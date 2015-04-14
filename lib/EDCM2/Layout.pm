package EDCM2;

=head Copyright licence and disclaimer

Copyright 2014-2015 Franck Latrémolière, Reckon LLP and others.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;

sub orderedLayout {

    my ( $model, @finalCalcTableList ) = @_;
    my $unnumbered = $model->{layout} =~ /unnumbered/i;
    my $wideLayout = $model->{layout} =~ /wide/i;
    @finalCalcTableList = [ map { @$_ } @finalCalcTableList ] if $wideLayout;

    my ( %calcTables, %dependencies, $addCalcTable );
    my $serialUplift = 0;
    $addCalcTable = sub {
        my ( $ob, $destination ) = @_;
        return
             if !UNIVERSAL::isa( $ob, 'SpreadsheetModel::Dataset' )
          || ref $ob->{location}
          || !UNIVERSAL::isa( $ob, 'SpreadsheetModel::Constant' )
          && !$ob->{sourceLines};
        $calcTables{ 0 + ( $ob->{rows} || 0 ) }{ 0 + $ob } = $ob;
        undef $dependencies{ 0 + $destination }{ 0 + $ob }
          if $destination;
        $addCalcTable->( $_, $ob ) foreach @{ $ob->{sourceLines} };
        $ob->{serialUplifted} ||= $serialUplift + $ob->{serial};
    };

    foreach ( grep { $_ } @finalCalcTableList ) {
        $serialUplift += 95_000;
        $addCalcTable->($_) foreach @$_;
    }

    my ( %deepDep, $getDeepDep );
    $getDeepDep = sub {
        my ($dst) = @_;
        my $dep = $dependencies{$dst} || {};
        $deepDep{$dst} ||=
          { %$dep, map { %{ $getDeepDep->($_) } } keys %$dep };
    };

    my ( %singlesRemaining, %tariffsRemaining, @constantSingle,
        @constantOther );
    while ( my ( $rows, $tset ) = each %calcTables ) {
        if ( !$rows ) {
            foreach ( values %$tset ) {
                if ( ref $_ eq 'SpreadsheetModel::Constant' ) {
                    push @constantSingle, $_;
                }
                else {
                    $singlesRemaining{ 0 + $_ } = $_;
                }
            }
        }
        elsif ( !%tariffsRemaining && values %$tset > 5 ) {
            %tariffsRemaining = %$tset;
        }
        else {
            push @constantOther,
              grep { ref $_ eq 'SpreadsheetModel::Constant'; } values %$tset;
        }
    }

    my $dataExtractor = sub {
        my ( $hashref, $maxSerial ) = @_;
        return unless %$hashref;
        my @columns = ();
        while (
            ( local $_ ) =
            sort {
                $hashref->{$a}{serialUplifted}
                  <=> $hashref->{$b}{serialUplifted}
            }
            grep {
                $hashref->{$_}{serialUplifted} < $maxSerial
                  and !grep { $singlesRemaining{$_} || $tariffsRemaining{$_} }
                  keys %{ $getDeepDep->($_) };
            } keys %$hashref
          )
        {
            push @columns, delete $hashref->{$_};
        }
        @columns;
    };

    my $groupMaker = sub {
        my ( $prefix, @extras ) = @_;
        my $grouper;
        if ( $model->{layout} =~ /matrix/i ) {
            return sub { @_; }
              unless $prefix =~ /Tariff-specific/i;
            my $matrix = new SpreadsheetModel::MatrixSheet(
                $model->{tariff1Row}
                ? (
                    dataRow            => $model->{tariff1Row},
                    captionDecorations => [qw(algae purple slime)],
                  )
                : (),
            );
            push @{ $model->{matrixTables} },
              my $copy = Stack( sources => [ $model->{table935}{columns}[0] ] );
            $matrix->addDatasetGroup(
                name    => 'Tariff name',
                columns => [$copy]
            );
            $grouper = sub {
                return unless @_;
                my $name = ' ';
                $name = $_ foreach grep { $_ } map { $_->{groupName} } @_;
                $matrix->addDatasetGroup(
                    name    => $name,
                    columns => [@_],
                );
                @_;
            };
        }
        elsif ($unnumbered) {
            $grouper = sub {
                return unless @_;
                Columnset(
                    name       => '',
                    logColumns => 1,
                    columns    => [@_],
                    @extras,
                );
            };
        }
        else {
            my $counter = 0;
            $grouper = sub {
                return unless @_;
                return @_ if @_ == 1 && !$_[0]{rows};
                my $name;
                $name = $_ foreach grep { $_ } map { $_->{groupName} } @_;
                Columnset(
                    name => "$prefix #"
                      . ++$counter
                      . ( $name ? ": $name" : '' ),
                    logColumns => 1,
                    columns    => [@_],
                    @extras,
                );
            };
        }
        sub {
            my @cols;
            my @result;
            foreach (
                ( sort { $a->{serialUplifted} <=> $b->{serialUplifted} } @_ ),
                { theEnd => 1 } )
            {
                if (    @cols
                    and $_->{theEnd} || $_->{newBlock} && !$wideLayout )
                {
                    push @result, $grouper->(@cols);
                    @cols = ();
                }
                push @cols, $_;
            }
            @result;
        };
    };

    push @{ $model->{generalTables} },
      $groupMaker->('Fixed parameters')->(
        sort { $a->{serialUplifted} <=> $b->{serialUplifted} }
        grep { !$_->{cols} } @constantSingle
      ),
      sort { $a->{serialUplifted} <=> $b->{serialUplifted} }
      ( @constantOther, grep { $_->{cols} } @constantSingle );

    my @ordered;
    my $singleMaker = $groupMaker->('Aggregate data');
    my $tariffMaker = $groupMaker->('Tariff-specific data');
    $serialUplift += 95_000;
    for (
        my $maxSerial = 50_000 ;
        $maxSerial < $serialUplift ;
        $maxSerial += 95_000
      )
    {
        while (
            (
                grep { $_->{serialUplifted} < $maxSerial }
                values %singlesRemaining
            )
            || ( grep { $_->{serialUplifted} < $maxSerial }
                values %tariffsRemaining )
          )
        {
            push @ordered,
              $tariffMaker->(
                $dataExtractor->( \%tariffsRemaining, $maxSerial ) );
            push @ordered,
              $singleMaker->(
                $dataExtractor->( \%singlesRemaining, $maxSerial ) );
        }
    }

    \@ordered;

}

1;
