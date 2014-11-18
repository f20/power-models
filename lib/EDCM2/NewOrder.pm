package EDCM2;

=head Copyright licence and disclaimer

Copyright 2014 Franck Latrémolière, Reckon LLP and others.

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

sub newOrder {

    my ( $model, @results ) = @_;

    my ( %calcTables, %dependencies );
    my $serialUplift = 0;
    my $addCalcTable;
    $addCalcTable = sub {
        my ( $ob, $destination ) = @_;
        $ob->{serial2} ||= $ob->{serial} + $serialUplift;
        return
             if !UNIVERSAL::isa( $ob, 'SpreadsheetModel::Dataset' )
          || $ob->{location}
          || !UNIVERSAL::isa( $ob, 'SpreadsheetModel::Constant' )
          && !$ob->{sourceLines};
        $calcTables{ 0 + ( $ob->{rows} || 0 ) }{ 0 + $ob } = $ob;
        undef $dependencies{ 0 + $destination }{ 0 + $ob }
          if $destination;
        $addCalcTable->( $_, $ob ) foreach @{ $ob->{sourceLines} };
    };
    foreach (@results) {
        $serialUplift += 10000;
        $addCalcTable->($_)
          foreach $_->{sourceLines} ? @{ $_->{sourceLines} } : ();
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
        my ($hashref) = @_;
        return unless %$hashref;
        my @columns = ();
        my $ncol    = 0;
        while (
            ( local $_ ) =
            sort { $hashref->{$a}{serial2} <=> $hashref->{$b}{serial2} }
            grep {
                !grep { $singlesRemaining{$_} || $tariffsRemaining{$_} }
                  keys %{ $getDeepDep->($_) };
            } keys %$hashref
          )
        {
            push @columns, delete $hashref->{$_};
        }
        @columns;
    };

    my @ordered;

    my $columnsetMaker = sub {
        my ( $prefix, @extras ) = @_;
        my $counter;
        sub {
            my @cols;
            my @result;
            foreach ( @_, { newBlock => 2 } ) {
                if ( $_->{newBlock}
                    && ( $model->{newOrder} < 2 || $_->{newBlock} == 2 ) )
                {
                    push
                      @result,  # NB: "#" has magical powers in a Columnset name
                      !@cols ? () : @cols == 1 ? @cols : Columnset(
                        name => "$prefix data #" . ++$counter,
                        columns =>
                          [ sort { $a->{serial} <=> $b->{serial} } @cols ],
                        @extras,
                      );
                    @cols = ();
                }
                push @cols, $_;
            }
            @result;
        };
    };

    push @{ $model->{generalTables} },
      $columnsetMaker->('Fixed parameter')->(
        sort   { $a->{serial} <=> $b->{serial} }
          grep { !$_->{cols} } @constantSingle,
        @constantOther
      ),
      grep { $_->{cols} } @constantSingle,
      @constantOther;

    my $singleMaker = $columnsetMaker->('Aggregate');
    my $tariffMaker = $columnsetMaker->('Tariff-specific');
    while ( %singlesRemaining || %tariffsRemaining ) {
        push @ordered, $tariffMaker->( $dataExtractor->( \%tariffsRemaining ) );
        push @ordered, $singleMaker->( $dataExtractor->( \%singlesRemaining ) );
    }

    $model->{newOrdering} = \@ordered;

}

sub makeCalcSheets {
    my $model   = shift;
    my $wbook   = shift;
    my $counter = 0;
    my @result;
    my @accum;
    foreach ( @{ $model->{newOrdering} }, undef ) {
        if (  !defined $_
            || ref $_ eq 'SpreadsheetModel::Columnset'
            && @{ $_->{columns} } > 8 )
        {
            my @a = @accum;
            @accum = $_;
            my $c = $counter++;
            push @a, @{ $model->{generalTables} } unless $c;
            push @result, $c ? "Calc$c" : 'General' => sub {
                my ($wsheet) = @_;
                $wsheet->{sheetNumber} = 40 + $c;
                $wsheet->{lastTableNumber} =
                  $model->{method} && $model->{method} =~ /LRIC/i ? 0 : -1;
                $wsheet->{tableNumberIncrement} = 2;
                $wsheet->freeze_panes( 1, 1 );
                $wsheet->set_column( 0, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes(
                    lines => $c ? "Calculation sheet $c" : 'General' ),
                  @a;
            };
        }
        else {
            push @accum, $_;
        }
    }
    @result;
}

1;
