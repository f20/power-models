package PowerModels::Extract::SelectedTablesJson;

=head Copyright licence and disclaimer

Copyright 2017 Franck Latrémolière, Reckon LLP and others.

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
use Encode;

use constant { OT_IGNORE_FILTER => 0, };

sub new {
    my ( $class, $ignoreFilter ) = @_;
    if ( $ignoreFilter && !ref $ignoreFilter ) {
        my @tableStems = $ignoreFilter =~ /([0-9]+)/g;
        my $tableRe = '^(?:' . join( '|', @tableStems ) . ')';
        if ( grep { $_ == 0; } @tableStems ) {
            $ignoreFilter = sub {
                $_[1] =~ /$tableRe/s ? undef : 0;
            };
        }
        else {
            my %sheetStems;
            foreach (@tableStems) {
                undef $sheetStems{$1} if /^([0-9]{2})/;
            }
            my $sheetRe = '^(?:' . join( '|', keys %sheetStems ) . ')';
            $ignoreFilter = sub {
                $_[1] !~ /$sheetRe/s ? 1 : $_[1] !~ /$tableRe/s ? 0 : undef;
            };
        }
    }
    bless [$ignoreFilter], $class;
}

sub writerAndParserOptions {
    my ($self) = @_;
    my %data;
    my ( @tableNumber, @rowZero, @notRowZero );
    sub {
        my ( $workbookFile, $workbookParseResult ) = @_;
        my $jsonMachine;
        foreach (qw(JSON JSON::PP)) {
            last if eval "require $_" and $jsonMachine = $_->new;
        }
        die 'No JSON module' unless $jsonMachine;
        $jsonMachine->canonical(1)->utf8;
        $workbookFile =~ s/\.xlsx?$//s;
        open my $h, '>', "$workbookFile.json$$";
        binmode $h;
        print {$h} $jsonMachine->encode( \%data );
        close $h;
        rename "$workbookFile.json$$", "$workbookFile.json";
      },
      Setup => sub { %data = (); },
      NotSetCell  => 1,
      CellHandler => sub {
        my ( $wbook, $sheetIdx, $row, $col, $cell ) = @_;
        my $v = $cell->unformatted;
        return unless defined $v;
        eval { $v = Encode::decode( 'UTF-16BE', $v ); }
          if $v =~ m/\x{0}/;
        if ( !$col ) {
            if ( $v =~ /^([0-9]+)\. /s ) {
                my $tn = $1;
                my $ignoreTable;
                $ignoreTable = $self->[OT_IGNORE_FILTER]->( $tn, $v )
                  if $self->[OT_IGNORE_FILTER];
                return $ignoreTable if defined $ignoreTable;
                $data{ $tableNumber[$sheetIdx] = $tn }[0][0] = $v;
                undef $rowZero[$sheetIdx];
                $notRowZero[$sheetIdx] = $row;
            }
            elsif ( $tableNumber[$sheetIdx] && $v ne '' ) {
                if ( defined $rowZero[$sheetIdx] ) {
                    if (
                        defined $data{ $tableNumber[$sheetIdx] }[0]
                        [ $row - $rowZero[$sheetIdx] - 1 ] )
                    {
                        $data{ $tableNumber[$sheetIdx] }[0]
                          [ $row - $rowZero[$sheetIdx] ] = $v;
                    }
                    else {
                        undef $tableNumber[$sheetIdx];
                        undef $rowZero[$sheetIdx];
                    }
                }
                else {
                    $notRowZero[$sheetIdx] = $row;
                }
            }
        }
        elsif ( $tableNumber[$sheetIdx] ) {
            if ( defined $rowZero[$sheetIdx] ) {
                $data{ $tableNumber[$sheetIdx] }[$col]
                  [ $row - $rowZero[$sheetIdx] ] = $v;
            }
            elsif ( $row > $notRowZero[$sheetIdx] ) {
                $rowZero[$sheetIdx] = $row;
                $data{ $tableNumber[$sheetIdx] }[$col][0] =
                  $v;
            }
        }
        0;
      };
}

1;

