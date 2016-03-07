package SpreadsheetModel::ColumnsetFilter;

=head Copyright licence and disclaimer

Copyright 2016 Franck Latrémolière, Reckon LLP and others.

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

require SpreadsheetModel::Columnset;
our @ISA = qw(SpreadsheetModel::Columnset);

use Spreadsheet::WriteExcel::Utility;

sub wsWrite {

    my ( $self, $wb, $ws, ) = @_;

    my ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 );
    if ( $self->{name} ) {
        $ws->set_row( $row, 21 );
        $ws->write( $row++, $col, "$self->{name}", $wb->getFormat('caption') );
        ++$row;
    }

    ++$row;    # for column headers

    my $thcFormat = $wb->getFormat( [ base => 'thc', locked => 0 ] );
    my $lastRow = $self->{rows} ? $#{ $self->{rows}{list} } : 0;
    my $lastCol = 0;

    my $dataset;
    $dataset = $self->{dataset}{'!'} if $self->{dataset};
    if ( $dataset && ref $dataset->[1] eq 'HASH' ) {
        $dataset->[$_] = [ @{ $dataset->[$_] }{ @{ $dataset->[0] } } ]
          foreach grep { ref $dataset->[$_] eq 'HASH'; } 1 .. $#$dataset;
        $_ = '' foreach grep { /^Anon-[0-9]+$/ } @{ $dataset->[0] };
    }

    $self->{$wb}{$ws} = 1;

    foreach ( @{ $self->{columns} } ) {

        @{ $_->{$wb} }{qw(worksheet row col)} = ( $ws, $row, $col + $lastCol );

        my $lCol = $_->lastCol;
        foreach my $c ( 0 .. $lCol ) {
            $ws->write(
                $row - 1,
                $col + $lastCol + $c,
                $_->objectShortName
                  . (
                    $lCol
                    ? (
                        ': '
                          . SpreadsheetModel::Object::_shortName(
                            $_->{cols}{list}[$c]
                          )
                      )
                    : ''
                  ),
                $thcFormat
            );
        }

        if ( ref $_ eq 'SpreadsheetModel::Dataset' ) {
            my $format = $wb->getFormat( $_->{defaultFormat} || 'texthard' );
            for my $c ( 0 .. $lCol ) {
                foreach my $r ( 0 .. $lastRow ) {
                    my $value =
                        $dataset
                      ? $dataset->[ $lastCol + $c ][$r]
                      : $_->{defaultData};
                    $value = "=$value"
                      if $value
                      and $value eq '#VALUE!' || $value eq '#N/A'
                      and $wb->formulaHashValues;
                    $ws->write( $row + $r, $col + $lastCol + $c,
                        $value, $format );
                }
            }
        }
        else {
            my $dobj = $_;
            my $drow = $row;
            my $dcol = $col + $lastCol;
            my $dws  = $ws;
            $wb->{deferralMaster}->(
                sub {
                    my $cell = $dobj->wsPrepare( $wb, $dws );
                    foreach my $c ( 0 .. $dobj->lastCol ) {
                        foreach my $r ( 0 .. $dobj->lastRow ) {
                            my ( $value, $format, $formula, @more ) =
                              $cell->( $c, $r );
                            if (@more) {
                                $dws->repeat_formula(
                                    $drow + $r, $dcol + $c, $formula,
                                    $format,    @more
                                );
                            }
                            elsif ($formula) {
                                $dws->write_formula(
                                    $drow + $r, $dcol + $c, $formula,
                                    $format,    $value
                                );
                            }
                            else {
                                $value = "=$value"
                                  if $value
                                  and $value eq '#VALUE!' || $value eq '#N/A'
                                  and $wb->formulaHashValues;
                                $dws->write( $drow + $r, $dcol + $c,
                                    $value, $format );
                            }
                        }
                    }
                }
            );
        }

        $_->dataValidation( $wb, $ws, $row, $col + $lastCol, $row + $lastRow )
          if $_->{validation};

        if ( $_->{validationDeferred} ) {
            my $dobj = $_;
            my $drow = $row;
            my $dcol = $col + $lastCol;
            my $dlst = $row + $lastRow;
            my $dws  = $ws;
            $wb->{deferralMaster}->(
                sub {
                    $dobj->{validation} =
                      $dobj->{validationDeferred}->( $wb, $dws );
                    $dobj->dataValidation( $wb, $dws, $drow, $dcol, $dlst );
                }
            );
        }

        if ( my $cf = $_->{conditionalFormatting} ) {
            foreach (
                ref $cf eq 'ARRAY'
                ? @$cf
                : $cf
              )
            {
                $_->{format} = $wb->getFormat( $_->{format} )
                  if $_->{format} && ( ref $_->{format} ) !~ /ormat/;
                eval {
                    $ws->conditional_formatting(
                        $row,
                        $col + $lastCol,
                        $row + $lastRow,
                        $col + $lastCol + $lCol, $_
                    );
                };
                if ($@) {
                    warn "Omitting conditional formatting: $@";
                    return;
                }
            }
        }

        $lastCol += $lCol + 1;

    }

    $ws->autofilter( $row - 1, $col, $row + $lastRow, $col + $lastCol - 1 );
    $ws->{protectionOptions} ||= {
        autofilter            => 1,
        select_locked_cells   => 0,
        select_unlocked_cells => 1,
        sort                  => 1,
    };

    $row += $lastRow;
    $self->requestForwardLinks( $wb, $ws, \$row, $col )
      if $wb->{forwardLinks};
    ++$row;
    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;
    if ( $self->{postWriteCalls}{$wb} ) {
        $_->($self) foreach @{ $self->{postWriteCalls}{$wb} };
    }

}

1;
