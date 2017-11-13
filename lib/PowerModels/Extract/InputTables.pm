package PowerModels::Extract::InputTables;

=head Copyright licence and disclaimer

Copyright 2008-2017 Franck Latrémolière, Reckon LLP and others.

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

sub extractInputData {
    my ( $workbook, $tree, $options ) = @_;
    my ( %byWorksheet, %used );
    my $conflictStatus = 0;
    for my $worksheet ( $workbook->worksheets() ) {
        my ( $row_min, $row_max ) = $worksheet->row_range();
        my ( $col_min, $col_max ) = $worksheet->col_range();
        my ( $tableNumber, $evenIfLocked, $columnHeadingsRow, $to1, $to2 );
        for my $row ( $row_min .. $row_max ) {
            my $rowName;
            for my $col ( $col_min .. $col_max ) {
                my $cell = $worksheet->get_cell( $row, $col );
                my $v;
                $v = $cell->unformatted if $cell;
                $evenIfLocked = 1
                  if $col == 0
                  && !$v
                  && defined $evenIfLocked
                  && !$evenIfLocked;
                next unless defined $v;
                eval { $v = Encode::decode( 'UTF-16BE', $v ); }
                  if $v =~ m/\x{0}/;
                if ( $col == 0 ) {
                    if ( !ref $cell->{Format} || $cell->{Format}{Lock} ) {
                        if ( $v =~ /^[0-9]{3,}\. .*⇒([0-9]{3,})/
                            && !( $evenIfLocked = 0 )
                            || $v =~ /^([0-9]{3,})\. /
                            && !( undef $evenIfLocked ) )
                        {
                            $tableNumber = $1;
                            undef $columnHeadingsRow;
                            $to2 = [];
                            $conflictStatus = defined $evenIfLocked ? 1 : 2
                              if $used{$tableNumber} && $conflictStatus < 2;
                            $to1 =
                              $used{$tableNumber}
                              ? [ map { +{%$_}; } @{ $tree->{$tableNumber} } ]
                              : $to2;
                            $used{$tableNumber} = 1;
                            $to1->[0]{_table} = $to2->[0]{_table} = $v
                              unless $options->{minimum};
                        }
                        elsif ($v) {
                            $v =~ s/[^A-Za-z0-9-]/ /g;
                            $v =~ s/- / /g;
                            $v =~ s/ +/ /g;
                            $v =~ s/^ //;
                            $v =~ s/ $//;
                            $rowName =
                              $v eq ''
                              ? 'Anon' . ( ( $columnHeadingsRow || 0 ) - $row )
                              : $v;
                        }
                        else {
                            undef $tableNumber;
                        }
                    }
                    elsif ( $worksheet->{Name} !~ /^(?:Index|Overview)$/s )
                    {    # unlocked cell in column 0
                        if ( defined $tableNumber ) {
                            next unless defined $columnHeadingsRow;
                            $rowName = $row - $columnHeadingsRow;
                            $to1->[0]{$rowName} = $to2->[0]{$rowName} = $v;
                        }
                        else {
                            $tableNumber       = '!';
                            $columnHeadingsRow = $row;
                        }
                    }
                }
                elsif ( defined $tableNumber ) {
                    if ( !defined $rowName ) {
                        $columnHeadingsRow = $row;
                        if ( $options->{minimum} ) {
                            $to1->[$col] ||= {};
                        }
                        else {
                            $to1->[$col]{'_column'} = $v;
                            $to2->[$col]{'_column'} = $v;
                        }
                    }
                    elsif ( $evenIfLocked
                        || ref $cell->{Format} && !$cell->{Format}{Lock}
                        and $v
                        || $to1->[$col] )
                    {
                        $to1->[$col]{$rowName} = $to2->[$col]{$rowName} = $v;
                        $tree->{$tableNumber} ||= $to2;
                        $byWorksheet{' combined'}{$tableNumber} = $to1;
                        $byWorksheet{" $worksheet->{Name}"}{$tableNumber} =
                          $to2;
                    }
                }
            }
        }
    }
    '', $tree,
       !$conflictStatus ? ()
      : $conflictStatus == 1 ? ( ' combined' => $byWorksheet{' combined'} )
      :                        %byWorksheet;
}

1;

