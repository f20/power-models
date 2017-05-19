package Multiyear;

=head Copyright licence and disclaimer

Copyright 2014-2017 Franck Latrémolière, Reckon LLP and others.

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
use Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Shortcuts ':all';

sub table1001Overrides {
    my ( $me, $model, $wb, $ws, $rowName ) = @_;
    my $dataset = $me->{table1001Overrides}{ 0 + $model };
    return unless $dataset;
    my ($row) = grep { $rowName eq $dataset->{rows}{list}[$_]; }
      0 .. $#{ $dataset->{rows}{list} };
    return unless defined $row;
    my ( $wsheet, $ro, $co ) = $dataset->wsWrite( $wb, $ws );
    return unless $wsheet;
    q%'% . $wsheet->get_name . q%'!% . xl_rowcol_to_cell( $ro + $row, $co );
}

sub schedule15 {

    my ( $me, $t1001version ) = @_;

    my @t1001 = map {
            $_->{"table1001_$t1001version"}
          ? $_->{"table1001_$t1001version"}
          : undef;
    } @{ $me->{models} };

    my ($first1001) = grep { $_ } @t1001 or return;
    my $rowset = $first1001->{columns}[0]{rows};
    my $rowformats = $first1001->{columns}[3]{rowFormats};
    my $needNote1 = 1;
    Columnset(
        name => 'Allowed revenue summary (Schedule 15 table 1, '
          . $t1001version
          . ' version)',
        columns => [
            (
                map {
                    Stack(
                        name    => $_->{name},
                        rows    => $rowset,
                        sources => [
                            $_,
                            Constant(
                                rows => $rowset,
                                data => [ [ map { '' } @{ $rowset->{list} } ] ]
                            )
                        ],
                        defaultFormat => 'textnocolour',
                    );
                } @{ $first1001->{columns} }[ 0 .. 2 ]
            ),
            (
                map {
                    my $t1001 = $t1001[ $_ - 1 ];
                    $t1001
                      ? SpreadsheetModel::Custom->new(
                        name          => "Model $_",
                        rows          => $rowset,
                        custom        => [ '=A1', '=A2' ],
                        defaultFormat => 'millioncopy',
                        arguments     => {
                            A1 => $t1001->{columns}[3],
                            A2 => $t1001->{columns}[4],
                        },
                        model     => $me->{models}[ $_ - 1 ],
                        wsPrepare => sub {
                            my ( $self, $wb, $ws, $format, $formula,
                                $pha, $rowh, $colh )
                              = @_;
                            $self->{name} =
                              $me->modelIdentifier( $self->{model}, $wb, $ws );
                            if ($needNote1) {
                                undef $needNote1;
                                push @{ $self->{location}{postWriteCalls}{$wb}
                                  }, sub {
                                    my ( $me, $wbMe, $wsMe, $rowrefMe, $colMe )
                                      = @_;
                                    $wsMe->write_string(
                                        $$rowrefMe += 2,
                                        $colMe - 1,
                                        'Note 1: '
                                          . 'Cost categories associated '
                                          . 'with excluded services should '
                                          . 'only be populated '
                                          . 'if the Company recovers the '
                                          . 'costs of providing '
                                          . 'these services from '
                                          . 'Use of System Charges.',
                                        $wb->getFormat('text')
                                    );
                                  };
                            }
                            my $boldFormat = $wb->getFormat(
                                [
                                    base => 'millioncopy',
                                    bold => 1
                                ]
                            );

                            sub {
                                my ( $x, $y ) = @_;
                                local $_ = $rowformats->[$y];
                                $_ && /hard/
                                  ? (
                                    '',
                                    /(0\.0+)hard/
                                    ? $wb->getFormat( $1 . 'copy' )
                                    : $format,
                                    $formula->[0],
                                    qr/\bA1\b/ => xl_rowcol_to_cell(
                                        $rowh->{A1} + $y,
                                        $colh->{A1}, 1
                                    )
                                  )
                                  : (
                                    '',
                                    $boldFormat,
                                    $formula->[1],
                                    qr/\bA2\b/ => xl_rowcol_to_cell(
                                        $rowh->{A2} + $y,
                                        $colh->{A2}, 1
                                    )
                                  );
                            };
                        },
                      )
                      : ();
                } 1 .. @t1001,
            )
        ]
    );
}

1;
