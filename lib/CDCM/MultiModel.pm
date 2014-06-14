package CDCM::MultiModel;

=head Copyright licence and disclaimer

Copyright 2014 Franck Latrémolière.

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

sub new {
    bless {
        historical       => [],
        scenario         => [],
        statsAssumptions => [],
        statsSections    => [ split /\n/, <<EOL ],
Annual charges for illustrative customers (£/year)
Distribution costs for illustrative customers (£/MWh)
DNO-wide aggregates
EOL
      },
      shift;
}

sub worksheetsAndClosuresMulti {

    my ( $me, $model, $wbook, @pairs ) = @_;

    push @{ $me->{finishClosures} }, sub {
        delete $wbook->{logger};
        delete $wbook->{titleAppend};
        delete $wbook->{noLinks};
    };

    unless ( @{ $me->{historical} } || @{ $me->{scenario} } ) {

        push @pairs,

          'Index$' => sub {
            my ($wsheet) = @_;
            push @{ $me->{finishClosures} }, sub {
                my $noLinks = delete $wbook->{noLinks};
                $wsheet->set_column( 0, 0,   70 );
                $wsheet->set_column( 1, 255, 14 );
                $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                    name  => 'Annual Review Pack',
                    lines => [
                             $model->{colour}
                          && $model->{colour} =~ /orange|gold/ ? <<EOL : (),

This document, model or dataset has been prepared by Reckon LLP on the instructions of the DCUSA Panel or one of its working
groups.  Only the DCUSA Panel and its working groups have authority to approve this material as meeting their requirements. 
Reckon LLP makes no representation about the suitability of this material for the purposes of complying with any licence
conditions or furthering any relevant objective.
EOL

                        <<'EOL',

UNLESS STATED OTHERWISE, THIS WORKBOOK IS ONLY A PROTOTYPE FOR TESTING PURPOSES AND ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
                        <<'EOL',

Copyright 2009-2011 Energy Networks Association Limited and others. Copyright 2011-2014 Franck Latrémolière, Reckon LLP and others. 
The code used to generate this spreadsheet includes open-source software published at https://github.com/f20/power-models.
Use and distribution of the source code is subject to the conditions stated therein. 
Any redistribution of this software must retain the following disclaimer:
THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL AUTHORS OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED
AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
EOL

                    ]
                );
                $_->wsWrite( $wbook, $wsheet ) foreach

                  @{ $me->{historical} }
                  ? Notes(
                    name        => 'Historical models',
                    sourceLines => [
                        map {
                            [ $_->{nickName}, undef, @{ $_->{sheetLinks} }, ];
                        } @{ $me->{historical} }
                    ],
                  )
                  : (),

                  @{ $me->{scenario} }
                  ? Notes(
                    name        => 'Scenario models',
                    sourceLines => [
                        map {
                            [
                                $me->{scenario}[$_]{nickName},
                                $me->{assumptionColumns}[$_],
                                @{ $me->{scenario}[$_]{sheetLinks} },
                            ];
                        } 0 .. $#{ $me->{scenario} }
                    ],
                  )
                  : ();

                $wbook->{noLinks} = $noLinks if defined $noLinks;
            };
          },

          'Schedule 15$' => sub {
            my ($wsheet) = @_;
            $wsheet->set_column( 0, 0,   60 );
            $wsheet->set_column( 1, 254, 16 );
            $wsheet->freeze_panes( 0, 1 );
            push @{ $me->{finishClosures} }, sub {
                Notes(
                    name => 'Analysis of allowed revenue (DCUSA schedule 15)' )
                  ->wsWrite( $wbook, $wsheet );
                my @t1001 = map {
                         $_->{table1001}
                      && $_->{targetRevenue} !~ /DCP132longlabels/i
                      ? $_->{table1001}
                      : undef
                } @{ $me->{models} };
                my ($first1001) = grep { $_ } @t1001 or return;
                my $rowset     = $first1001->{columns}[0]{rows};
                my $rowformats = $first1001->{columns}[3]{rowFormats};
                $wbook->{noLinks} = 1;
                my $needNote1 = 1;
                Columnset(
                    name    => 'Schedule 15 table 1',
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
                                            data => [
                                                [
                                                    map { '' }
                                                      @{ $rowset->{list} }
                                                ]
                                            ]
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
                                    custom        => [ '=IV1', '=IV2' ],
                                    defaultFormat => 'millioncopy',
                                    arguments     => {
                                        IV1 => $t1001->{columns}[3],
                                        IV2 => $t1001->{columns}[4],
                                    },
                                    model     => $me->{models}[ $_ - 1 ],
                                    wsPrepare => sub {
                                        my ( $self, $wb, $ws, $format, $formula,
                                            $pha, $rowh, $colh )
                                          = @_;
                                        $self->{name} = $self->{model}
                                          ->modelIdentification( $wb, $ws );
                                        if ($needNote1) {
                                            undef $needNote1;
                                            push @{ $self->{location}
                                                  {postWriteCalls}{$wb} }, sub {
                                                $ws->write_string(
                                                    $ws->{nextFree}++,
                                                    0,
                                                    'Note 1: '
                                                      . 'Cost categories associated '
                                                      . 'with excluded services should only be populated '
                                                      . 'if the Company recovers the costs of providing '
                                                      . 'these services from Use of System Charges.',
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
                                                qr/\bIV1\b/ =>
                                                  xl_rowcol_to_cell(
                                                    $rowh->{IV1} + $y,
                                                    $colh->{IV1},
                                                    1
                                                  )
                                              )
                                              : (
                                                '',
                                                $boldFormat,
                                                $formula->[1],
                                                qr/\bIV2\b/ =>
                                                  xl_rowcol_to_cell(
                                                    $rowh->{IV2} + $y,
                                                    $colh->{IV2},
                                                    1
                                                  )
                                              );
                                        };
                                    },
                                  )
                                  : ();
                            } 1 .. @t1001,
                        )
                    ]
                )->wsWrite( $wbook, $wsheet );
            };
          };

        unshift @pairs,

          'Statistics$' => sub {
            my ($wsheet) = @_;
            $wsheet->set_column( 0, 255, 50 );
            $wsheet->set_column( 1, 255, 16 );
            $wsheet->freeze_panes( 0, 1 );
            my $noData = delete $wbook->{noData};
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name  => 'Statistical time series',
                lines => [
                        'The items below are only examples of some of'
                      . ' the statistical outputs that could'
                      . ' be included in this sheet.',
                ],
              ),
              @{ $me->{statsAssumptions} };
            $wbook->{noData} = $noData if defined $noData;
            push @{ $me->{finishClosures} }, sub {
                $wbook->{noLinks} = 1;
                $_->wsWrite( $wbook, $wsheet )
                  foreach $me->statisticsColumnsets( $wbook, $wsheet );
            };
          };

    }

    push @{ $me->{models} }, $model;

    if ( ref $model->{dataset} eq 'HASH' && !$model->{dataset}{baseDataset} ) {
        push @{ $me->{historical} }, $model;
        return @pairs;
    }

    unless ( $me->{assumptionColumns} ) {
        $me->{assumptionColumns} = [];
        $me->{assumptionRowset}  = Labelset(
            list => [
                'Change in the price control index (RPI)',              #  0
                'MEAV change: 132kV',                                   #  1
                'MEAV change: 132kV/EHV',                               #  2
                'MEAV change: EHV',                                     #  3
                'MEAV change: EHV/HV',                                  #  4
                'MEAV change: 132kV/HV',                                #  5
                'MEAV change: HV network',                              #  6
                'MEAV change: HV service',                              #  7
                'MEAV change: HV/LV',                                   #  8
                'MEAV change: LV network',                              #  9
                'MEAV change: LV service',                              # 10
                'Cost change: direct costs',                            # 11
                'Cost change: indirect costs',                          # 12
                'Cost change: network rates',                           # 13
                'Cost change: transmission exit',                       # 14
                'Volume change: supercustomer metered demand units',    # 15
                'Volume change: supercustomer metered demand MPANs',    # 16
                'Volume change: site-specific metered demand units',    # 17
                'Volume change: site-specific metered demand MPANs',    # 18
                'Volume change: demand capacity',                       # 19
                'Volume change: demand excess reactive',                # 20
                'Volume change: unmetered demand units',                # 21
                'Volume change: generation units',                      # 22
                'Volume change: generation MPANs',                      # 23
                'Volume change: generation excess reactive',            # 24
            ]
        );
        unshift @pairs, 'Assumptions$' => sub {
            my ($wsheet) = @_;
            $wsheet->set_column( 0, 255, 50 );
            $wsheet->set_column( 1, 255, 20 );
            $wsheet->freeze_panes( 0, 1 );
            my $logger      = delete $wbook->{logger};
            my $titleAppend = delete $wbook->{titleAppend};
            my $noLinks     = $wbook->{noLinks};
            $wbook->{noLinks} = 1;
            $_->wsWrite( $wbook, $wsheet )
              foreach $me->{assumptions} = Notes( name => 'Assumptions' );
            my $headerRowForLater = ++$wsheet->{nextFree};
            ++$wsheet->{nextFree};
            push @{ $me->{finishClosures} }, sub {

                for ( my $i = 0 ; $i < @{ $me->{assumptionColumns} } ; ++$i ) {
                    my $model = $me->{assumptionColumns}[$i]{model};
                    $wsheet->write(
                        $headerRowForLater + 1,
                        $i + 1,
                        $model->modelIdentification( $wbook, $wsheet ),
                        $wbook->getFormat('thc')
                    );
                    $wsheet->write(
                        $headerRowForLater,
                        $i + 1,
                        $model->{sourceModel}
                          ->modelIdentification( $wbook, $wsheet ),
                        $wbook->getFormat('thc')
                    );
                }
            };
            $_->wsWrite( $wbook, $wsheet ) foreach Columnset(
                name      => '',
                noHeaders => 1,
                columns   => $me->{assumptionColumns},
            );
            $wbook->{logger}      = $logger;
            $wbook->{titleAppend} = $titleAppend;
            $wbook->{noLinks}     = $noLinks;
        };
    }

    push @{ $me->{scenario} }, $model;

    push @{ $me->{assumptionColumns} },
      $me->{assumptionsByModel}{ 0 + $model } = Constant(
        name          => 'Assumptions',
        model         => $model,
        rows          => $me->{assumptionRowset},
        defaultFormat => '%hardpm',
        data          => [
            [
                qw(0.035
                  0.02 0.02 0.02 0.02 0.02 0.02 0.02 0.02 0.02 0.02
                  0.02 0.02 0.02 0.02
                  -0.01 0
                  0.01 0.01 0.01 0.01
                  0
                  0.03 0.03 0.03)
            ]
        ],
      );

    @pairs;

}

sub assumptionsLocator {
    my ( $me, $model, $sourceModel ) = @_;
    my @assumptionsColumnLocationArray;
    sub {
        my ( $wb, $ws, $row ) = @_;
        unless ( $row =~ /^[0-9]+$/s ) {
            my $q = qr/$row/;
            ($row) = grep { $me->{assumptionRowset}{list}[$_] =~ /$q/; }
              0 .. $#{ $me->{assumptionRowset}{list} };
        }
        unless (@assumptionsColumnLocationArray) {
            @assumptionsColumnLocationArray =
              $me->{assumptionsByModel}{ 0 + $model }->wsWrite( $wb, $ws );
            $assumptionsColumnLocationArray[0] =
              q%'% . $assumptionsColumnLocationArray[0]->get_name . q%'!%;
        }
        $assumptionsColumnLocationArray[0]
          . xl_rowcol_to_cell(
            $assumptionsColumnLocationArray[1] + $row,
            $assumptionsColumnLocationArray[2],
            1, 1
          );
    };
}

sub statisticsColumnsets {
    my ( $me, $wbook, $wsheet ) = @_;
    map {
        if ( $me->{statsRows}[$_] ) {
            my $rows = Labelset( list => $me->{statsRows}[$_] );
            my $statsMaps = $me->{statsMap}[$_];
            $rows->{groups} = 'fake';
            for ( my $r = 0 ; $r < @{ $rows->{list} } ; ++$r ) {
                $rows->{groupid}[$r] = 'fake'
                  if grep { $_->[$r] } values %$statsMaps;
            }
            my @columns =
              map {
                if ( my $relevantMap = $statsMaps->{ 0 + $_ } ) {
                    SpreadsheetModel::Custom->new(
                        name => $_->modelIdentification( $wbook, $wsheet ),
                        rows => $rows,
                        custom    => [ map { "=IV1$_"; } 0 .. $#$relevantMap ],
                        arguments => {
                            map {
                                my $t;
                                $t = $relevantMap->[$_][0]
                                  if $relevantMap->[$_];
                                $t ? ( "IV1$_" => $t ) : ();
                            } 0 .. $#$relevantMap
                        },
                        defaultFormat => '0.000copy',
                        rowFormats    => [
                            map {
                                if ( $_ && $_->[0] ) {
                                    local $_ = $_->[0]{rowFormats}[ $_->[2] ]
                                      || $_->[0]{defaultFormat};
                                    s/(?:soft|hard|con)/copy/ if $_ && !ref $_;
                                    $_;
                                }
                                else { 'unavailable'; }
                            } @$relevantMap
                        ],
                        wsPrepare => sub {
                            my ( $self, $wb, $ws, $format, $formula,
                                $pha, $rowh, $colh )
                              = @_;
                            sub {
                                my ( $x, $y ) = @_;
                                my $cellFormat =
                                    $self->{rowFormats}[$y]
                                  ? $wb->getFormat( $self->{rowFormats}[$y] )
                                  : $format;
                                return '', $cellFormat
                                  unless $relevantMap->[$y];
                                my ( $table, $offx, $offy ) =
                                  @{ $relevantMap->[$y] };
                                my $ph = "IV1$y";
                                '', $cellFormat, $formula->[$y], $ph,
                                  xl_rowcol_to_cell(
                                    $rowh->{$ph} + $offy,
                                    $colh->{$ph} + $offx,
                                    1, 1,
                                  );
                            };
                        },
                    );
                }
                else {
                    ();
                }
              } @{ $me->{models} };
            Columnset(
                name    => $me->{statsSections}[$_],
                columns => \@columns,
            );
        }
        else { (); }
    } 0 .. $#{ $me->{statsSections} };
}

sub addStats {
    my $me      = shift;
    my $section = shift;
    my $model;
    if ( ref $section eq 'CDCM' ) {
        $model   = $section;
        $section = 'General aggregates';
    }
    else {
        $model = shift;
    }
    my ($sectionNumber) = grep { $section eq $me->{statsSections}[$_]; }
      0 .. $#{ $me->{statsSections} };
    unless ( defined $sectionNumber ) {
        push @{ $me->{statsSections} }, $section;
        $sectionNumber = $#{ $me->{statsSections} };
    }
    foreach my $table (@_) {
        if ( my $lastRow = $table->lastRow ) {
            for ( my $row = 0 ; $row <= $lastRow ; ++$row ) {
                my $name      = "$table->{rows}{list}[$row]";
                my $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name};
                unless ( defined $rowNumber ) {
                    push @{ $me->{statsRows}[$sectionNumber] }, $name;
                    $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name} =
                      $#{ $me->{statsRows}[$sectionNumber] };
                }
                $me->{statsMap}[$sectionNumber]{ 0 + $model }[$rowNumber] =
                  [ $table, 0, $row ]
                  unless $table->{rows}{groupid}
                  && !defined $table->{rows}{groupid}[$row];
            }
        }
        else {
            my $name =
              UNIVERSAL::can( $table->{name}, 'shortName' )
              ? $table->{name}->shortName
              : "$table->{name}";
            my $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name};
            unless ( defined $rowNumber ) {
                push @{ $me->{statsRows}[$sectionNumber] }, $name;
                $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name} =
                  $#{ $me->{statsRows}[$sectionNumber] };
            }
            $me->{statsMap}[$sectionNumber]{ 0 + $model }[$rowNumber] =
              [ $table, 0, 0 ];
        }
    }
}

sub finish {
    $_->() foreach @{ $_[0]{finishClosures} };
}

1;
