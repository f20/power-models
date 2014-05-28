package CDCM::ARP;

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
        historical => [],
        scenario   => [],
        allmodels  => 0,
      },
      shift;
}

sub worksheetsAndClosuresWithArp {

    my ( $arp, $model, $wbook, @pairs ) = @_;

    push @{ $arp->{finishClosures} }, sub {
        delete $wbook->{logger};
        delete $wbook->{titleAppend};
        delete $wbook->{noLinks};
    };

    push @pairs, 'Index$' => sub {
        my ($wsheet) = @_;
        push @{ $arp->{finishClosures} }, sub {
            $wsheet->set_column( 0, 0,   70 );
            $wsheet->set_column( 1, 255, 14 );
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name  => 'Annual Review Pack',
                lines => [
                         $model->{colour}
                      && $model->{colour} =~ /orange|gold/ ? <<EOL : (),

This document, model or dataset has been prepared by Reckon LLP on the instructions of the DCUSA Panel or
one of its working groups.  Only the DCUSA Panel and its working groups have authority to approve this
material as meeting their requirements.  Reckon LLP makes no representation about the suitability of this
material for the purposes of complying with any licence conditions or furthering any relevant objective.
EOL
                ]
              ),
              $model->licenceNotes;
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name        => 'Historical models',
                sourceLines => [
                    map { [ $_->{nickName}, undef, @{ $_->{sheetLinks} }, ]; }
                      @{ $arp->{historical} }
                ],
              ),
              Notes(
                name        => 'Scenario models',
                sourceLines => [
                    map {
                        [
                            $_->{nickName}, $arp->{assumptions},
                            @{ $_->{sheetLinks} },
                        ];
                    } @{ $arp->{scenario} }
                ],
              );
        };
      },

      1 ? () : (

        'Timebands$' => sub {
            my ($wsheet) = @_;
            push @{ $arp->{finishClosures} }, sub {
                $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                    name  => 'Specification of distribution time bands',
                    lines => 'This sheet has not yet been designed.',
                );
            };
        },

        'EDCM$' => sub {
            my ($wsheet) = @_;
            push @{ $arp->{finishClosures} }, sub {
                $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                    name  => 'EDCM information',
                    lines => 'This sheet has not yet been designed.',
                );
            };
        },

        'DCP 087$' => sub {
            my ($wsheet) = @_;
            push @{ $arp->{finishClosures} }, sub {
                $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                    name  => 'DCP 087 (smoothing) calculations',
                    lines => 'This sheet is under construction. '
                      . 'It will mirror "Smoothed Input Details" in the current ARP.',
                );
              }
        },

      ),

      'Schedule 15$' => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 0,   60 );
        $wsheet->set_column( 1, 254, 20 );
        $wsheet->freeze_panes( 0, 1 );
        push @{ $arp->{finishClosures} }, sub {
            Notes( name => 'Analysis of allowed revenue (DCUSA schedule 15)' )
              ->wsWrite( $wbook, $wsheet );
            my @t1001 = map {
                     $_->{table1001}
                  && $_->{targetRevenue} !~ /DCP132longlabels/i
                  ? $_->{table1001}
                  : undef
            } @{ $arp->{models} };
            my ($first1001) = grep { $_ } @t1001 or return;
            my $rowset     = $first1001->{columns}[0]{rows};
            my $rowformats = $first1001->{columns}[3]{rowFormats};
            $wbook->{noLinks} = 1;
            my $needNote1 = 1;
            Columnset(
                name => 'Schedule 15 table 1',
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
                                            [ map { '' } @{ $rowset->{list} } ]
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
                                table1000 =>
                                  $arp->{models}[ $_ - 1 ]{table1000},
                                wsPrepare => sub {
                                    my ( $self, $wb, $ws, $format, $formula,
                                        $pha, $rowh, $colh )
                                      = @_;
                                    my ( $w, $r, $c ) =
                                      $self->{table1000}->wsWrite( $wb, $ws );
                                    $self->{name} = q%='%
                                      . $w->get_name . q%'!%
                                      . xl_rowcol_to_cell( $r, $c + 1 );
                                    if ($needNote1) {
                                        undef $needNote1;
                                        push
                                          @{ $self->{location}{postWriteCalls}
                                              {$wb} }, sub {
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
                                            IV1 => xl_rowcol_to_cell(
                                                $rowh->{IV1} + $y,
                                                $colh->{IV1},
                                                1
                                            )
                                          )
                                          : (
                                            '',
                                            $boldFormat,
                                            $formula->[1],
                                            IV2 => xl_rowcol_to_cell(
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
      },

      'Statistics$' => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 255, 50 );
        $wsheet->set_column( 1, 255, 20 );
        $wsheet->freeze_panes( 0, 1 );
        push @{ $arp->{finishClosures} }, sub {
            $wbook->{noLinks} = 1;
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name  => 'Statistical time series',
                lines => [
                        'The items below are only examples of some of'
                      . ' the statistical outputs that could'
                      . ' be included in this sheet.',
                ],
              ),
              $arp->statisticsColumnsets( $wbook, $wsheet );
        };
      }
      unless @{ $arp->{historical} } || @{ $arp->{scenario} };

    push @{ $arp->{models} }, $model;

    if ( ref $model->{dataset} eq 'HASH' && !$model->{dataset}{baseDataset} ) {
        push @{ $arp->{historical} }, $model;
        return @pairs;
    }

    unless ( $arp->{assumptionColumns} ) {
        $arp->{assumptionColumns} = [];
        $arp->{assumptionRowset}  = Labelset(
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
              foreach $arp->{assumptions} = Notes( name => 'Assumptions' );
            my $headerRowForLater = ++$wsheet->{nextFree};
            ++$wsheet->{nextFree};
            push @{ $arp->{finishClosures} }, sub {

                for ( my $i = 0 ; $i < @{ $arp->{assumptionColumns} } ; ++$i ) {
                    my $model = $arp->{assumptionColumns}[$i]{model};
                    my ( $w, $r, $c ) =
                      $model->{table1000}->wsWrite( $wbook, $wsheet );
                    $wsheet->write(
                        $headerRowForLater + 1,
                        $i + 1,
                        q%="To "&'%
                          . $w->get_name . q%'!%
                          . xl_rowcol_to_cell( $r, $c + 1 ),
                        $wbook->getFormat('thc')
                    );
                    ( $w, $r, $c ) =
                      $model->{sourceModel}{table1000}
                      ->wsWrite( $wbook, $wsheet );
                    $wsheet->write(
                        $headerRowForLater,
                        $i + 1,
                        q%="From "&'%
                          . $w->get_name . q%'!%
                          . xl_rowcol_to_cell( $r, $c + 1 ),
                        $wbook->getFormat('thc')
                    );
                }
            };
            $_->wsWrite( $wbook, $wsheet ) foreach Columnset(
                name      => '',
                noHeaders => 1,
                columns   => $arp->{assumptionColumns},
            );
            $wbook->{logger}      = $logger;
            $wbook->{titleAppend} = $titleAppend;
            $wbook->{noLinks}     = $noLinks;
        };
    }

    push @{ $arp->{scenario} }, $model;

    push @{ $arp->{assumptionColumns} },
      $arp->{assumptionsByModel}{ 0 + $model } = Constant(
        model         => $model,
        rows          => $arp->{assumptionRowset},
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
    my ( $arp, $model, $sourceModel ) = @_;
    my @assumptionsColumnLocationArray;
    sub {
        my ( $wb, $ws, $row ) = @_;
        unless ( $row =~ /^[0-9]+$/s ) {
            my $q = qr/$row/;
            ($row) = grep { $arp->{assumptionRowset}{list}[$_] =~ /$q/; }
              0 .. $#{ $arp->{assumptionRowset}{list} };
        }
        unless (@assumptionsColumnLocationArray) {
            @assumptionsColumnLocationArray =
              $arp->{assumptionsByModel}{ 0 + $model }->wsWrite( $wb, $ws );
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
    my ( $arp, $wbook, $wsheet ) = @_;
    return unless $arp->{statsRows};
    my $rows = Labelset( list => $arp->{statsRows} );
    my @columns =
      map {
        if ( my $tableref = $arp->{statsMap}{ 0 + $_ } ) {
            my ( $w, $r, $c ) =
              $_->{table1000}->wsWrite( $wbook, $wsheet );
            SpreadsheetModel::Custom->new(
                name => q%='%
                  . $w->get_name . q%'!%
                  . xl_rowcol_to_cell( $r, $c + 1 ),
                rows      => $rows,
                custom    => [ map { "=IV1$_"; } 0 .. $#$tableref ],
                arguments => {
                    map {
                        my $t = $tableref->[$_];
                        $t ? ( "IV1$_" => $t ) : ();
                    } 0 .. $#$tableref
                },
                wsPrepare => sub {
                    my ( $self, $wb, $ws, $format, $formula,
                        $pha, $rowh, $colh ) = @_;
                    my @formats = map {
                        if ($_) {
                            local $_ = $_->{defaultFormat}
                              || '0.001soft';
                            s/(?:soft|hard|con)/copy/ unless ref $_;
                            $wb->getFormat($_);
                        }
                        else { undef; }
                    } @$tableref;
                    sub {
                        my ( $x, $y ) = @_;
                        my $t = $y;
                        $x = 0;
                        $y = 0;
                        my $ph = "IV1$t";
                        '', $formats[$t], $formula->[$t], $ph,
                          xl_rowcol_to_cell(
                            $rowh->{$ph} + $x,
                            $colh->{$ph} + $y,
                            1, 1,
                          );
                    };
                },
            );
        }
        else {
            ();
        }
      } @{ $arp->{models} };
    Columnset( name => '', columns => \@columns, );
}

sub addStats {
    my ( $arp, $model, @tables ) = @_;
    foreach my $table (@tables) {
        push @{ $model->{statisticsForArp} }, $table;

        # Only single-cell tables implemented at present
        my $name      = "$table->{name}";
        my $rowNumber = $arp->{statsRowMap}{$name};
        unless ( defined $rowNumber ) {
            push @{ $arp->{statsRows} }, $table->{name};
            $rowNumber = $arp->{statsRowMap}{$name} = $#{ $arp->{statsRows} };
        }
        $arp->{statsMap}{ 0 + $model }[$rowNumber] = $table;
    }
}

sub finish {
    $_->() foreach @{ $_[0]{finishClosures} };
}

1;
