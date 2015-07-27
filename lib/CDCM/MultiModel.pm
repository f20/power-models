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
General input data
Load characteristics
Network model and related input data
Service models
DNO-wide aggregated
Revenue by tariff
Units distributed by tariff
MPANs by tariff
Average pence per unit for each tariff
Average charge per MPAN for each tariff
Illustrative charges (£/MWh)
Illustrative charges (£/year)
EOL
      },
      shift;
}

sub modelIdentifier {
    my ( $me, $model, $wb, $ws ) = @_;
    unless ( $me->{identification} ) {
        my ( %coHash, %yrHash, %vHash );
        foreach ( @{ $me->{models} } ) {
            ++$vHash{ $_->{version} || '' };
            local $_ = $_->{'~datasetName'} || '';
            tr/-/ /;
            s/ (20[0-9][0-9] [0-9][0-9])//;
            ++$yrHash{ $1 || '' };
            ++$coHash{ $_ || '' };
        }
        $me->{identification} = {
              meth => ( !grep { $_ > 1 } values %yrHash ) ? 1
            : ( !grep { $_ > 1 } values %coHash ) ? 0
            : ( !grep { $_ > 1 } values %vHash )  ? 3
            :                                       4
        };
    }
    return $me->{identification}{ 0 + $model } ||=
      '='
      . $model->modelIdentification( $wb, $ws )->[ $me->{identification}{meth} ]
      if $me->{identification}{meth} < 3;
    return $me->{identification}{ 0 + $model } ||= join ' ',
      grep { $_ } @{$model}{qw(version ~datasetName)}
      if $me->{identification}{meth} == 3;
    return $me->{identification}{ 0 + $model } ||= '=' . join '&" "&',
      @{ $model->modelIdentification( $wb, $ws ) };
}

sub indexFinishClosure {
    my ( $me, $wbook, $wsheet, $model ) = @_;
    sub {
        my $noLinks = delete $wbook->{noLinks};
        $wsheet->set_column( 0, 0,   70 );
        $wsheet->set_column( 1, 255, 14 );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name  => 'Multiple CDCM models',
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

Copyright 2009-2011 Energy Networks Association Limited and others. Copyright 2011-2015 Franck Latrémolière, Reckon LLP and others. 
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
        $_->wsWrite( $wbook, $wsheet ) foreach @{ $me->{historical} }
          ? Notes(
            name        => 'Historical models',
            sourceLines => [
                map {
                    [
                        $_->{nickName} || 'Historical model',
                        undef,
                        @{ $_->{sheetLinks} },
                    ];
                } @{ $me->{historical} }
            ],
          )
          : (), @{ $me->{scenario} } ? Notes(
            name        => 'Scenario models',
            sourceLines => [
                map {
                    [
                        $me->{scenario}[$_]{nickName} || 'Scenario model',
                        $me->{assumptionColumns}[$_],
                        @{ $me->{scenario}[$_]{sheetLinks} },
                    ];
                } 0 .. $#{ $me->{scenario} }
            ],
          )
          : ();
        $wbook->{noLinks} = $noLinks if defined $noLinks;
    };
}

sub sheetsForFirstModel {

    my ( $me, $model, $wbook ) = @_;

    push @{ $me->{finishClosures} }, sub {
        delete $wbook->{logger};
        delete $wbook->{titleAppend};
        delete $wbook->{noLinks};
    };

    'Index$' => sub {
        my ($wsheet) = @_;
        push @{ $me->{finishClosures} },
          $me->indexFinishClosure( $wbook, $wsheet, $model );
      },

      'Schedule 15$' => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 0,   60 );
        $wsheet->set_column( 1, 254, 16 );
        $wsheet->freeze_panes( 0, 1 );
        push @{ $me->{finishClosures} }, sub {

            my @t1001 = map {
                     $_->{table1001}
                  && $_->{targetRevenue} !~ /DCP132longlabels/i
                  ? $_->{table1001}
                  : undef;
            } @{ $me->{models} };
            Notes( name => 'Allowed revenue summary (DCUSA schedule 15)', )
              ->wsWrite( $wbook, $wsheet );

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
                                      $me->modelIdentifier( $self->{model},
                                        $wb, $ws );
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
                                            qr/\bA1\b/ => xl_rowcol_to_cell(
                                                $rowh->{A1} + $y,
                                                $colh->{A1},
                                                1
                                            )
                                          )
                                          : (
                                            '',
                                            $boldFormat,
                                            $formula->[1],
                                            qr/\bA2\b/ => xl_rowcol_to_cell(
                                                $rowh->{A2} + $y,
                                                $colh->{A2},
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

      'Illustrative$' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber}     = 12;
        $wsheet->set_column( 0, 255, 64 );
        $wsheet->set_column( 1, 255, 16 );
        $wsheet->freeze_panes( 0, 1 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Illustrative charges', ),
          @{ $me->{statsAssumptions} };
        push @{ $me->{finishClosures} }, sub {
            $wbook->{noLinks} = 1;
            $_->wsWrite( $wbook, $wsheet )
              foreach $me->statisticsColumnsets( $wbook, $wsheet,
                sub { $_[0] =~ /illustrative/i; } );
        };
      },

      'Other$' => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 255, 50 );
        $wsheet->set_column( 1, 255, 16 );
        $wsheet->freeze_panes( 0, 1 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Other statistics', );
        push @{ $me->{finishClosures} }, sub {
            $wbook->{noLinks} = 1;
            $_->wsWrite( $wbook, $wsheet )
              foreach $me->statisticsColumnsets( $wbook, $wsheet,
                sub { $_[0] !~ /input|illustrative/i; } );
        };
      },

      'Changes$' => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 255, 50 );
        $wsheet->set_column( 1, 255, 16 );
        $wsheet->freeze_panes( 0, 1 );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes( name => 'Changes', );
        push @{ $me->{finishClosures} }, sub {
            $wbook->{noLinks} = 1;
            $_->wsWrite( $wbook, $wsheet ) foreach $me->changeColumnsets(
                sub {
                    # $_[0] =~ /input/i ||
                    $_[0] =~ /illustrative/i && $_[0] !~ m¢£/year¢;
                }
            );
        };
      };

}

sub assumptionsClosure {
    my ( $me, $wbook ) = @_;
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
    sub {
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

        my $table1001headerRowForLater;
        if (
            my @table1001Overridable =
            map {
                    !$_->{table1001}
                  || $_->{targetRevenue} =~ /DCP132longlabels/i
                  ? ()
                  : [ $_, $_->{table1001}{columns}[3] ];
            } @{ $me->{scenario} }
          )
        {
            my $rows = $table1001Overridable[0][1]{rows};
            $me->{table1001Overrides} = {
                map {
                    (
                        0 + $_->[0] => Dataset(
                            name          => '',
                            rows          => $rows,
                            defaultFormat => '0.0hard',
                            rowFormats    => [
                                map { /RPI|\bIndex\b/i ? '0.000hard' : undef; }
                                  @{ $rows->{list} }
                            ],
                            data => [
                                map { defined $_ ? '#N/A' : undef; }
                                  @{ $_->[1]{data} }
                            ],
                            usePlaceholderData => 1,
                        )
                    );
                } @table1001Overridable
            };
            Notes( name => 'DCUSA schedule 15 input data in £ million' )
              ->wsWrite( $wbook, $wsheet );
            $table1001headerRowForLater = ++$wsheet->{nextFree};
            Columnset(
                name            => '',
                noHeaders       => 1,
                ignoreDatasheet => 1,
                columns         => [
                    map { $me->{table1001Overrides}{ 0 + $_->[0] } }
                      @table1001Overridable
                ],
            )->wsWrite( $wbook, $wsheet );
        }

        Notes( name => 'Assumed rates of change in costs and volumes' )
          ->wsWrite( $wbook, $wsheet );
        my $headerRowForLater = ++$wsheet->{nextFree};
        ++$wsheet->{nextFree};
        $_->wsWrite( $wbook, $wsheet ) foreach Columnset(
            name            => '',
            noHeaders       => 1,
            ignoreDatasheet => 1,
            columns         => $me->{assumptionColumns},
        );
        push @{ $me->{finishClosures} }, sub {
            my $thc = $wbook->getFormat('thc');
            for ( my $i = 0 ; $i < @{ $me->{assumptionColumns} } ; ++$i ) {
                my $model = $me->{assumptionColumns}[$i]{model};
                my $id = $me->modelIdentifier( $model, $wbook, $wsheet );
                $wsheet->write( $headerRowForLater + 1, $i + 1, $id, $thc );
                $wsheet->write( $table1001headerRowForLater, $i + 1, $id, $thc )
                  if defined $table1001headerRowForLater;
                $wsheet->write(
                    $headerRowForLater,
                    $i + 1,
                    $me->modelIdentifier(
                        $model->{sourceModel}, $wbook, $wsheet
                    ),
                    $thc
                );
            }
        };
        $wbook->{logger}      = $logger;
        $wbook->{titleAppend} = $titleAppend;
        $wbook->{noLinks}     = $noLinks;
    };
}

sub worksheetsAndClosuresMulti {
    my ( $me, $model, $wbook, @pairs ) = @_;
    unshift @pairs, $me->sheetsForFirstModel( $model, $wbook )
      unless @{ $me->{historical} } || @{ $me->{scenario} };
    push @{ $me->{models} }, $model;
    my $assumptionZero;
    if ( ref $model->{dataset} eq 'HASH' ) {
        if ( my $sourceModel = $me->{modelByDataset}{ 0 + $model->{dataset} } )
        {
            $model->{sourceModel} = $sourceModel;
            $assumptionZero = 1;
        }
        elsif ( !$model->{sourceModel} ) {
            $me->{modelByDataset}{ 0 + $model->{dataset} } = $model;
            push @{ $me->{historical} }, $model;
            return @pairs;
        }
    }
    else {
        push @{ $me->{historical} }, $model;
        return @pairs;
    }
    unshift @pairs, 'Assumptions$' => $me->assumptionsClosure($wbook)
      unless $me->{assumptionColumns};
    push @{ $me->{scenario} }, $model;
    push @{ $me->{assumptionColumns} },
      $me->{assumptionsByModel}{ 0 + $model } = Dataset(
        name          => 'Assumptions',
        model         => $model,
        rows          => $me->{assumptionRowset},
        defaultFormat => '%hardpm',
        data          => [
            [
                $assumptionZero
                ? ( map { '' } @{ $me->{assumptionRowset}{list} } )
                : (
                    0.03,
                    qw(0.02 0.02 0.02 0.02 0.02 0.02 0.02 0.02 0.02 0.02
                      0.02 0.02 0.02 0.02),
                    1
                    ? qw(0 0 0 0 0 0 0 0 0 0)
                    : qw(-0.01 0
                      0.01 0.01 0.01 0.01
                      0
                      0.03 0.03 0.03)
                )
            ]
        ],
        usePlaceholderData => 1,
      );
    @pairs;
}

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
    my ( $me, $wbook, $wsheet, $filter ) = @_;
    map {
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
                    name => $me->modelIdentifier( $_, $wbook, $wsheet ),
                    rows => $rows,
                    custom    => [ map { "=A1$_"; } 0 .. $#$relevantMap ],
                    arguments => {
                        map {
                            my $t;
                            $t = $relevantMap->[$_][0]
                              if $relevantMap->[$_];
                            $t ? ( "A1$_" => $t ) : ();
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
                            my $ph = "A1$y";
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
        $me->{statsColumnsets}[$_] = Columnset(
            name    => $me->{statsSections}[$_],
            columns => \@columns,
        );
      } grep {
        $me->{statsRows}[$_]
          && @{ $me->{statsRows}[$_] }
          && $filter->( $me->{statsSections}[$_] );
      } 0 .. $#{ $me->{statsSections} };
}

sub changeColumnsets {
    my ( $me, $filter ) = @_;
    my %modelMap =
      map { ( 0 + $me->{models}[$_], $_ ) } 0 .. $#{ $me->{models} };
    my @modelNumbers;
    foreach ( 1 .. $#{ $me->{historical} } ) {
        my $old = $me->{historical}[ $_ - 1 ]{dataset}{1000}[2]
          {'Company charging year data version'};
        my $new = $me->{historical}[$_]{dataset}{1000}[2]
          {'Company charging year data version'};
        next unless $old && $new && $old ne $new;
        push @modelNumbers,
          [
            $modelMap{ 0 + $me->{historical}[ $_ - 1 ] },
            $modelMap{ 0 + $me->{historical}[$_] },
          ];
    }
    foreach ( @{ $me->{assumptionColumns} } ) {
        my $model = $_->{model};
        push @modelNumbers,
          [ $modelMap{ 0 + $model->{sourceModel} }, $modelMap{ 0 + $model }, ];
    }
    map {
        my $cols = $me->{statsColumnsets}[$_]{columns};
        my ( @cola, @colb );
        foreach (@modelNumbers) {
            my ( $before, $after ) = @{$cols}[@$_];
            next unless $before && $after;
            push @cola, Arithmetic(
                name          => $after->{name},
                defaultFormat => '0.000softpm',
                rowFormats    => [
                    map {
                             !defined $before->{rowFormats}[$_]
                          || !defined $after->{rowFormats}[$_] ? undef
                          : $before->{rowFormats}[$_] eq 'unavailable'
                          || $after->{rowFormats}[$_] eq 'unavailable'
                          ? 'unavailable'
                          : eval {
                            local $_ = $after->{rowFormats}[$_];
                            s/copy|soft/softpm/;
                            $_;
                          };
                    } 0 .. $#{ $after->{rows}{list} }
                ],
                arithmetic => '=A1-A2',
                arguments  => { A1 => $after, A2 => $before, },
            );
            push @colb, Arithmetic(
                name          => $after->{name},
                defaultFormat => '%softpm',
                rowFormats    => [
                    map {
                             !defined $before->{rowFormats}[$_]
                          || !defined $after->{rowFormats}[$_] ? undef
                          : $before->{rowFormats}[$_] eq 'unavailable'
                          || $after->{rowFormats}[$_] eq 'unavailable'
                          ? 'unavailable'
                          : undef;
                    } 0 .. $#{ $after->{rows}{list} }
                ],
                arithmetic => '=IF(A2,A1/A3-1,"")',
                arguments => { A1 => $after, A2 => $before, A3 => $before, },
            );
        }
        (
            @cola ? Columnset(
                name    => "Change: $me->{statsColumnsets}[$_]{name}",
                columns => \@cola,
              )
            : (),
            @colb ? Columnset(
                name    => "Relative change: $me->{statsColumnsets}[$_]{name}",
                columns => \@colb,
              )
            : ()
        );
      } grep {
        $me->{statsColumnsets}[$_] && $filter->( $me->{statsSections}[$_] );
      } 0 .. $#{ $me->{statsSections} };

}

sub addStats {
    my ( $me, $section, $model, @tables ) = @_;
    my ($sectionNumber) = grep { $section eq $me->{statsSections}[$_]; }
      0 .. $#{ $me->{statsSections} };
    unless ( defined $sectionNumber ) {
        push @{ $me->{statsSections} }, $section;
        $sectionNumber = $#{ $me->{statsSections} };
    }
    foreach my $table (@tables) {
        my $lastCol = $table->lastCol;
        if ( my $lastRow = $table->lastRow ) {
            for ( my $col = 0 ; $col <= $lastCol ; ++$col ) {
                for ( my $row = 0 ; $row <= $lastRow ; ++$row ) {
                    my $groupid;
                    $groupid = $table->{rows}{groupid}[$row]
                      if $table->{rows}{groupid};
                    my $name = "$table->{rows}{list}[$row]";
                    $name .= " $table->{cols}{list}[$col]" if $lastCol;
                    my $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name};
                    unless ( defined $rowNumber ) {
                        if ( defined $groupid ) {
                            my $group = "$table->{rows}{groups}[$groupid]";
                            my $groupRowNumber =
                              $me->{statsRowMap}[$sectionNumber]{$group};
                            if ( defined $groupRowNumber ) {
                                for (
                                    my $i = $groupRowNumber + 1 ;
                                    $i <=
                                    $#{ $me->{statsRows}[$sectionNumber] } ;
                                    ++$i
                                  )
                                {
                                    if ( $me->{statsRows}[$sectionNumber][$i] !~
                                        /^$group \(/ )
                                    {
                                        $rowNumber = $i;
                                        last;
                                    }
                                }
                            }
                            if ( defined $rowNumber ) {
                                splice @{ $me->{statsRows}[$sectionNumber] },
                                  $rowNumber, 0, $name;
                                foreach my $ma (
                                    values %{ $me->{statsMap}[$sectionNumber] }
                                  )
                                {
                                    for ( my $i = @$ma ;
                                        $i > $rowNumber ; --$i )
                                    {
                                        $ma->[$i] = $ma->[ $i - 1 ];
                                    }
                                }
                                map    { ++$_ }
                                  grep { $_ >= $rowNumber }
                                  values %{ $me->{statsRowMap}[$sectionNumber]
                                  };
                                $me->{statsRowMap}[$sectionNumber]{$name} =
                                  $rowNumber;
                            }
                        }
                        unless ( defined $rowNumber ) {
                            push @{ $me->{statsRows}[$sectionNumber] }, $name;
                            $rowNumber =
                              $me->{statsRowMap}[$sectionNumber]{$name} =
                              $#{ $me->{statsRows}[$sectionNumber] };
                        }
                    }
                    $me->{statsMap}[$sectionNumber]{ 0 + $model }[$rowNumber] =
                      [ $table, $col, $row ]
                      unless $table->{rows}{groupid} && !defined $groupid;
                }
            }
        }
        else {
            for ( my $col = 0 ; $col <= $lastCol ; ++$col ) {
                my $name =
                  UNIVERSAL::can( $table->{name}, 'shortName' )
                  ? $table->{name}->shortName
                  : "$table->{name}";
                $name .= " $table->{cols}{list}[$col]" if $lastCol;
                my $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name};
                unless ( defined $rowNumber ) {
                    push @{ $me->{statsRows}[$sectionNumber] }, $name;
                    $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name} =
                      $#{ $me->{statsRows}[$sectionNumber] };
                }
                $me->{statsMap}[$sectionNumber]{ 0 + $model }[$rowNumber] =
                  [ $table, $col, 0 ];
            }
        }
    }
}

sub finish {
    $_->() foreach @{ $_[0]{finishClosures} };
}

1;
