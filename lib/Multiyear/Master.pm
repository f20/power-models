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
use Multiyear::DataDerivative;
use Multiyear::Assumptions;
use Multiyear::Schedule15;
use Multiyear::Statistics;
use SpreadsheetModel::Shortcuts ':all';

sub sheetPriority {
    my ( $model, $sheet ) = @_;
    my $score = {
        'Index$'        => 9,
        'Assumptions$'  => 7,
        'Schedule 15$'  => 6,
        'Tariffs$'      => 5,
        'Aggregates$'   => 4,
        'Illustrative$' => 3,
    }->{$sheet};
    $score;
}

sub new {
    my $class = shift;
    my $model = bless {
        historical       => [],
        scenario         => [],
        statsAssumptions => [],
        statsSections    => [ split /\n/, <<EOL ],
Average pence per unit for each tariff
Average charge per MPAN for each tariff
Illustrative charges (£/MWh)
Illustrative charges (£/year)
LDNO margins for illustrative customers
QNO margins for illustrative customers
General input data
Load characteristics
Network model and related input data
Service models
DNO-wide aggregated
Revenue by tariff
Units distributed by tariff
MPANs by tariff
EOL
        @_,
    }, $class;
    ${ $model->{sharingObjectRef} } = $model if $model->{sharingObjectRef};
    $model;
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
    my ( $me, $wbook, $wsheet ) = @_;
    sub {
        my $noLinks = delete $wbook->{noLinks};
        $wsheet->set_column( 0, 0,   70 );
        $wsheet->set_column( 1, 255, 14 );
        require SpreadsheetModel::Book::FrontSheet;
        my $noticeMaker =
          SpreadsheetModel::Book::FrontSheet->new( model => $me );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name      => 'Multiple CDCM models',
            copyright => 'Copyright 2009-2011 Energy Networks '
              . 'Association Limited and others. '
              . 'Copyright 2011-2017 Franck Latrémolière, Reckon LLP and others.'
          ),
          $noticeMaker->extraNotes,   $noticeMaker->dataNotes,
          $noticeMaker->licenceNotes, @{ $me->{historical} }
          ? Notes(
            name        => 'Historical models',
            sourceLines => [
                map {
                    [
                        $_->{nickName} || 'Historical model',
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
                        @{ $me->{scenario}[$_]{sheetLinks} },
                    ];
                } 0 .. $#{ $me->{scenario} }
            ],
          )
          : ();
        $wbook->{noLinks} = $noLinks if defined $noLinks;
    };
}

sub worksheetsAndClosures {

    my ( $me, $wbook ) = @_;

    push @{ $me->{finishClosures} }, sub {
        delete $wbook->{titleAppend};
        $wbook->{noLinks} = 1;
        delete $wbook->{logger};
    };

    'Index$' => sub {
        my ($wsheet) = @_;
        push @{ $me->{finishClosures} },
          $me->indexFinishClosure( $wbook, $wsheet );
      },

      $me->{assumptionsSheet}
      ? (
        'Assumptions$' => sub {
            my ($wsheet) = @_;
            $wsheet->set_column( 0, 255, 50 );
            $wsheet->set_column( 1, 255, 20 );
            $wsheet->freeze_panes( 0, 1 );
            my $logger      = delete $wbook->{logger};
            my $titleAppend = delete $wbook->{titleAppend};
            my $noLinks     = $wbook->{noLinks};
            $wbook->{noLinks} = 1;
            Notes( name => 'Assumptions' )->wsWrite( $wbook, $wsheet );

            my $table1001headerRowForLater;
            if (
                my @table1001Overridable =
                map {
                    !$_->{table1001_2016}
                      ? ()
                      : [ $_, $_->{table1001_2016}{columns}[3] ];
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
                                    map {
                                            /Price index|RPI index/i
                                          ? '0.000hard'
                                          : undef;
                                    } @{ $rows->{list} }
                                ],
                                data => [
                                    map { defined $_ ? '#N/A' : undef; }
                                      @{ $_->[1]{data} }
                                ],
                            )
                        );
                    } @table1001Overridable
                };

                Notes( name => '151. Schedule 15 input data in £ million' )
                  ->wsWrite( $wbook, $wsheet );
                $table1001headerRowForLater = ++$wsheet->{nextFree};
                Columnset(
                    name            => '',
                    number          => 151,
                    dataset         => $me->{dataset},
                    noHeaders       => 1,
                    ignoreDatasheet => 1,
                    columns         => [
                        map { $me->{table1001Overrides}{ 0 + $_->[0] } }
                          @table1001Overridable
                    ],
                )->wsWrite( $wbook, $wsheet );

            }

            Notes( name => '160. Percentage changes in costs and volumes' )
              ->wsWrite( $wbook, $wsheet );
            my $headerRowForLater = ++$wsheet->{nextFree};
            ++$wsheet->{nextFree};
            Columnset(
                name            => '',
                number          => 160,
                dataset         => $me->{dataset},
                noHeaders       => 1,
                ignoreDatasheet => 1,
                columns         => $me->{percentageAssumptionColumns},
            )->wsWrite( $wbook, $wsheet );

            my $headerRow165;
            if ( $me->{overrideAssumptionColumns} ) {
                Notes( name => '165. Other input data' )
                  ->wsWrite( $wbook, $wsheet );
                $headerRow165 = ++$wsheet->{nextFree};
                Columnset(
                    name            => '',
                    number          => 165,
                    dataset         => $me->{dataset},
                    noHeaders       => 1,
                    ignoreDatasheet => 1,
                    columns         => $me->{overrideAssumptionColumns},
                )->wsWrite( $wbook, $wsheet );
            }

            push @{ $me->{finishClosures} }, sub {
                my $thc = $wbook->getFormat('thc');
                for (
                    my $i = 0 ;
                    $i < @{ $me->{percentageAssumptionColumns} } ;
                    ++$i
                  )
                {
                    my $model = $me->{percentageAssumptionColumns}[$i]{model};
                    my $id = $me->modelIdentifier( $model, $wbook, $wsheet );
                    $wsheet->write( $_, $i + 1, $id, $thc )
                      foreach grep { defined $_; } $table1001headerRowForLater,
                      $headerRow165;
                    $id =~ s/^=/="To "&/;
                    $wsheet->write( $headerRowForLater + 1, $i + 1, $id, $thc );
                    $id =
                      $me->modelIdentifier( $model->{sourceModel}, $wbook,
                        $wsheet );
                    $id =~ s/^=/="From "&/;
                    $wsheet->write( $headerRowForLater, $i + 1, $id, $thc );
                }
            };

            $wbook->{logger}      = $logger;
            $wbook->{titleAppend} = $titleAppend;
            $wbook->{noLinks}     = $noLinks;

        }
      )
      : (),

      'Schedule 15$' => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 0,   60 );
        $wsheet->set_column( 1, 254, 16 );
        $wsheet->freeze_panes( 0, 1 );
        push @{ $me->{finishClosures} }, sub {
            $_->wsWrite( $wbook, $wsheet )
              foreach map { $me->schedule15($_); } qw(2016 2012);
        };
      },

      'Illustrative$' => sub {
        my ($wsheet) = @_;
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
                sub { $_[0] =~ /illustrative/i; } ),
              $me->changeColumnsets(
                sub {
                    $_[0] =~ /illustrative/i || $_[0] !~ m¢£/year¢;
                }
              );
        };
      },

      'Aggregates$' => sub {
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
                sub { $_[0] !~ /input|illustrative/i; } ),
              $me->changeColumnsets(
                sub {
                    $_[0] !~ /illustrative/i;
                }
              );
        };
      };

}

sub finish {
    $_->() foreach @{ $_[0]{finishClosures} };
}

1;
