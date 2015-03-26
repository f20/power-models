package CDCM;

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
use SpreadsheetModel::Derivative;

sub derivativeDataset {

    my ( $model, $sourceModel ) = @_;

    my $addSourceDatasetAdjuster =
      SpreadsheetModel::Derivative::setupDerivativeDataset( $model,
        $sourceModel );

    my $table1001data;
    $table1001data = delete $model->{dataset}{1001}
      if $model->{dataset} && $model->{dataset}{1001};

    if (
        $model->{sharedData}
        && ( my $getAssumptionCell =
            $model->{sharedData}->assumptionsLocator( $model, $sourceModel ) )
      )
    {

        $addSourceDatasetAdjuster->(
            1001 => sub {
                my ( $cell, $row, $col, $wb, $ws, $irow ) = @_;
                my $hardData;
                $hardData = $table1001data->[$col]{$irow}
                  || 0
                  if $table1001data
                  && $table1001data->[$col]
                  && defined $table1001data->[$col]{$irow};
                return defined $hardData ? $hardData : "=$cell"
                  unless $col == 4;
                my $preOverride =
                  defined $hardData                  ? $hardData
                  : $row =~ /RPI Indexation Factor/i ? ( "(1+"
                      . $getAssumptionCell->( $wb, $ws, 'RPI' )
                      . ")*$cell" )
                  : $cell;
                if ( my $override =
                    $model->{sharedData}
                    ->table1001Overrides( $model, $wb, $ws, $row ) )
                {
                    return
                      qq%=IF(ISERROR(0+$override),$preOverride,1e6*$override)%;
                }
                "=$preOverride";
            }
        );

        $addSourceDatasetAdjuster->(
            1020 => sub {
                ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                my $ac = $getAssumptionCell->(
                    $wb, $ws,
                    /132kV\/EHV/     ? 2
                    : /132kV\/HV/    ? 5
                    : /132kV/        ? 1
                    : /EHV\/HV/      ? 4
                    : /EHV/          ? 3
                    : /HV\/LV/       ? 8
                    : /HV/           ? 6
                    : /LV circuits/i ? 9
                    :                  10
                );
                "=(1+$ac)*$cell";
            }
        );

        {
            my $ac;
            $addSourceDatasetAdjuster->(
                1022 => sub {
                    ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                    $ac ||= $getAssumptionCell->( $wb, $ws, 10 );
                    "=(1+$ac)*$cell";
                }
            );
        }

        {
            my $ac;
            $addSourceDatasetAdjuster->(
                1023 => sub {
                    ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                    $ac ||= $getAssumptionCell->( $wb, $ws, 7 );
                    "=(1+$ac)*$cell";
                }
            );
        }

        my $unauthInSource =
          $sourceModel->{unauth} && $sourceModel->{unauth} =~ /day/i;
        $addSourceDatasetAdjuster->(
            1053 => sub {
                ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                my $ac = $getAssumptionCell->(
                    $wb, $ws,
                    $col < 4
                    ? ( /unmet/i ? 21 : /gener/i ? 22 : /half[ -]hourly/i
                          && !/aggreg/i ? 17 : 15 )
                    : $col == 4 ? ( /gener/i ? 23 : /half[ -]hourly/i
                          && !/aggreg/i ? 18 : 16 )
                    : $col == 5 || $unauthInSource && $col == 6 ? 19
                    : ( /gener/i ? 24 : 20 ),
                );
                "=(1+$ac)*$cell";
            }
        );

        if (  !$unauthInSource
            && $model->{unauth}
            && $model->{unauth} =~ /day/i )
        {    # Adjust if unauthorised capacity is being introduced
            my $original = $model->{dataset}{1053};
            $model->{dataset}{1053} = sub {
                my $d = $original->(@_);
                splice @$d, 6, 0, { map { $_ => ''; } keys %{ $d->[5] } };
                $d;
            };
        }
        elsif ( $unauthInSource
            and !$model->{unauth} || $model->{unauth} !~ /day/i )
        {    # Adjust if unauthorised capacity is being removed
            my $original = $model->{dataset}{1053};
            $model->{dataset}{1053} = sub {
                my $d = $original->(@_);
                splice @$d, 5, 2, {
                    map {
                        $_ => '=' . join '+',
                          map { local $_ = $_ || 0; s/^=//s; $_; } $d->[5]{$_},
                          $d->[6]{$_};
                    } keys %{ $d->[5] }
                };
                $d;
            };
        }

        $addSourceDatasetAdjuster->(
            1055 => sub {
                ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                my $ac = $getAssumptionCell->( $wb, $ws, 14 );
                "=(1+$ac)*$cell";
            }
        );

        $addSourceDatasetAdjuster->(
            1059 => sub {
                ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                my $no;
                $no = 11 if $col == 1;
                $no = 12 if $col == 2;
                $no = 13 if $col == 4;
                join '', '=',
                  $no
                  ? ( '(1+', $getAssumptionCell->( $wb, $ws, $no ), ')*' )
                  : (), $cell;
            }
        );

    }

    $addSourceDatasetAdjuster->( $_ => sub { my ($cell) = @_; "=$cell"; } )
      foreach grep { /^[0-9]+$/s } keys %{ $model->{dataset} };

}

1;
