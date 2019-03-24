package CDCM;

# Copyright 2009-2011 Energy Networks Association Limited and others.
# Copyright 2011-2017 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';
use SpreadsheetModel::SegmentRoot;

sub flexibleAdder {

    # single adder with caps and collars, with flexibility to include generation

    my ( $model, $allTariffsByEndUser, $nonExcludedComponents, $componentMap,
        $volumeData, $tariffsExMatching, $adderAmount, )
      = @_;

    my $generationWeight = $model->{scaler} =~ /gen([0-9.-]+)/ ? $1 : 0;

    my @adderWeightsItems = (
        name => 'Tariff-specific weighting for revenue matching adder',
        rows => $allTariffsByEndUser,
        data => [
            map { /gener/i ? $generationWeight : 1; }
              @{ $allTariffsByEndUser->{list} }
        ],
    );
    my $adderWeights =
      $model->{scaler} =~ /editable/i
      ? Dataset(
        @adderWeightsItems,
        appendTo           => $model->{inputTables},
        dataset            => $model->{dataset},
        number             => 1079,
        usePlaceholderData => 1,
      )
      : Constant(@adderWeightsItems);

    my @columns = grep { /kWh/ } @$nonExcludedComponents;
    my @slope = map {
        Arithmetic(
            name          => "Effect through $_",
            arithmetic    => '=10*A1*A2',
            defaultFormat => '0soft',
            arguments     => {
                A2 => $adderWeights,
                A1 => $volumeData->{$_},
            },
        );
    } @columns;

    my $slopeSet = Columnset(
        name    => 'Marginal revenue effect of adder',
        columns => \@slope
    );

    my %min = map {
        my $tariffComponent = $_;
        $_ => $model->{scaler} =~ /zero/i
          ? undef
          : Dataset(
            name       => "Minimum $_",
            rows       => $allTariffsByEndUser,
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => -999_999.999,
                maximum  => 999_999.999,
            },
            usePlaceholderData => 1,
            data               => [
                map { $componentMap->{$_}{$tariffComponent} ? 0 : undef; }
                  @{ $allTariffsByEndUser->{list} }
            ],
          );
    } @columns;

    if ( my @cols = grep { $_ } @min{@columns} ) {
        Columnset(
            name     => 'Minimum rate',
            number   => 1077,
            columns  => \@cols,
            dataset  => $model->{dataset},
            appendTo => $model->{inputTables},
        );
    }

    my %max = map {
        my $tariffComponent = $_;
        $_ => $model->{scaler} =~ /zero/i
          ? undef
          : Dataset(
            name       => "Maximum $_",
            rows       => $allTariffsByEndUser,
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => -999_999.999,
                maximum  => 999_999.999,
            },
            usePlaceholderData => 1,
            data               => [
                map { $componentMap->{$_}{$tariffComponent} ? '' : undef; }
                  @{ $allTariffsByEndUser->{list} }
            ],
          );
    } @columns;

    if ( my @cols = grep { $_ } @max{@columns} ) {
        Columnset(
            name     => 'Maximum rate',
            number   => 1078,
            columns  => \@cols,
            dataset  => $model->{dataset},
            appendTo => $model->{inputTables},
        );
    }

    my %minAdder = map {
        my $tariffComponent = $_;
        $_ => Arithmetic(
            name       => "Minimum adder threshold for $_",
            arithmetic => '=IF(A91>0,'
              . (
                $min{$_} ? 'IF(ISNUMBER(A3),(A2-A1)/A92,"")'
                : 'IF(A3<0,"",0-A1/A92)'
              )
              . ',IF(A93<0,'
              . (
                $max{$_} ? 'IF(ISNUMBER(A31),(A21-A11)/A94,"")'
                : 'IF(A31<0,0-A11/A94,"")'
              )
              . ',0))',
            arguments => {
                A1  => $tariffsExMatching->{$_},
                A11 => $tariffsExMatching->{$_},
                A2  => $min{$_} || $tariffsExMatching->{$_},
                A3  => $min{$_} || $tariffsExMatching->{$_},
                A21 => $max{$_} || $tariffsExMatching->{$_},
                A31 => $max{$_} || $tariffsExMatching->{$_},
                A91 => $adderWeights,
                A92 => $adderWeights,
                A93 => $adderWeights,
                A94 => $adderWeights,
            },
            rowFormats => [
                map {
                    $componentMap->{$_}{$tariffComponent} ? undef
                      : 'unavailable';
                } @{ $allTariffsByEndUser->{list} }
            ]
        );
    } @columns;

    my %maxAdder = map {
        my $tariffComponent = $_;
        $_ => Arithmetic(
            name       => "Maximum adder threshold for $_",
            arithmetic => '=IF(A91<0,'
              . (
                $min{$_} ? 'IF(ISNUMBER(A3),(A2-A1)/A92,"")'
                : 'IF(A3<0,"",0-A1/A92)'
              )
              . ',IF(A93>0,'
              . (
                $max{$_} ? 'IF(ISNUMBER(A31),(A21-A11)/A94,"")'
                : 'IF(A31<0,0-A11/A94,"")'
              )
              . ',0))',
            arguments => {
                A1  => $tariffsExMatching->{$_},
                A11 => $tariffsExMatching->{$_},
                A2  => $min{$_} || $tariffsExMatching->{$_},
                A3  => $min{$_} || $tariffsExMatching->{$_},
                A21 => $max{$_} || $tariffsExMatching->{$_},
                A31 => $max{$_} || $tariffsExMatching->{$_},
                A91 => $adderWeights,
                A92 => $adderWeights,
                A93 => $adderWeights,
                A94 => $adderWeights,
            },
            rowFormats => [
                map {
                    $componentMap->{$_}{$tariffComponent} ? undef
                      : 'unavailable';
                } @{ $allTariffsByEndUser->{list} }
            ]
        );
    } @columns;

    my $adderRate = new SpreadsheetModel::SegmentRoot(
        name   => 'General adder rate (p/kWh)',
        slopes => $slopeSet,
        target => $adderAmount,
        min    => Columnset(
            name    => 'Adder value at which the minimum is breached',
            columns => [ @minAdder{@columns} ]
        ),
        max => Columnset(
            name    => 'Adder value at which the maximum is breached',
            columns => [ @maxAdder{@columns} ]
        ),
    );

    +{
        map {
            my $tariffComponent = $_;
            my $iv              = 'A1';
            $iv = "MAX($iv,A5)" if $minAdder{$_};
            $iv = "MIN($iv,A6)" if $maxAdder{$_};
            $_ => Arithmetic(
                name       => "Adder on $_",
                rows       => $allTariffsByEndUser,
                cols       => Labelset( list => ['Adder'] ),
                arithmetic => "=A2*$iv",
                arguments  => {
                    A1 => $adderRate,
                    A2 => $adderWeights,
                    $minAdder{$_} ? ( A5 => $minAdder{$_} ) : (),
                    $maxAdder{$_} ? ( A6 => $maxAdder{$_} ) : ()
                },
                rowFormats => [
                    map {
                        $componentMap->{$_}{$tariffComponent}
                          ? undef
                          : 'unavailable';
                    } @{ $allTariffsByEndUser->{list} }
                ]
            );
        } @columns
    };

}

1;
