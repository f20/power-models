package Financial;

=head Copyright licence and disclaimer

Copyright 2015 Franck Latrémolière, Reckon LLP and others.

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
use Financial::Balance;
use Financial::CashCalc;
use Financial::Cashflow;
use Financial::Debt;
use Financial::FixedAssetsUK;
use Financial::Income;
use Financial::Periods;
use Financial::Ratios;
use Financial::Reserve;
use Financial::Sheets;
use Financial::Stream;
use SpreadsheetModel::CalcBlock;
use SpreadsheetModel::Shortcuts ':all';

sub new {

    my $class = shift;

    my $model = bless {
        balanceTables       => [],
        cashflowTables      => [],
        equityRaisingTables => [],
        incomeTables        => [],
        inputTables         => [],
        ratioTables         => [],
        @_,
    }, $class;

    my $sales = Financial::Stream->new(
        $model,
        inputTableNumber => 1430,
        numLines         => $model->{numSales},
    );

    my $costSales = Financial::Stream->new(
        $model,
        signAdjustment   => '-1*',
        databaseName     => 'cost of sales items',
        flowName         => 'cost of sales (£)',
        balanceName      => 'cost of sales creditors (£)',
        bufferName       => 'cost of sales creditor cash buffer (£)',
        inputTableNumber => 1440,
        numLines         => $model->{numCostSales},
    );

    my $adminExp = Financial::Stream->new(
        $model,
        signAdjustment   => '-1*',
        databaseName     => 'administrative expense items',
        flowName         => 'administrative expenses (£)',
        balanceName      => 'administrative expense creditors (£)',
        bufferName       => 'administrative expense creditor cash buffer (£)',
        inputTableNumber => 1442,
        numLines         => $model->{numAdminExp},
    );

    my $assets = Financial::FixedAssetsUK->new($model);

    my $debt = Financial::Debt->new($model);

    my $cashCalc = Financial::CashCalc->new(
        model     => $model,
        sales     => $sales,
        costSales => $costSales,
        adminExp  => $adminExp,
        assets    => $assets,
        debt      => $debt,
    );

    my $income = Financial::Income->new(
        model     => $model,
        sales     => $sales,
        costSales => $costSales,
        adminExp  => $adminExp,
        assets    => $assets,
        debt      => $debt,
    );

    my $balanceFrictionless = Financial::Balance->new(
        model     => $model,
        sales     => $sales,
        costSales => $costSales,
        adminExp  => $adminExp,
        assets    => $assets,
        cashCalc  => $cashCalc,
        debt      => $debt,
        suffix    => ' assuming frictionless equity',
    );

    my $cashflowFrictionless = Financial::Cashflow->new(
        model   => $model,
        income  => $income,
        balance => $balanceFrictionless,
        suffix  => ' assuming frictionless equity',
    );

    $model->{numYears}   ||= 2;
    $model->{startMonth} ||= 7;
    $model->{startYear}  ||= 2015;

    my $years = Financial::Periods->new(
        model => $model,
        1   ? ( numYears    => $model->{numYears} )
        : 0 ? ( numQuarters => 4 * $model->{numYears} )
        : ( numMonths => 12 * $model->{numYears} ),
        periodsAreFixed     => 1,
        periodsAreInputData => 0,
        priorPeriod         => 1,
        reverseTime         => 0,
        startMonth          => $model->{startMonth},
        startYear           => $model->{startYear},
    );

    my $months = Financial::Periods->new(
        model           => $model,
        numMonths       => 12 * $model->{numYears},
        periodsAreFixed => 1,
        priorPeriod     => 1,
        startMonth      => $model->{startMonth},
        startYear       => $model->{startYear},
        suffix          => 'monthly',
    );

    my $reserve = Financial::Reserve->new(
        model    => $model,
        cashflow => $cashflowFrictionless,
        periods  => $months,
    );

    my $balance = Financial::Balance->new(
        model     => $model,
        sales     => $sales,
        costSales => $costSales,
        adminExp  => $adminExp,
        assets    => $assets,
        cashCalc  => $cashCalc,
        reserve   => $reserve,
        debt      => $debt,
    );

    my $cashflow = Financial::Cashflow->new(
        model   => $model,
        income  => $income,
        balance => $balance,
    );

    my $ratios = Financial::Ratios->new(
        model    => $model,
        income   => $income,
        balance  => $balance,
        cashflow => $cashflow,
    );

    push @{ $model->{incomeTables} }, $income->statement($years);

    push @{ $model->{equityRaisingTables} }, $reserve->raisingSchedule;

    push @{ $model->{balanceTables} },
      $balance->statement($years), $balance->fixedAssetAnalysis($years),
      $balance->workingCapital($years);

    push @{ $model->{cashflowTables} },
      $cashflow->statement($years),
      $cashflow->workingCapital($years), $cashflow->investors($years);

    push @{ $model->{ratioTables} }, $ratios->statement($years);

    $_->finish
      foreach grep { $_->can('finish'); } $months, $years, $assets, $sales,
      $costSales,           $adminExp,             $debt,    $income,
      $balanceFrictionless, $cashflowFrictionless, $reserve, $balance,
      $cashflow,            $ratios;

    $model;

}

1;

