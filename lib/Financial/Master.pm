package Financial;

=head Copyright licence and disclaimer

Copyright 2015, 2016 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Shortcuts ':all';
use Financial::Sheets;

sub requiredModulesForRuleset {
    my ( $module, $ruleset ) = @_;
    map { __PACKAGE__ . '::' . $_; } qw(
      Balance
      CashCalc
      Cashflow
      Debt
      FixedAssetsUK
      FlowAnnual
      Income
      Periods
      Ratios
      Reserve
      ), $ruleset->{numExceptional}
      || $ruleset->{numCapitalExp} ? qw(FlowOnce) : ();
}

sub AUTOLOAD {
    my $model = shift;
    our $AUTOLOAD;
    eval "require $AUTOLOAD";
    my $obj = $AUTOLOAD->new( model => $model, @_ );
    push @{ $model->{finishList} }, $obj;
    $obj;
}

sub new {

    my $class = shift;

    my $model = bless { inputTables => [], @_, }, $class;

    my $sales = $model->FlowAnnual(
        lines        => $model->{numSales},
        name         => 'Sales',
        number       => 1430,
        show_balance => 'debtors (£)',
        show_buffer  => 'debtors cash buffer (£)',
        show_flow    => 'sales (£)',
    );

    my $costSales = $model->FlowAnnual(
        is_cost      => 1,
        lines        => $model->{numCostSales},
        name         => 'Cost of sales',
        number       => 1440,
        show_balance => 'cost of sales creditors (£)',
        show_buffer  => 'cost of sales creditor cash buffer (£)',
        show_flow    => 'cost of sales (£)',
    );

    my $adminExp = $model->FlowAnnual(
        is_cost      => 1,
        lines        => $model->{numAdminExp},
        name         => 'Administrative expenses',
        number       => 1442,
        show_balance => 'administrative expense creditors (£)',
        show_buffer  => 'administrative expense creditor cash buffer (£)',
        show_flow    => 'administrative expenses (£)',
    );

    my $exceptional;
    if ( $model->{numExceptional} ) {
        $exceptional = $model->FlowOnce(
            is_cost      => 1,
            lines        => $model->{numExceptional},
            name         => 'Exceptional costs',
            number       => 1444,
            show_balance => 'exceptional cost creditors (£)',
            show_buffer  => 'exceptional cost creditor cash buffer (£)',
            show_flow    => 'exceptional cost (£)',
        );
    }

    my $capitalExp;
    if ( $model->{numCapitalExp} ) {
        $capitalExp = $model->FlowOnce(
            is_cost      => 1,
            lines        => $model->{numCapitalExp},
            name         => 'Capital expenditure',
            number       => 1447,
            show_balance => 'capital expenditure creditors (£)',
            show_buffer  => 'capital expenditure cash buffer (£)',
            show_flow    => 'capital expenditure (£)',
        );
    }

    my $assets = $model->FixedAssetsUK( capitalExp => $capitalExp, );

    my $debt = $model->Debt;

    my @expensesForCreditors = grep { $_ } $costSales, $adminExp,
      $exceptional, $capitalExp;

    my $cashCalc = $model->CashCalc(
        sales    => $sales,
        expenses => \@expensesForCreditors,
        assets   => $assets,
        debt     => $debt,
    );

    my $income = $model->Income(
        sales     => $sales,
        costSales => $costSales,
        expenses  => [ grep { $_ } $adminExp, $exceptional, ],
        assets    => $assets,
        debt      => $debt,
    );

    my $balanceFrictionless = $model->Balance(
        sales    => $sales,
        expenses => \@expensesForCreditors,
        assets   => $assets,
        cashCalc => $cashCalc,
        debt     => $debt,
        suffix   => ' assuming frictionless equity',
    );

    my $cashflowFrictionless = $model->Cashflow(
        income  => $income,
        balance => $balanceFrictionless,
        suffix  => ' assuming frictionless equity',
    );

    $model->{numYears}   ||= 2;
    $model->{startMonth} ||= 7;
    $model->{startYear}  ||= 2015;

    my $years = $model->Periods(
        numYears            => $model->{numYears},
        periodsAreFixed     => 1,
        periodsAreInputData => 0,
        priorPeriod         => 1,
        reverseTime         => 0,
        startMonth          => $model->{startMonth},
        startYear           => $model->{startYear},
    );

    my $months = $model->Periods(
        $model->{quarterly}
        ? ( numQuarters => 4 * $model->{numYears} )
        : ( numMonths => 12 * $model->{numYears} ),
        periodsAreFixed => 1,
        priorPeriod     => 1,
        startMonth      => $model->{startMonth},
        startYear       => $model->{startYear},
        suffix          => 'monthly',
    );

    my $reserve = $model->Reserve(
        cashflow => $cashflowFrictionless,
        periods  => $months,
    );

    my $balance = $model->Balance(
        sales    => $sales,
        expenses => \@expensesForCreditors,
        assets   => $assets,
        cashCalc => $cashCalc,
        reserve  => $reserve,
        debt     => $debt,
    );

    my $cashflow = $model->Cashflow(
        income  => $income,
        balance => $balance,
    );

    my $ratios = $model->Ratios(
        income   => $income,
        balance  => $balance,
        cashflow => $cashflow,
    );

    push @{ $model->{incomeTables} }, $income->statement($years);

    push @{ $model->{equityRaisingTables} }, $reserve->raisingSchedule;

    push @{ $model->{balanceTables} },
      $balance->statement($years),
      $balance->fixedAssetAnalysis($years),
      $balance->workingCapital($years);

    push @{ $model->{cashflowTables} },
      $cashflow->statement($years),
      $cashflow->profitAndLossReserveMovements($years),
      $cashflow->workingCapitalMovements($years),
      $cashflow->equityInitialAndRaised($years);

    push @{ $model->{ratioTables} },
      $ratios->statement($years),
      $ratios->reference($years),
      $ratios->chart_ebitda_cover($years);

    push @{ $model->{inputCharts} },
      $cashflow->chart_equity_dividends($years),
      $ratios->chart_gearing($years);

    push @{ $model->{standaloneCharts} }, $income->chart($years);

    $_->finish
      foreach grep { UNIVERSAL::can( $_, 'finish' ); }
      @{ $model->{finishList} };

    $model;

}

1;

