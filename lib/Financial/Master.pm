package Financial;

=head Copyright licence and disclaimer

Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Chart;
use Financial::Sheets;

sub requiredModulesForRuleset {
    my ( $module, $ruleset ) = @_;
    map { __PACKAGE__ . '::' . $_; } qw(
      Balance
      CashCalc
      Cashflow
      Debt
      FlowAnnual
      Income
      Periods
      Reserve
      ), $ruleset->{numExceptional}
      || $ruleset->{numCapitalExp} ? qw(FlowOnce) : (),
      $ruleset->{inputDataModule}
      || 'Inputs',      $ruleset->{fixedAssetsModule}
      || 'FixedAssets', $ruleset->{noEquityReturns} ? () : qw(EquityReturns),
      $ruleset->{noRatios} ? () : qw(Ratios);
}

sub AUTOLOAD {
    my ( $model, @arguments ) = @_;
    our $AUTOLOAD;
    my $module = $AUTOLOAD;
    eval "require $module";
    sub {
        return unless @_;
        my (@obj) = $module->new( model => $model, @arguments, @_ );
        push @{ $model->{finishList} }, @obj;
        wantarray ? @obj : $obj[0];
    };
}

sub new {

    my $class           = shift;
    my $model           = bless { inputTables => [], @_, }, $class;
    my $inputDataModule = $model->{inputDataModule} || 'Inputs';
    my $input =
      $model->$inputDataModule->( inputTables => $model->{inputTables} );

    my $sales = $model->FlowAnnual(
        show_balance => 'trade receivables (£)',
        show_buffer  => 'trade receivables cash buffer (£)',
        show_flow    => 'sales (£)',
    )->( $input->sales );

    my $costSales = $model->FlowAnnual(
        is_cost      => 1,
        show_balance => 'cost of sales trade payables (£)',
        show_buffer  => 'cost of sales trade payables cash buffer (£)',
        show_flow    => 'cost of sales (£)',
    )->( $input->costSales );

    my $adminExp = $model->FlowAnnual(
        is_cost      => 1,
        show_balance => 'administrative expense trade payables (£)',
        show_buffer  => 'administrative expense trade payables cash buffer (£)',
        show_flow    => 'administrative expenses (£)',
    )->( $input->adminExp );

    my $exceptional = $model->FlowOnce(
        is_cost      => 1,
        show_balance => 'exceptional cost trade payables (£)',
        show_buffer  => 'exceptional cost trade payables cash buffer (£)',
        show_flow    => 'exceptional cost (£)',
    )->( $input->exceptional );

    my $capitalExp = $model->FlowOnce(
        is_cost      => 1,
        show_balance => 'capital expenditure trade payables (£)',
        show_buffer  => 'capital expenditure cash buffer (£)',
        show_flow    => 'capital expenditure (£)',
    )->( $input->capitalExp );

    my $fixedAssetsModule = $model->{fixedAssetsModule} || 'FixedAssets';
    my $assets =
      $model->$fixedAssetsModule( capitalExp => $capitalExp, )
      ->( $input->assets );

    my $debt = $model->Debt->( $input->debt );

    my @expensesForPayables = grep { $_ } $costSales, $adminExp,
      $exceptional, $capitalExp;

    my $cashCalc = $model->CashCalc->(
        sales    => $sales,
        expenses => \@expensesForPayables,
        assets   => $assets,
        debt     => $debt,
    );

    my $income = $model->Income->(
        sales     => $sales,
        costSales => $costSales,
        expenses  => [ grep { $_ } $adminExp, $exceptional, ],
        assets    => $assets,
        debt      => $debt,
    );

    my $balanceFrictionless = $model->Balance->(
        sales    => $sales,
        expenses => \@expensesForPayables,
        assets   => $assets,
        cashCalc => $cashCalc,
        debt     => $debt,
        suffix   => ' if frictionless equity was available',
    );

    my $cashflowFrictionless = $model->Cashflow->(
        income  => $income,
        balance => $balanceFrictionless,
        suffix  => ' if frictionless equity was available',
    );

    $model->{numYears}   ||= 2;
    $model->{startMonth} ||= 7;
    $model->{startYear}  ||= 2015;

    my $years = $model->Periods->(
        numYears            => $model->{numYears},
        periodsAreFixed     => 1,
        periodsAreInputData => 0,
        priorPeriod         => 1,
        reverseTime         => 0,
        startMonth          => $model->{startMonth},
        startYear           => $model->{startYear},
    );

    my $months = $model->Periods->(
        $model->{quarterly}
        ? ( numQuarters => 4 * $model->{numYears} )
        : ( numMonths => 12 * $model->{numYears} ),
        periodsAreFixed => 1,
        priorPeriod     => 1,
        startMonth      => $model->{startMonth},
        startYear       => $model->{startYear},
        suffix          => $model->{quarterly} ? 'quarterly' : 'monthly',
    );

    my $reserve = $model->Reserve->(
        cashflow => $cashflowFrictionless,
        periods  => $months,
    );

    my $balance = $model->Balance->(
        sales    => $sales,
        expenses => \@expensesForPayables,
        assets   => $assets,
        cashCalc => $cashCalc,
        reserve  => $reserve,
        debt     => $debt,
    );

    my $cashflow = $model->Cashflow->(
        income  => $income,
        balance => $balance,
    );

    push @{ $model->{incomeTables} }, $income->statement($years);

    push @{ $model->{monthlyTables} }, $reserve->raisingSchedule,
      $reserve->openingSpareCash, $reserve->{periods}->openingDay;

    push @{ $model->{balanceTables} },
      $balance->statement($years),
      $balance->fixedAssetAnalysis($years),
      $balance->workingCapital($years);

    push @{ $model->{cashflowTables} },
      $cashflow->statement($years),
      $cashflow->retainedEarningsMovements($years),
      $cashflow->workingCapitalMovements($years);

    unless ( $model->{noEquityReturns} ) {
        my $equityReturns = $model->EquityReturns->(
            income   => $income,
            balance  => $balance,
            cashflow => $cashflow,
        );
        push @{ $model->{cashflowTables} },
          @{ $equityReturns->equityInternalRateOfReturn($years) },
          $equityReturns->npv($years);
        push @{ $model->{standaloneCharts} },
          $equityReturns->chart_equity_dividends($years),
          $equityReturns->chart_npv($years);
    }

    unless ( $model->{noRatios} ) {
        my $ratios = $model->Ratios->(
            income   => $income,
            balance  => $balance,
            cashflow => $cashflow,
        );
        push @{ $model->{ratioTables} },
          $ratios->statement($years),          $ratios->reference($years),
          $ratios->chart_ebitda_cover($years), $ratios->chart_roce($years);
        push @{ $model->{inputCharts} }, $ratios->chart_gearing($years);
    }

    push @{ $model->{standaloneCharts} }, $income->chart($years);

    $_->finish
      foreach grep { UNIVERSAL::can( $_, 'finish' ); }
      @{ $model->{finishList} };

    $model;

}

1;

