﻿package Starter;

# Copyright 2020-2021 Franck Latrémolière and others.
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
use SpreadsheetModel::Book::FrontSheet;

sub serviceMapForRuleset {
    my ($ruleset) = @_;
    {
        fruitCounter => __PACKAGE__ . '::FruitCounter',
        chartTester  => __PACKAGE__ . '::WaterfallTester',
        runk         => __PACKAGE__ . '::Runk',
    };
}

sub myFruitCounter {
    my ($model) = @_;
    $model->{myFruitCounter} //= $model->instantiate( fruitCounter => );
}

sub myChartTester {
    my ($model) = @_;
    $model->{myChartTester} //=
      $model->instantiate(
        chartTester => chartOptions => { mergeFirstStep => 1 } );
}

sub inputsSheetWriter {
    my ( $model, $wbook ) = @_;
    sub {
        my ($wsheet) = @_;
        $model->{inputSheet}{$wbook} = $wsheet;
        $model->{_input_sheet} = $wsheet;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Inputs and charts' ),
          $model->myFruitCounter->inputTables,
          $model->myChartTester->inputTables;
    };
}

sub resultSheetWriter {
    my ( $model, $wbook ) = @_;
    sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Calculations and results' ),
          $model->myFruitCounter->resultTables,
          $model->myChartTester->calculationTables,
          $model->instantiate('runk')->resultTables;
    };
}

sub worksheetsAndClosures {
    my ( $model, $wbook ) = @_;
    $wbook->{titleWriter} = sub { push @{ $model->{_titwrt}{$wbook} }, [@_]; };
    (
        'Inputs+Charts' => $model->inputsSheetWriter($wbook),
        'Calculations'  => $model->resultSheetWriter($wbook),
        'Index'         => SpreadsheetModel::Book::FrontSheet->new(
            model     => $model,
            copyright => 'Copyright 2020-2021 Franck Latrémolière and others.'
        )->closure($wbook),
    );
}

sub finishModel {
    my ( $model, $wbook ) = @_;
    my $append;
    foreach ( @{ $model->{_titwrt}{$wbook} } ) {
        my ( $wsheet, $row, $col, $title, $fmt ) = @$_;
        $append //= $model->myFruitCounter->appendCode( $wbook, $wsheet );
        $wsheet->write( $row, $col, qq%="$title"$append%, $fmt,
                'Not calculated: '
              . 'open in spreadsheet app and allow calculations' );
    }
    $_->wsWrite( $wbook, $model->{inputSheet}{$wbook} )
      foreach $model->myChartTester->charts;
}

sub new {
    my $class = shift;
    bless {@_}, $class;
}

sub requiredModulesForRuleset {
    my ( $class, $ruleset ) = @_;
    values %{ serviceMapForRuleset($ruleset) };
}

sub instantiate {
    my ( $model, $service, @options ) = @_;
    $model->{serviceMap} //= serviceMapForRuleset($model);
    $model->{serviceMap}{$service}->new( $model, @options );
}

1;
