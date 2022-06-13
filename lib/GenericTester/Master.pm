package GenericTester;

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

use strict;
use utf8;
use v5.10.0;
use warnings;
use SpreadsheetModel::Book::FrontSheet;
use SpreadsheetModel::Shortcuts ':all';

my %serviceMap = (
    calcBlockTester => __PACKAGE__ . '::CalcBlockTester',
    checksummer     => __PACKAGE__ . '::Checksummer',
    fruitCounter    => __PACKAGE__ . '::FruitCounter',
    waterfallTester => __PACKAGE__ . '::WaterfallTester',
);

sub requiredModulesForRuleset {
    my ( $class, $ruleset ) = @_;
    grep { defined $_; } map { $serviceMap{ $_->[0] }; } @{ $ruleset->{tests} };
}

sub new {
    my $class = shift;
    my $model = bless {@_}, $class;
    foreach ( @{ $model->{tests} } ) {
        my ( $service, @options ) = @$_;
        push @{ $model->{instances} },
          $serviceMap{$service}->new( $model, @options );
    }
    $model;
}

sub inputsSheetWriter {
    my ( $model, $wbook ) = @_;
    sub {
        my ($wsheet) = @_;
        $model->{inputSheet}{$wbook} = $wsheet;
        $model->{_input_sheet} = $wsheet;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   48 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes(
            name => 'Inputs'
              . (
                ( grep { $_->can('charts') } @{ $model->{instances} } )
                ? ' and charts'
                : ''
              )
          ),
          map { $_->inputTables }
          grep { $_->can('inputTables') } @{ $model->{instances} };
    };
}

sub resultSheetWriter {
    my ( $model, $wbook ) = @_;
    sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 0,   48 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Calculations' ), map { $_->calcTables }
          grep { $_->can('calcTables') } @{ $model->{instances} };
    };
}

sub worksheetsAndClosures {
    my ( $model, $wbook ) = @_;
    $wbook->{titleWriter} =
      sub { push @{ $model->{_titwrt}{$wbook} }, [@_]; };
    (
        'Inputs'
          . (
            ( grep { $_->can('charts') } @{ $model->{instances} } )
            ? ' & Charts'
            : ''
          ) => $model->inputsSheetWriter($wbook),
        'Calcs' => $model->resultSheetWriter($wbook),
        'Index' => SpreadsheetModel::Book::FrontSheet->new(
            model     => $model,
            copyright => 'Copyright 2020-2021 Franck Latrémolière and others.'
        )->closure($wbook),
    );
}

sub finishModel {
    my ( $model, $wbook ) = @_;
    $_->wsWrite( $wbook, $model->{inputSheet}{$wbook} )
      foreach map { $_->charts }
      grep { $_->can('charts') } @{ $model->{instances} };
    my ($appendCodeProvider) =
      grep { $_->can('appendCode') } @{ $model->{instances} };
    my $append;
    foreach ( @{ $model->{_titwrt}{$wbook} } ) {
        my ( $wsheet, $row, $col, $title, $fmt ) = @$_;
        $append //=
            $appendCodeProvider
          ? $appendCodeProvider->appendCode( $wbook, $wsheet )
          : '';
        $wsheet->write( $row, $col, qq%="$title"$append%, $fmt,
                'Not calculated: '
              . 'open in spreadsheet app and allow calculations' );
    }
}

1;
