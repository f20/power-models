package ElecHarness::Extension;

# Copyright 2021 Franck Latrémolière and others.
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
use Spreadsheet::WriteExcel::Utility;

sub process {
    my ( $self, $maker, $options ) = @_;
    $maker->{setting}->( customPicker => \&_customPicker );
}

sub _customPicker {

    my ( $addToList,      $datasetsRef,  $rulesetsRef )   = @_;
    my ( $datasetHarness, $datasetAfter, $datasetBefore ) = @$datasetsRef;
    my ( $rulesetHarness, $rulesetAfter, $rulesetBefore ) = @$rulesetsRef;
    my $datasetName = $datasetHarness->{'~datasetName'};
    $_->{'~datasetName'} = $datasetName foreach $datasetAfter, $datasetBefore;
    $_->{'template'} = '%' foreach @$rulesetsRef;
    $rulesetBefore->{'nickName'} ||= 'Start';
    $rulesetAfter->{'nickName'}  ||= 'Finish';
    my %sourceModelsIds = (
        '_finish' => ( $datasetAfter->{'~datasetId'} = "_finish $datasetName" ),
        '_start' => ( $datasetBefore->{'~datasetId'} = "_start $datasetName" ),
    );
    $addToList->( $datasetHarness, $rulesetHarness );
    $addToList->( $datasetAfter,   $rulesetAfter );
    $addToList->( $datasetBefore,  $rulesetBefore );

    my $tableSourcesMap;
    my $hybridDatasetFactory = sub {
        my ( $step, $nick ) = @_;
        my %dataset;
        $dataset{sourceModelsIds} = \%sourceModelsIds;
        $dataset{1500}[3] =
          { 'Company charging year data version' => $nick || "Step $step", };
        $dataset{datasetCallback} = sub {
            my ($model) = @_;
            $tableSourcesMap ||= _buildTableSourcesMap($model);
            $model->{dataset}{$_} =
              _tableClosureFactory( $_, $tableSourcesMap->{$_}, $step )
              foreach keys %$tableSourcesMap;
        };
        \%dataset;
    };

    my $ruleset = $rulesetBefore;
    foreach ( 0 .. $#{ $rulesetHarness->{ruleChanges} } ) {
        $ruleset = {%$ruleset};
        while ( my ( $k, $v ) = each %{ $rulesetHarness->{ruleChanges}[$_] } ) {
            $ruleset->{$k} = $v;
        }
        $addToList->(
            {
                '~datasetName' => $datasetName,
                dataset =>
                  $hybridDatasetFactory->( $_ + 1, $ruleset->{nickName} ),
            },
            $ruleset
        );
    }

}

sub _tableClosureFactory {
    my ( $table, $srcs, $step ) = @_;
    sub {
        my ( $table, $wb, $ws ) = @_;
        my $get_coord = sub {
            my ($obj) = @_;
            return ( undef, undef, undef, [], 0 ) unless defined $obj;
            my $o2 = $obj->{columns} ? $obj->{columns}[0] : $obj;
            my ( $sh, $ro, $co ) = $o2->wsWrite( $wb, $ws );
            my @rowLabels = !$o2->lastRow ? 'Single row' : map {
                local $_ = $_;
                s/.*\n//s;
                s/[^A-Za-z0-9 -]/ /g;
                s/- / /g;
                s/ +/ /g;
                s/^ //;
                s/ $//;
                $_;
            } @{ $o2->{rows}{list} };
            my $colCount = 0;
            $colCount += 1 + $_->lastCol
              foreach $obj->{columns} ? @{ $obj->{columns} } : $obj;
            ( "'" . $sh->get_name . "'!", $ro, $co, \@rowLabels, $colCount );
        };
        my ( $pre1, $row1, $col1, $lab1, $cc1 ) =
          $get_coord->( $srcs->{start} );
        my ( $pre2, $row2, $col2, $lab2, $cc2 ) =
          $get_coord->( $srcs->{finish} );
        my ( $pre3, $row3, $col3, $lab3, $cc3 ) =
          $get_coord->( $srcs->{stepControl} );
        my $colCount = $cc1;
        $colCount = $cc2 if $colCount < $cc2;
        [
            undef,
            map {
                my $c = $_ - 1;
                my %d;
                for ( my $r = 0 ; $r < @$lab3 ; ++$r ) {
                    my $cell1 = $pre1
                      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $row1 + $r, $col1 + $c );
                    my $cell2 = $pre2
                      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $row2 + $r, $col2 + $c );
                    my $cell3 = $pre3
                      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $row3 + $r, $col3 + $c );
                    $d{ $lab3->[$r] } = "=IF($step<$cell3,$cell1,$cell2)";
                }
                for ( my $r = 0 ; $r < @$lab2 ; ++$r ) {
                    $d{ $lab2->[$r] } ||= '='
                      . $pre2
                      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $row2 + $r, $col2 + $c );
                }
                for ( my $r = 0 ; $r < @$lab1 ; ++$r ) {
                    $d{ $lab1->[$r] } ||= '='
                      . $pre1
                      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $row1 + $r, $col1 + $c );
                }
                \%d;
            } 1 .. $colCount
        ];
    };
}

sub _buildTableSourcesMap {
    my ($model) = @_;
    my %map;
    foreach ( @{ $model->{sourceModels}{_start}{inputTables} } ) {
        $map{ $_->{number} }{start} = $_;
    }
    foreach ( @{ $model->{sourceModels}{_finish}{inputTables} } ) {
        $map{ $_->{number} }{finish} = $_;
    }
    foreach my $table ( sort keys %map ) {
        my ( $start, $finish ) = @{ $map{$table} }{qw(start finish)};
        next unless $start && $finish;

        my $name1 = $start->objectShortName;
        my $name2 = $finish->objectShortName;
        my $name  = $name1 eq $name2 ? $name1 : "$name1 or $name2";

        my @rows1 =
          $start->{rows} ? @{ $start->{rows}{list} } : $start->objectShortName;
        my @rows2 =
          $finish->{rows}
          ? @{ $finish->{rows}{list} }
          : $finish->objectShortName;
        my @rows;
        if ( @rows1 == 1 && @rows2 == 1 ) {
            @rows =
              $rows1[0] eq $rows2[0] ? $rows1[0] : "$rows1[0] or $rows2[0]";
        }
        else {
            my %rows;
            foreach ( @rows1, @rows2 ) {
                next if exists $rows{$_};
                push @rows, $_;
                undef $rows{$_};
            }
        }

        my ( @cols1, @cols2 );
        push @cols1, $_->{cols} ? @{ $_->{cols}{list} } : $_->objectShortName
          foreach $start->{columns} ? @{ $start->{columns} } : $start;
        push @cols2, $_->{cols} ? @{ $_->{cols}{list} } : $_->objectShortName
          foreach $finish->{columns} ? @{ $finish->{columns} } : $finish;
        my @cols = map {
            $cols1[$_] eq $cols2[$_] ? $cols1[$_] : "$cols1[$_] or $cols2[$_]";
        } 0 .. ( @cols2 > @cols1 ? $#cols2 : $#cols1 );
        push @{ ${ $model->{sharingObjectRef} }->{stepRules} },
          $map{$table}{stepControl} = Dataset(
            name          => "$table. Step definition for $name",
            number        => $table,
            dataset       => ${ $model->{sharingObjectRef} }->{dataset},
            defaultFormat => 'stephard',
            rows          => Labelset(
                list => \@rows,
                $finish->{rows} && $finish->{rows}{defaultFormat}
                ? ( defaultFormat => $finish->{rows}{defaultFormat} )
                : (),
            ),
            cols => Labelset( list => \@cols ),
            data => [
                map {
                    [ map { '2'; } @rows ];
                } @cols
            ],
          );
    }
    \%map;
}

1;
