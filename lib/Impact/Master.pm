package Impact;

=head Copyright licence and disclaimer

Copyright 2017 Franck Latrémolière, Reckon LLP and others.

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

sub requiredModulesForRuleset {
    my ( $class, $ruleset ) = @_;
    my $impact = $ruleset->{impact} or die 'No impact specification';
    "Impact::$impact", $ruleset->{byDno}
      ? qw(SpreadsheetModel::Data::DnoAreas
      SpreadsheetModel::MatrixSheet)
      : qw(SpreadsheetModel::ColumnsetFilter);
}

sub worksheetsAndClosures {
    my ($model) = @_;
    @{ $model->{worksheetsAndClosures} };
}

sub new {
    my $class = shift;
    my $model = bless {@_}, $class;
    study $model->{baselinePattern};
    study $model->{scenarioPattern};
    my $method = 'process' . $model->{impact};
    if ( $model->{byDno} ) {
        my $titleSuffix = $model->{title} ? ": $model->{title}" : '';
        my @short       = SpreadsheetModel::Data::DnoAreas::dnoShortNames();
        my @long        = SpreadsheetModel::Data::DnoAreas::dnoLongNames();
        for ( my $i = 0 ; $i < @short ; ++$i ) {
            my ( $baselineData, $scenarioData ) =
              $model->selectDatasets( $short[$i] )
              or next;
            $model->$method( $baselineData, $scenarioData, $short[$i],
                $long[$i] . $titleSuffix,
            );
        }
    }
    else {
        my %areas;
        foreach ( @{ $model->{datasetArray} } ) {
            local $_ = $_->{'~datasetName'};
            s/-20.*//;
            undef $areas{$_};
        }
        $model->$method( $model->selectDatasets($_), $_ )
          foreach sort keys %areas;
        $model->consolidate;
    }
    $model;
}

sub selectDatasets {
    my ( $model, $area ) = @_;
    $area =~ tr/ /-/;
    my ( $baselineData, $scenarioData );
    foreach my $ds ( @{ $model->{datasetArray} } ) {
        local $_ = $ds->{'~datasetName'} or next;
        next unless /$area/;
        if (/$model->{baselinePattern}/) {
            die "Several candidate datasets for $area baseline"
              if $baselineData;
            $baselineData = $ds->{dataset};
        }
        if (/$model->{scenarioPattern}/) {
            die "Several candidate datasets for $area scenario"
              if $scenarioData;
            $scenarioData = $ds->{dataset};
        }
    }
    warn "No dataset for $area baseline" unless $baselineData;
    warn "No dataset for $area scenario" unless $scenarioData;
    return unless $baselineData && $scenarioData;
    $baselineData, $scenarioData;
}

sub consolidate {
    my ($model) = @_;
    my ( $rowset, $areaLabels, $rowLabels, @dataColumns, @calcColumns );
    foreach ( @{ $model->{columnsetFilterFood} } ) {
        my ( $area, @datasets ) = @$_;
        my $rowslist = $datasets[0]{rows}{list};
        unless ($rowset) {
            $rowset = Labelset(
                defaultFormat => 'thitem',
                list          => [ 1 .. @$rowslist ],
            );
            $areaLabels = Constant(
                name          => 'Area',
                defaultFormat => 'th',
                rows          => $rowset,
                data          => [ map { $area; } @$rowslist ],
            );
            $rowLabels = Constant(
                name          => 'Row',
                defaultFormat => 'th',
                rows          => $rowset,
                data          => $rowslist,
            );
            $_->{rows} = $rowset foreach @datasets;
            @dataColumns = grep { $_->{data}; } @datasets;
            @calcColumns = grep { !$_->{data}; } @datasets;
            delete $_->{rowFormats} foreach @calcColumns;
            next;
        }
        push @{ $rowset->{list} },
          ( 1 + @{ $rowset->{list} } ) .. ( @{ $rowset->{list} } + @$rowslist );
        push @{ $areaLabels->{data} }, map { $area; } @$rowslist;
        push @{ $rowLabels->{data} }, @$rowslist;
        @datasets = grep { $_->{data}; } @datasets;
        for ( my $i = 0 ; $i < @datasets ; ++$i ) {
            my $acc = $dataColumns[$i]{data};
            my $add = $datasets[$i]{data};
            if ( ref $acc->[0] ) {
                for ( my $j = 0 ; $j < @$acc ; ++$j ) {
                    push @{ $acc->[$j] }, @{ $add->[$j] };
                }
            }
            else {
                push @$acc, @$add;
            }
        }
    }
    push @{ $model->{worksheetsAndClosures} }, All => sub {
        my ( $wsheet, $wbook ) = @_;
        $wsheet->set_column( 0, 0,   16 );
        $wsheet->set_column( 1, 1,   48 );
        $wsheet->set_column( 2, 254, 16 );
        $wsheet->freeze_panes( 3, 2 );
        $_->wsWrite( $wbook, $wsheet )
          foreach SpreadsheetModel::ColumnsetFilter->new(
            name    => ucfirst( $model->{title} ),
            columns => [ $areaLabels, $rowLabels, @dataColumns, @calcColumns, ],
          );
    };
}

1;
