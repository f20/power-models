package ModelM::MultiModel;

# Copyright 2011 The Competitive Networks Association and others.
# Copyright 2014-2019 Franck Latrémolière and others.
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
require Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my ( $class, @content, ) = @_;
    bless { @content, controlSheetNotGeneratedYet => 1 }, $class;
}

sub addModelIdentificationCells {
    my ( $me, @cells ) = @_;
    push @{ $me->{modelNames} }, $me->{waterfalls}
      ? qq%=$cells[1]%
      : qq%=$cells[0]&" "&$cells[1]&" "&$cells[2]%;
    $me->{waterfallIdentificationCells} = [ @cells[ 0, 1 ] ];
}

sub addImpactTableSet {
    push @{ $_[0]{impactTableSets} }, $_[1];
}

sub worksheetsAndClosuresWithController {

    my ( $me, $model, $wbook, @pairs ) = @_;

    return @pairs unless delete $me->{controlSheetNotGeneratedYet};

    unshift @pairs, 'Control$' => sub {

        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 14;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   60 );
        $wsheet->set_column( 1, 250, 20 );
        require SpreadsheetModel::Book::FrontSheet;
        my $noticeMaker = SpreadsheetModel::Book::FrontSheet->new(
            model     => $model,
            copyright => 'Copyright 2009-2012 The Competitive Networks'
              . ' Association and others. '
              . 'Copyright 2012-2022 Franck Latrémolière and others.',
        );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Controller' ), $noticeMaker->extraNotes,
          $noticeMaker->dataNotes,
          $noticeMaker->licenceNotes,
          @{ $me->{commonAllocationRules} };

        push @{ $me->{finishClosures} }, sub {
            my @modelOrder =
              $me->{waterfalls}
              ? ( 0, 2 .. $#{ $me->{modelNames} }, 1 )
              : ( 0 .. $#{ $me->{modelNames} } );
            my $modelNameset =
              Labelset( list => [ @{ $me->{modelNames} }[@modelOrder] ] );

            my @summaryTables = map {
                my $tableNo    = $_;
                my $leadTable  = $me->{impactTableSets}[0][$tableNo];
                my $leadColumn = $leadTable->{columns}[0];
                my ( $sh, $ro, $co ) = $leadColumn->wsWrite( $wbook, $wsheet );
                $sh = $sh->get_name;
                my $colset = Labelset(
                    list => [
                        map {
                            $leadTable->{columns}[$_]{name} =~ /checksum/i
                              ? ()
                              : $leadTable->{columns}[$_]->objectShortName;
                        } 0 .. $#{ $leadTable->{columns} }
                    ]
                );
                my $defaultFormat =
                    $leadTable->{columns}
                  ? $leadTable->{columns}[0]{defaultFormat}
                  : $leadTable->{defaultFormat};
                $defaultFormat =~ s/(con|soft)/copy/
                  unless $defaultFormat =~ /pm$/;
                my $lastRow =
                  $leadColumn->{rows} ? $#{ $leadColumn->{rows}{list} } : 0;
                map {
                    my $row = $_;
                    Constant(
                        name => "From $leadTable->{name}"
                          . (
                            $lastRow ? " $leadColumn->{rows}{list}[$row]" : ''
                          ),
                        defaultFormat => $defaultFormat,
                        rows          => $modelNameset,
                        cols          => $colset,
                        byrow         => 1,
                        data          => [
                            map {
                                my $table = $_->[$tableNo];
                                my ( $sh, $ro, $co ) =
                                  $table->{columns}[0]
                                  ->wsWrite( $wbook, $wsheet );
                                $sh = $sh->get_name;
                                [
                                    map {
                                        qq%='$sh'!%
                                          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                            $ro + $row, $co + $_ );
                                    } 0 .. $#{ $colset->{list} }
                                ];
                            } @{ $me->{impactTableSets} }[@modelOrder]
                        ]
                    );
                } 0 .. $lastRow;
            } grep { $me->{impactTableSets}[0][$_]{columns}; }
              0 .. $#{ $me->{impactTableSets}[0] };

            delete $wbook->{logger};
            $_->wsWrite( $wbook, $wsheet ) foreach @summaryTables;
            if ( $me->{waterfalls} ) {
                my ( $t, $c ) = $me->waterfallCharts( '', @summaryTables );
                $_->wsWrite( $wbook, $wsheet ) foreach @$t;
                $_->wsWrite( $wbook, $me->{sheetForCharts} ) foreach @$c;
            }

        };

    };

    unshift @pairs, 'Charts$' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 14;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0, 86 * ( $me->{scaling_factor} || 1 ) );
        $me->{sheetForCharts} = $wsheet;
        unshift @{ $me->{finishClosures} }, sub {
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name => $me->{waterfallIdentificationCells}
                ? '="Waterfall charts for "&'
                  . $me->{waterfallIdentificationCells}[0]
                  . '&" ("&'
                  . $me->{waterfallIdentificationCells}[1] . '&")"'
                : 'Waterfall charts'
            );
        };
      }
      if $me->{waterfalls} && !$me->{waterfalls} !~ /standalone/i;

    @pairs;

}

sub finishMultiModelSharing {
    my ($me) = @_;
    return unless $me->{finishClosures};
    $_->() foreach @{ $me->{finishClosures} };
}

1;
