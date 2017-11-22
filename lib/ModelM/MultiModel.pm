package ModelM::MultiModel;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2014-2017 Franck Latrémolière and others.

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
require Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my ( $class, @content, ) = @_;
    bless { @content, controlSheetNotGeneratedYet => 1 }, $class;
}

sub addModelIdentificationCells {
    my ( $mms, @cells ) = @_;
    push @{ $mms->{modelNames} }, $mms->{waterfalls}
      ? qq%=$cells[2]%
      : qq%=$cells[0]&" "&$cells[1]&" "&$cells[2]%;
    $mms->{waterfallIdentificationCells} ||= [ @cells[ 0, 1 ] ];
}

sub addImpactTableSet {
    push @{ $_[0]{impactTableSets} }, $_[1];
}

sub worksheetsAndClosuresWithController {

    my ( $mms, $model, $wbook, @pairs ) = @_;

    return @pairs unless delete $mms->{controlSheetNotGeneratedYet};

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
              . 'Copyright 2012-2017 Franck Latrémolière, Reckon LLP and others.',
        );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Controller' ), $noticeMaker->extraNotes,
          $noticeMaker->dataNotes,
          $noticeMaker->licenceNotes,
          @{ $mms->{commonAllocationRules} };

        push @{ $mms->{finishClosures} }, sub {
            my @modelOrder =
              $mms->{waterfalls}
              ? ( 0, 2 .. $#{ $mms->{modelNames} }, 1 )
              : ( 0 .. $#{ $mms->{modelNames} } );
            my $modelNameset =
              Labelset( list => [ @{ $mms->{modelNames} }[@modelOrder] ] );

            my @summaryTables = map {
                my $tableNo    = $_;
                my $leadTable  = $mms->{impactTableSets}[0][$tableNo];
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
                            } @{ $mms->{impactTableSets} }[@modelOrder]
                        ]
                    );
                } 0 .. $lastRow;
              } grep { $mms->{impactTableSets}[0][$_]{columns}; }
              0 .. $#{ $mms->{impactTableSets}[0] };

            delete $wbook->{logger};
            $_->wsWrite( $wbook, $wsheet ) foreach @summaryTables;
            if ( $mms->{waterfalls} )
            {    # Useful charts should perhaps be standalone as it is
                    #  difficult to copy charts embedded in locked worksheets.
                my ( $t, $c ) = $mms->waterfallCharts( '', @summaryTables );
                $_->wsWrite( $wbook, $wsheet ) foreach @$t;
                $_->wsWrite( $wbook, $mms->{sheetForCharts} ) foreach @$c;
            }

        };

    };

    unshift @pairs, 'Charts$' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 14;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   6 );
        $wsheet->set_column( 1, 250, 20 );
        $mms->{sheetForCharts} = $wsheet;
        unshift @{ $mms->{finishClosures} }, sub {
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name => $mms->{waterfallIdentificationCells}
                ? '="Waterfall charts for "&'
                  . $mms->{waterfallIdentificationCells}[0]
                  . '&" ("&'
                  . $mms->{waterfallIdentificationCells}[1] . '&")"'
                : 'Waterfall charts'
            );
        };
      }
      if $mms->{waterfalls} && !$mms->{waterfalls} !~ /standalone/i;

    @pairs;

}

sub finishMultiModelSharing {
    my ($mms) = @_;
    return unless $mms->{finishClosures};
    $_->() foreach @{ $mms->{finishClosures} };
}

1;
