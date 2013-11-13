package ModelM;

=head Copyright licence and disclaimer

Copyright 2009-2012 The Competitive Networks Association and others.

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

use POSIX ();

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    $wbook->{lastSheetNumber} = 13;

    my @pairs;

    push @pairs, 'Input' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 13;

        # reset in case of building several models in a single workbook
        delete $wbook->{highestAutoTableNumber};
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   36 );
        $wsheet->set_column( 1, 250, 20 );
        return if $model->{oneSheet};
        $wsheet->{nextFree} = 2;    # One comment line under "Input data" title
        $model->{inputTables} ||= [];
        $model->{dataset}{1300}[3]{'Company charging year data version'} =
          $model->{version}
          if $model->{version} && $model->{dataset};
        my $modelInformationTable = Dataset(
            number        => 1300,
            dataset       => $model->{dataset},
            name          => 'Company, charging year, data version',
            cols          => Labelset( list => [qw(Company Year Version)] ),
            defaultFormat => 'texthard',
            data          => [
                'Illustrative company',
                'Illustrative year',
                'Illustrative dataset'
            ]
        );
        my ( $sh, $ro, $co ) =
          $modelInformationTable->wsWrite( $wbook, $wsheet );
        $sh = $sh->get_name;
        $wbook->{titleAppend} =
            qq%" for "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co )
          . qq%&" in "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 1 )
          . qq%&" ("&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 2 )
          . '&")"';
        push @{ $model->{multiModelSharing}{modelNameList} },
            qq%='$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co )
          . qq%&" "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 2 )
          if $model->{multiModelSharing};

        $_->wsWrite( $wbook, $wsheet )
          foreach sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
          @{ $model->{inputTables} };
        my $nextFree = delete $wsheet->{nextFree};
        Notes( lines =>
              [ 'Input data', '', 'This sheet contains the input data.' ] )
          ->wsWrite( $wbook, $wsheet );
        $wsheet->{nextFree} = $nextFree;
    };

    push @pairs, 'Calculations' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   36 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Calculations' ), @{ $model->{calcTables} };
    };

    push @pairs, 'Results' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   36 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Results' ), @{ $model->{impactTables} };
        push @{ $model->{multiModelSharing}{impactTablesColl} },
          $model->{impactTables}
          if $model->{multiModelSharing};
    };

    if ( my $mms = $model->{multiModelSharing} ) {

        return @pairs if $mms->{controller};
        $mms->{controller} = 1;

        unshift @pairs, Controller => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   60 );
            $wsheet->set_column( 1, 250, 20 );
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name => 'Controller',
                lines =>
                  [ $model->illustrativeNotice, $model->disclaimerNotice, ]
              ),
              @{ $mms->{optionsColumns} };

            $mms->{finish} = sub {
                delete $wbook->{logger};
                my $modelNameset = Labelset( list => $mms->{modelNameList} );
                $_->wsWrite( $wbook, $wsheet ) foreach map {
                    my $tableNo    = $_;
                    my $leadTable  = $mms->{impactTablesColl}[0][$tableNo];
                    my $leadColumn = $leadTable->{columns}[0] || $leadTable;
                    my ( $sh, $ro, $co ) =
                      $leadColumn->wsWrite( $wbook, $wsheet );
                    $sh = $sh->get_name;
                    my $colset = Labelset(
                        list => [
                            map {
                                qq%='$sh'!%
                                  . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                    $ro - 1, $co + $_ )
                            } 0 .. $#{ $leadTable->{columns} }
                        ]
                    );
                    my $defaultFormat = $leadTable->{defaultFormat}
                      || $leadTable->{columns}[0]{defaultFormat};
                    $defaultFormat =~ s/soft/copy/
                      unless $defaultFormat =~ /pm$/;
                    Constant(
                        name          => "From $leadTable->{name}",
                        defaultFormat => $defaultFormat,
                        rows          => $modelNameset,
                        cols          => $colset,
                        byrow         => 1,
                        data          => [
                            map {
                                my $table = $_->[$tableNo];
                                my ( $sh, $ro, $co ) = $table->{columns}[0]
                                  ->wsWrite( $wbook, $wsheet );
                                $sh = $sh->get_name;
                                [
                                    map {
                                        qq%='$sh'!%
                                          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                            $ro, $co + $_ );
                                    } 0 .. $#{ $leadTable->{columns} }
                                ];
                            } @{ $mms->{impactTablesColl} }
                        ]
                    );
                } 0 .. $#{ $mms->{impactTablesColl}[0] };
            };

        };

        return @pairs;

    }

    return @pairs, $model->{oneSheet}
      ? ( 'All' => sub { } )
      : (
        'Overview' => sub {

            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->fit_to_pages( 1, 2 );
            $wsheet->set_column( 0, 0,   30 );
            $wsheet->set_column( 1, 1,   90 );
            $wsheet->set_column( 2, 250, 30 );
            $model->generalNotes->wsWrite( $wbook, $wsheet );

            $wsheet->write_string(
                2, 2,
                'Colour coding',
                $wbook->getFormat('thc')
            );
            $wsheet->write_string( 3, 2, 'Data input',
                $wbook->getFormat('0.000hard') );
            $wsheet->write_string(
                4, 2,
                'Unused cell in input data table',
                $wbook->getFormat('unused')
            );
            $wsheet->write_string( 5, 2, 'Calculation',
                $wbook->getFormat('0.000soft') );
            $wsheet->write_string( 6, 2, 'Copy data',
                $wbook->getFormat('0.000copy') );
            $wsheet->write_string(
                7, 2,
                'Unused cell in calculation table',
                $wbook->getFormat('unavailable')
            );
            $wsheet->write_string(
                8, 2,
                'Constant value',
                $wbook->getFormat('0.000con')
            );
            $wsheet->write_string(
                9, 2,
                'Unlocked cell for notes',
                $wbook->getFormat('scribbles')
            );

            $wbook->{logger}->wsWrite( $wbook, $wsheet );

            local $/ = "\n";

            Notes(
                name       => '',
                rowFormats => ['caption'],
                lines      => [
                    'Model identification and configuration',
                    $model->{yaml},
                    '',
                    'Generated on '
                      . POSIX::strftime( '%a %e %b %Y %H:%M:%S',
                        @{ $model->{localTime} } )
                      . ( $ENV{SERVER_NAME} ? " by $ENV{SERVER_NAME}" : '' ),
                ]
            )->wsWrite( $wbook, $wsheet );

        }
      );

}

sub generalNotes {
    my ($model) = @_;
    Notes(
        name  => 'Overview',
        lines => [
            $model->{colour} && $model->{colour} =~ /orange|gold/ ? <<EOL : (),

This document, model or dataset has been prepared by Reckon LLP on the instructions of the DCUSA Panel or
one of its working groups.  Only the DCUSA Panel and its working groups have authority to approve this
material as meeting their requirements.  Reckon LLP makes no representation about the suitability of this
material for the purposes of complying with any licence conditions or furthering any relevant objective.
EOL
            $model->illustrativeNotice,
            $model->{noLinks} ? () : <<EOL,

This workbook is structured as a series of named and numbered tables. There is a list of tables below, with
hyperlinks.  Above each calculation table, there is a description of the calculations made, and a hyperlinked
list of the tables or parts of tables from which data are used in the calculation. Hyperlinks point to the
relevant table column heading of the relevant table. Scrolling up or down is usually required after clicking
a hyperlink in order to bring the relevant data and/or headings into view. Some versions of Microsoft Excel
can display a "Back" button, which can be useful when using hyperlinks to navigate around the workbook.
EOL
            $model->disclaimerNotice,
        ]
    );
}

sub illustrativeNotice {
    my ($model) = @_;
    $model->{colour} && $model->{colour} =~ /gold/ ? <<EOL :

UNLESS STATED OTHERWISE, ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
      <<EOL;

UNLESS STATED OTHERWISE, THIS WORKBOOK IS ONLY A PROTOTYPE FOR TESTING
PURPOSES AND ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
}

sub disclaimerNotice {
    <<'EOL',

Copyright 2009-2012 The Competitive Networks Association and others. Copyright 2012-2013 Franck Latrémolière,
Reckon LLP and others. Redistribution and use in source and binary forms, with or without modification, are
permitted provided that the following conditions are met:
1. Redistributions of source code must retain the above copyright notice, this list of conditions and the
following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the
following disclaimer in the documentation and/or other materials provided with the distribution.
THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES,
INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS
OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY
WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
EOL
}

1;
