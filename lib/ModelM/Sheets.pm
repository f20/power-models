package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.

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
use Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::ColourCodeWriter;

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    my @pairs;

    push @pairs, 'Input' => sub {
        my ($wsheet) = @_;

        $wbook->{lastSheetNumber} = 13;
        $wsheet->{sheetNumber}    = 13;

        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   60 );
        $wsheet->set_column( 1, 250, 20 );
        $wsheet->{nextFree} = 2;    # One comment line under "Input data" title
        $model->{objects}{inputTables} ||= [];
        $model->{dataset}{1300}[3]{'Company charging year data version'} =
          $model->{version}
          if $model->{version} && $model->{dataset};
        my ( $sh, $ro, $co ) = Dataset(
            number        => 1300,
            dataset       => $model->{dataset},
            name          => 'Company, charging year, data version',
            cols          => Labelset( list => [qw(Company Year Version)] ),
            defaultFormat => 'puretexthard',
            data          => [ 'no company', 'no year', 'no data version' ],
            usePlaceholderData => 1,
        )->wsWrite( $wbook, $wsheet );
        $sh = $sh->get_name;
        $wbook->{titleAppend} =
            qq%" for "&'$sh'!%
          . xl_rowcol_to_cell( $ro, $co )
          . qq%&" in "&'$sh'!%
          . xl_rowcol_to_cell( $ro, $co + 1 )
          . qq%&" ("&'$sh'!%
          . xl_rowcol_to_cell( $ro, $co + 2 ) . '&")"';
        $model->{multiModelSharing}->addModelName( qq%='$sh'!%
              . xl_rowcol_to_cell( $ro, $co )
              . qq%&" "&'$sh'!%
              . xl_rowcol_to_cell( $ro, $co + 2 ) )
          if $model->{multiModelSharing};
        $_->wsWrite( $wbook, $wsheet )
          foreach sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
          @{ $model->{objects}{inputTables} };
        my $nextFree = delete $wsheet->{nextFree};
        Notes( lines =>
              [ 'Input data', '', 'This sheet contains the input data.' ] )
          ->wsWrite( $wbook, $wsheet );
        $wsheet->{nextFree} = $nextFree;
    };

    my %tablesBySheet;

    foreach (
        $model->{objects}{calcSheets}
        ? @{ $model->{objects}{calcSheets} }
        : ()
      )
    {
        my ( $sheetName, @tables ) = @$_;
        $sheetName ||= 'Calculations';
        if ( $tablesBySheet{$sheetName} ) {
            push @{ $tablesBySheet{$sheetName} }, @tables;
        }
        else {
            $tablesBySheet{$sheetName} = \@tables;
            push @pairs, ( $sheetName || 'Calculations' ) => sub {
                my ($wsheet) = @_;
                $wsheet->freeze_panes( 1, 0 );
                $wsheet->set_column( 0, 0,   36 );
                $wsheet->set_column( 1, 250, 20 );
                $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                    name => $sheetName eq 'Calculations'
                    ? 'Calculations'
                    : "Calculations ($sheetName)"
                  ),
                  @{ $tablesBySheet{$sheetName} };
            };
        }
    }

    push @pairs, 'Results' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   36 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Results' ),
          @{ $model->{objects}{resultsTables} };
        $model->{multiModelSharing}
          ->addImpactTableSet( $model->{objects}{resultsTables} )
          if $model->{multiModelSharing};
    };

    return $model->{multiModelSharing}
      ->worksheetsAndClosuresWithController( $model, $wbook, @pairs )
      if $model->{multiModelSharing};

    @pairs, 'Index' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->fit_to_pages( 1, 2 );
        $wsheet->set_column( 0, 0,   30 );
        $wsheet->set_column( 1, 1,   105 );
        $wsheet->set_column( 2, 250, 30 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->topNotes, $model->licenceNotes,
          SpreadsheetModel::ColourCodeWriter->new,
          $wbook->{logger}, $model->technicalNotes;
    };

}

sub technicalNotes {
    my ($model) = @_;
    require POSIX;
    Notes(
        name       => '',
        rowFormats => ['caption'],
        lines      => [
            'Technical model rules and version control',
            $model->{yaml},
            '',
            'Generated on '
              . POSIX::strftime( '%a %e %b %Y %H:%M:%S',
                $model->{localTime} ? @{ $model->{localTime} } : localtime )
              . ( $ENV{SERVER_NAME} ? " by $ENV{SERVER_NAME}" : '' ),
        ]
    );
}

sub topNotes {
    my ($model) = @_;
    Notes(
        name  => 'Index',
        lines => [
            $model->{colour} && $model->{colour} =~ /orange|gold/ ? <<EOL : (),

This document, model or dataset has been prepared by Reckon LLP on the instructions of the DCUSA Panel or one of its working
groups.  Only the DCUSA Panel and its working groups have authority to approve this material as meeting their requirements.
Reckon LLP makes no representation about the suitability of this material for the purposes of complying with any licence
conditions or furthering any relevant objective.
EOL
            $model->dataNotes,
            $model->{noLinks} ? () : <<EOL,

This workbook is structured as a series of named and numbered tables. There is a list of tables below, with hyperlinks. Above
each calculation table, there is a description of the calculations and hyperlinks to tables from which data are used. Hyperlinks
point to the relevant table column heading of the relevant table. Scrolling up or down is usually required after clicking a
hyperlink in order to bring the relevant data and/or headings into view. Some versions of Microsoft Excel can display a "Back"
button, which can be useful when using hyperlinks to navigate around the workbook.
EOL
        ]
    );
}

sub dataNotes {
    my ($model) = @_;
    $model->{colour} && $model->{colour} =~ /gold/ ? <<EOL :

UNLESS STATED OTHERWISE, ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
      <<EOL;

UNLESS STATED OTHERWISE, THIS WORKBOOK IS ONLY A PROTOTYPE FOR TESTING PURPOSES AND ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
}

sub licenceNotes {
    Notes(
        name  => '',
        lines => <<'EOL',
Copyright 2009-2012 The Competitive Networks Association and others.  Copyright 2012-2015 Franck Latrémolière, Reckon LLP and others.
The code used to generate this spreadsheet includes open-source software published at https://github.com/f20/power-models.
Use and distribution of the source code is subject to the conditions stated therein. 
Any redistribution of this software must retain the following disclaimer:
THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL AUTHORS OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED
AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
EOL
    );
}

1;
