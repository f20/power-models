﻿package ModelM;

# Copyright 2011 The Competitive Networks Association and others.
# Copyright 2012-2017 Franck Latrémolière, Reckon LLP and others.
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

sub finishModel {
    my ( $model, $wbook ) = @_;
    delete $wbook->{titleWriter};
    my $append = ( $model->{idAppend}{$wbook} || '' )
      . ( $model->{checksumAppend}{$wbook} || '' );
    foreach ( @{ $model->{titleWrites}{$wbook} } ) {
        my ( $ws, $row, $col, $n, $fmt ) = @$_;
        $ws->write( $row, $col, qq%="$n"$append%, $fmt,
                'Not calculated: '
              . 'open in spreadsheet app and allow calculations' );
    }
}

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
        my $dataSheet;
        $dataSheet = delete $wbook->{dataSheet} if $model->{noSingleInputSheet};
        $model->{titleWrites}{$wbook} = [];
        $wbook->{titleWriter} =
          sub { push @{ $model->{titleWrites}{$wbook} }, [@_]; };
        $model->{objects}{inputTables} ||= [];
        my $idTable = Dataset(
            number        => 1300,
            dataset       => $model->{dataset},
            name          => 'Company, charging year, data version',
            cols          => Labelset( list => [qw(Company Year Version)] ),
            defaultFormat => 'puretexthard',
            data          => [ 'no company', 'no year', 'no data version' ],
            usePlaceholderData => 1,
            forwardLinks       => {},
        );
        Notes(
            $model->{noSingleInputSheet}
            ? ( name => 'Method M input data' )
            : ( lines =>
                  [ 'Input data', '', 'This sheet contains the input data.' ] )
        )->wsWrite( $wbook, $wsheet );
        require Spreadsheet::WriteExcel::Utility;
        my ( $sh, $ro, $co ) = $idTable->wsWrite( $wbook, $wsheet );
        $sh = $sh->get_name;
        my @cells = map {
            Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro,
                $co + $_ );
        } 0 .. 2;
        $model->{idAppend}{$wbook} =
            qq%&" for "&'$sh'!$cells[0]&" in "&'$sh'!$cells[1]%
          . qq%&" ("&'$sh'!$cells[2]&")"%;

        if ( $model->{noSingleInputSheet} ) {
            $_->wsWrite( $wbook, $wsheet )
              foreach
              sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
              @{ $model->{objects}{inputTables} };
        }
        else {
            $_->wsWrite( $wbook, $wsheet )
              foreach
              sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
              @{ $model->{objects}{inputTables} };
            $model->{multiModelSharing}
              ->addModelIdentificationCells( map { "'$sh'!$cells[$_]"; }
                  0 .. 2 )
              if $model->{multiModelSharing} && !$wbook->{findForwardLinks};
        }
        $wbook->{dataSheet} = $dataSheet if defined $dataSheet;
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
                    ? 'Method M calculations'
                    : "Method M calculations ($sheetName)"
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
          if $model->{multiModelSharing} && !$wbook->{findForwardLinks};
    };

    return $model->{multiModelSharing}
      ->worksheetsAndClosuresWithController( $model, $wbook, @pairs )
      if $model->{multiModelSharing};

    @pairs,
      'Index' => SpreadsheetModel::Book::FrontSheet->new(
        model     => $model,
        copyright => 'Copyright 2009-2012 The Competitive Networks'
          . ' Association and others. '
          . 'Copyright 2012-2020 Franck Latrémolière, Reckon LLP and others.',
      )->closure($wbook);

}

sub licenceNotes {
    Notes(
        name  => '',
        lines => <<'EOL',
Copyright 2009-2012 The Competitive Networks Association and others.  Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.
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
