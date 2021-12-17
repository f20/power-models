package Sampler;

# Copyright 2015-2021 Franck Latrémolière and others.
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
use SpreadsheetModel::Book::FrontSheet;

sub new {
    my $class = shift;
    bless {@_}, $class;
}

sub requiredModulesForRuleset {
    my ( $class, $model ) = @_;
    'Sampler::Writers', $model->{showColourCode}
      || $model->{showNumFormatColours} ? 'Sampler::ColoursList'  : (),
      $model->{showColourMatrix}        ? 'Sampler::ColoursArray' : (),
      $model->{omitLegend} ? () : 'SpreadsheetModel::Book::FormatLegend';
}

sub worksheetsAndClosures {
    my ( $model, $wbook ) = @_;
    Sampler => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 255, $model->{omitLegend} ? 18 : 24 );
        Notes( name => 'Spreadsheet format sampler' )
          ->wsWrite( $wbook, $wsheet );
        SpreadsheetModel::Book::FormatLegend->new->wsWrite( $wbook, $wsheet )
          unless $model->{omitLegend};
        $model->writeFormatList( $wbook, $wsheet ) unless $model->{omitFormats};
        $model->writeColourCode( $wbook, $wsheet )
          if $model->{showColourCode};
        $model->writeNumFormatColours( $wbook, $wsheet )
          if $model->{showNumFormatColours};
        $model->writeMpanFormats( $wbook, $wsheet )
          unless $model->{omitMpanFormats};
        $model->writeColourMatrix( $wbook, $wsheet )
          if $model->{showColourMatrix};
        SpreadsheetModel::Book::FrontSheet->new(
            model     => $model,
            copyright => 'Copyright 2015-2021 Franck Latrémolière and others.'
        )->technicalNotes->wsWrite( $wbook, $wsheet );
    };
}

1;
