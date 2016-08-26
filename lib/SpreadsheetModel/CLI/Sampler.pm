package SpreadsheetModel::CLI::CommandRunner;

=head Copyright licence and disclaimer

Copyright 2011-2015 Franck Latrémolière and others. All rights reserved.

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

sub sampler {
    my ( $self, @options ) = @_;
    my %options;
    $options{colour} = 'gold'   if grep { /gold/i } @options;
    $options{colour} = 'orange' if grep { /orange/i } @options;
    use SpreadsheetModel::Book::WorkbookXLSX;
    my $wbook = SpreadsheetModel::Book::WorkbookXLSX->new($$);
    $wbook->setFormats( \%options );
    my $wsheet = $wbook->add_worksheet('Sampler');
    $wsheet->set_paper(9);
    $wsheet->fit_to_pages( 1, 0 );
    $wsheet->hide_gridlines(2);
    $wsheet->set_column( 0, 5, 20 );
    my $includeJsonColumn = undef;
    $wsheet->set_column( 6, 6, 120 ) if $includeJsonColumn;
    require SpreadsheetModel::FormatSampler;
    SpreadsheetModel::FormatSampler->new($includeJsonColumn)
      ->wsWrite( $wbook, $wsheet );
    undef $wbook;
    rename $$, join( ' ', 'Format sampler', @options ) . '.xlsx';
}

1;
