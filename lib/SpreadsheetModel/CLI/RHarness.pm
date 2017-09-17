package SpreadsheetModel::CLI::RHarness;

=head Copyright licence and disclaimer

Copyright 2011-2017 Franck Latrémolière and others. All rights reserved.

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

sub R {
    my ( $self, @commands ) = @_;
    open my $r, '| R --vanilla --slave';
    binmode $r, ':utf8';
    require SpreadsheetModel::Data::RCode;
    print {$r} SpreadsheetModel::Data::RCode->rCode(@commands);
}

sub Rcode {
    my ( $self, @commands ) = @_;
    open my $r, '>', "$$.R";
    binmode $r, ':utf8';
    print $r "# R code from power-models\n\n";
    require SpreadsheetModel::Data::RCode;
    print {$r} SpreadsheetModel::Data::RCode->rCode(@commands);
    close $r;
    rename "$$.R", 'power-models.R';
    warn <<EOW
To use this R code, say:
    source("power-models.R");
EOW
}

1;
