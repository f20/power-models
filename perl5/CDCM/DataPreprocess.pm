package CDCM;

=head Copyright licence and disclaimer

Copyright 2012 Reckon LLP and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation and/or
other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY RECKON LLP AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL RECKON LLP OR CONTRIBUTORS BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;

sub preprocessDataset {

    my ($model) = @_;
    my $d = $model->{dataset};

    $d->{1000}[3]{'Company charging year data version'} = $model->{version}
      if $model->{version};

    if ( $model->{single1076} && $d->{1076}[4] ) {
        foreach ( grep { !/^_/ } keys %{ $d->{1076}[1] } ) {
            $d->{1076}[1]{$_} += $d->{1076}[2]{$_} if $d->{1076}[2]{$_};
            $d->{1076}[2]{$_} = '';
            $d->{1076}[1]{$_} += $d->{1076}[3]{$_} if $d->{1076}[3]{$_};
            $d->{1076}[3]{$_} = '';
            $d->{1076}[1]{$_} -= $d->{1076}[4]{$_} if $d->{1076}[4]{$_};
            $d->{1076}[4]{$_} = '';
        }
    }

}

1;
