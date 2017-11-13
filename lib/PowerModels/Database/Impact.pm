package PowerModels::Database;

=head Copyright licence and disclaimer

Copyright 2009-2016 Franck Latrémolière, Reckon LLP and others.

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
use PowerModels::Database::ImpactGeneric;
use PowerModels::Database::ImpactSpecialist;

sub cdcmRevenueMatrixImpact {
    my ( $self, $wbmodule, %options ) = @_;
    $options{tableNumber} = 3901;
    $options{col1}        = 0;
    $options{col2}        = 25;
    $options{columns}     = [24];
    $self->findLines( \%options, 1 );
    $self->revenueMatrixImpact( $wbmodule, %options );
}

sub cdcmTariffImpact {
    my ( $self, $wbmodule, %options ) = @_;
    $self->defaultOptions( \%options );
    $options{tableNumber} ||= 3701;
    $self->findComponents( \%options );
    $self->genericTariffImpact( $wbmodule, %options );
}

sub edcmTariffImpact {
    my ( $self, $wbmodule, %options ) = @_;
    $options{components} ||= [ split /\n/, <<EOL];
Import super-red unit rate (p/kWh)
Import fixed charge (p/day)
Import capacity rate (p/kVA/day)
Import exceeded capacity rate (p/kVA/day)
Export super-red unit rate (p/kWh)
Export fixed charge (p/day)
Export capacity rate (p/kVA/day)
Export exceeded capacity rate (p/kVA/day)
EOL
    $options{tableNumber}       ||= 4501;
    $options{firstColumnBefore} ||= 2;
    $options{firstColumnAfter}  ||= 2;
    $options{nameExtraColumn}   ||= 1;
    $self->genericTariffImpact( $wbmodule, %options );
}

sub edcmRevenueMatrixImpact {
    my ( $self, $wbmodule, %options ) = @_;
    $options{tableNumber}     = 4601;
    $options{col1}            = 9;
    $options{col2}            = 33;
    $options{columns}         = [ 16, 20 ];
    $options{nameExtraColumn} = 1;
    $self->revenueMatrixImpact( $wbmodule, %options );
}

1;
