package DataManagement::Chedam;

=head Copyright licence and disclaimer

Copyright 2013 Franck Latrémolière and others. All rights reserved.

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

sub locateHidamModelled {
    my ( $self, $filename ) = @_;
    my %hash = (
        name          => "$filename (modelled)",
        MVA_BSP       => [ 5613, 4, 1 ],
        MVA_Dis       => [ 6604, 1, 1 ],
        MVA_GSP_132kV => [ 5613, 1, 1 ],
        MVA_GSP_EHV   => [ 5613, 2, 1 ],
        MVA_GSP_HV    => [ 5613, 3, 1 ],
        MVA_Pri_132kV => [ 5613, 5, 1 ],
        MVA_Pri_EHV   => [ 5613, 6, 1 ],
        count_BSP       => [ 1702, 8 ],
        count_GSP_132kV => [ 1702, 2 ],
        count_GSP_EHV   => [ 1702, 4 ],
        count_GSP_HV    => [ 1702, 6 ],
        count_Pri_132kV => [ 1702, 10 ],
        count_Pri_EHV   => [ 1702, 12 ],
        km_132kV        => [ 5705, 1, 1 ],
        km_EHV          => [ 5705, 2, 1 ],
        km_HV           => [ 6708, 1, 1 ],
        km_LV           => [ 6708, 2, 1 ],
        MVA_Dis_Ground  => [ 6605, 1, 1 ],
        MVA_Dis_Pole    => [ 6606, 1, 1 ],
    );
    bless \%hash, $self;
}

sub locateHidamAdjMMD {
    my ( $self, $filename ) = @_;
    my %hash = (
        name           => "$filename (adjusted MMD)",
        MW_BSP         => [ 5611, 1, 1 ],
        MW_GSP_132kV   => [ 5608, 1, 1 ],
        MW_GSP_EHV     => [ 5608, 2, 1 ],
        MW_GSP_HV      => [ 5608, 3, 1 ],
        MW_Pri_132kV   => [ 5611, 2, 1 ],
        MW_Pri_EHV     => [ 5618, 1, 1 ],
        MVA_Dis_Pole   => [ 6606, 1, 1 ],
        MVA_Dis_Ground => [ 6605, 1, 1 ],
    );
    bless \%hash, $self;
}

sub locateHidamActualCap {
    my ( $self, $filename ) = @_;
    my %hash = (
        name           => "$filename (actual capacity)",
        MVA_GSP_132kV  => [ 5601, 1, 1 ],
        MVA_GSP_EHV    => [ 5601, 4, 1 ],
        MVA_GSP_HV     => [ 5601, 7, 1 ],
        MVA_BSP        => [ 5601, 10, 1 ],
        MVA_Pri_132kV  => [ 5601, 13, 1 ],
        MVA_Pri_EHV    => [ 5601, 16, 1 ],
        MVA_132kV      => [ 5605, 1, 1 ],
        MVA_EHV        => [ 5605, 4, 1 ],
        MVA_Dis_Pole   => [ 1738, 1, 1 ],
        MVA_Dis_Ground => [ 1738, 2, 1 ],
    );
    bless \%hash, $self;
}

sub locateHidamActualMD {
    my ( $self, $filename ) = @_;
    my %hash = (
        name           => "$filename (actual AMD)",
        MW_GSP_132kV   => [ 5601, 2, 1 ],
        MW_GSP_EHV     => [ 5601, 5, 1 ],
        MW_GSP_HV      => [ 5601, 8, 1 ],
        MW_BSP         => [ 5601, 11, 1 ],
        MW_Pri_132kV   => [ 5601, 14, 1 ],
        MW_Pri_EHV     => [ 5601, 17, 1 ],
        MW_132kV       => [ 5605, 2, 1 ],
        MW_EHV         => [ 5605, 5, 1 ],
        MVA_Dis_Pole   => [ 1738, 1, 1 ],
        MVA_Dis_Ground => [ 1738, 2, 1 ],
    );
    bless \%hash, $self;
}

sub locateDrm {
    my ( $self, $filename ) = @_;
    my %hash = (
        name          => "$filename (CDCM AMD)",
        Div_GSP       => [ 2612, 1, 1 ],
        Div_BSP       => [ 2612, 3, 1 ],
        Div_Pri       => [ 2612, 5, 1 ],
        Div_Dis       => [ 2612, 8, 1 ],
        SMD_132kV     => [ 2108, 1, 1 ],
        SMD_132kV_EHV => [ 2108, 1, 2 ],
        SMD_EHV       => [ 2108, 1, 3 ],
        SMD_EHV_HV    => [ 2108, 1, 4 ],
        SMD_132kV_HV  => [ 2108, 1, 5 ],
        SMD_HV        => [ 2108, 1, 6 ],
        SMD_HV_LV     => [ 2108, 1, 7 ],
        SMD_LV        => [ 2108, 1, 8 ],
    );
    bless \%hash, $self;
}

1;
