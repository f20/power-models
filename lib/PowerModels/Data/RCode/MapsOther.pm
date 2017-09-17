package PowerModels::Data::RCode;

=head Copyright licence and disclaimer

Copyright 2014-2017 Franck Latrémolière, Reckon LLP and others.

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

use PowerModels::Data::RCode::AreaMaps;

sub svgNoText {
    my ( $self, $rIncluded ) = @_;
    PowerModels::Data::RCode::AreaMaps->rCode($rIncluded) . <<'EOR';
plot.dno.map(
    data.frame(1:14),
    file.name='Map without text',
    file.type='svg',
    number.show=F,
    legend.show=F
);
EOR
}

sub highlightEachArea {
    my ( $self, $rIncluded ) = @_;
    PowerModels::Data::RCode::AreaMaps->rCode($rIncluded) . <<'EOR';
for (i in 1:14) {
    f<-rep(NA, 14);
    names(f)<-dno.areas;
    f[i]<-1;
    plot.dno.map(
        f,
        file.name=i,
        file.type="640",
        number.show=FALSE,
        legend.show=FALSE
    );
}
EOR
}

sub edcmOpacity {
    my ( $self, $rIncluded ) = @_;
    PowerModels::Data::RCode::AreaMaps->rCode($rIncluded) . <<'EOR';
csv <- textConnection('
    Opacity 2013,Opacity 2014,DNO
    0.25,0.25,ENWL
    0.25,10,NPG Northeast
    0.25,10,NPG Yorkshire
    0.25,10,SPEN SPD
    0.25,10,SPEN SPM
    0.25,0.5,SSEPD SEPD
    0.25,0.5,SSEPD SHEPD
    0,0,UKPN EPN
    0.25,0,UKPN LPN
    0.25,0,UKPN SPN
    0.25,10,WPD EastM
    0.25,10,WPD SWales
    0.25,10,WPD SWest
    0.25,10,WPD WestM
');
v <- read.csv(csv);
names(v)[1:2] <- paste('Opacity', 2013:2014);
for (t in c('pdf', 'Word')) {
    plot.dno.map(
        v[,1:2],
        file.name='EDCM opacity ratings', file.type=t, number.show=F, legend.show=F,
        title='EDCM opacity ratings',
        box=paste(
            "Green is good, red is bad.",
            "Ratings reflect Franck's experience in seeking non-confidential",
            "aggregate data that would help understand how each",
            "DNO's EDCM charging model works in practice."
        )
    );
}
EOR
}

sub mapCdcmEdcm2013 {
    my ( $self, $rIncluded ) = @_;
    PowerModels::Data::RCode::AreaMaps->rCode($rIncluded) . <<'EOR';
v <- read.csv(textConnection('
CDCM,EDCM Zero,EDCM Avg,DNO
9.230136098,4.697991667,5.961880556,ENWL
10.18231688,3.526735185,4.017475926,NPG Northeast
8.182795992,3.45787963,3.888435185,NPG Yorkshire
9.469356279,4.014021296,4.412169444,SPEN SPD
13.49652633,6.942557407,7.78515,SPEN SPM
8.62859187,3.878862037,4.332565741,SSEPD SEPD
17.19461702,5.597428704,5.949280556,SSEPD SHEPD
6.466388597,3.012714815,3.836788889,UKPN EPN
6.819068962,3.312843519,4.067473148,UKPN LPN
8.706796271,3.974465741,4.895762037,UKPN SPN
6.692516502,4.549932407,4.786043519,WPD EastM
15.6681223,5.692940741,6.239237037,WPD SWales
12.97384748,4.284894444,5.044153704,WPD SWest
8.052028412,3.748900926,4.628530556,WPD WestM
'));
names(v)[1:3] <- c('CDCM', 'EDCM, zero charge 1', 'EDCM, average charge 1');
plot.dno.map(
    v[,1:3],
    file.name='CDCM HV v EDCM 1111, February 2013 data',
    file.type='Word',
    title='DUoS charge for HV continuous load (\U{a3}/MWh)'
);
EOR
}

1;
