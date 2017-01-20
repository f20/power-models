package SpreadsheetModel::Data::RCode::PriceMaps;

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

use SpreadsheetModel::Data::RCode::AreaMaps;

sub maps3701rate3ts {
    my ( $self, $script ) = @_;
    SpreadsheetModel::Data::RCode::AreaMaps->rCode($script) . <<'EOR';
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t <- dbGetQuery(db, paste(
    'select company, period, option, a.v as tariff, b.v as value',
    'from data as a',
    'inner join data as b using (bid, tab, row)',
    'inner join data as c using (bid, tab)',
    'left join books using (bid)',
    'where a.tab=3701 and a.row>0 and a.col=0 and b.col=c.col and c.row=0',
    'and c.v like "Unit rate 3 p/kWh"'
    )
);
period <- factor(t$period);
periodList <- gsub(' 02', '', fixed=T, levels(period));
levels(period) <- periodList;
numPeriods <- length(periodList);
company <- factor(t$company);
v <- as.numeric(t$value);
names(v) <- company;
tariff <- factor(t$tariff);
testkey <- factor(paste(tariff, period, company));
if (length(testkey) > length(levels(testkey))) {
	tariff <- factor(paste(tariff, t$option));
}
tariffList <- levels(tariff);

pdf('Green unit rates.pdf', width=11.69, height=8.27);
for (t in tariffList[order(tariffList)]) {
    if (sum(v[tariff==t] != 0) > 0) {
        l <- list();
        for (o in 1:numPeriods) {
            l[[o]] <- v[tariff==t&period==periodList[o]];
        }
        try(plot.dno.map(
            l,
            file.name=NA,
            title=t,
            option.names=paste(periodList, 'off-peak p/kWh'),
            legend.digit=1,
            number.format='%1.2f'
        ));
    }
}
graphics.off();

for (t in tariffList[order(tariffList)]) {
    if (t == "HV HH Metered" || t == "LV Sub HH Metered" || t == "LV HH Metered") {
        l <- list();
        for (o in 1:numPeriods) {
            l[[o]] <- v[tariff==t&period==periodList[o]];
        }
        try(plot.dno.map(
            l,
            file.name=paste(t, 'green'),
            file.type=1200,
            title=paste(t, 'Green unit rate', sep=' \U{2014} '),
            option.names=paste(periodList, 'off-peak p/kWh'),
            legend.digit=1,
            number.format='%1.2f'
        ));
    }
}

EOR
}

sub maps3701reactivets {
    my ( $self, $script ) = @_;
    SpreadsheetModel::Data::RCode::AreaMaps->rCode($script) . <<'EOR';
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t <- dbGetQuery(db, paste(
    'select company, period, option, a.v as tariff, b.v as value',
    'from data as a',
    'inner join data as b using (bid, tab, row)',
    'inner join data as c using (bid, tab)',
    'left join books using (bid)',
    'where a.tab=3701 and a.row>0 and a.col=0 and b.col=c.col and c.row=0',
    'and c.v like "%kVArh%"'
    )
);
period <- factor(t$period);
periodList <- gsub(' 02', '', fixed=T, levels(period));
levels(period) <- periodList;
numPeriods <- length(periodList);
company <- factor(t$company);
v <- as.numeric(t$value);
names(v) <- company;
tariff <- factor(t$tariff);
testkey <- factor(paste(tariff, period, company));
if (length(testkey) > length(levels(testkey))) {
	tariff <- factor(paste(tariff, t$option));
}
tariffList <- levels(tariff);

pdf('Reactive power charges.pdf', width=11.69, height=8.27);
for (t in tariffList[order(tariffList)]) {
    if (sum(v[tariff==t] != 0) > 0) {
        l <- list();
        for (o in 1:numPeriods) {
            l[[o]] <- v[tariff==t&period==periodList[o]];
        }
        try(plot.dno.map(
            l,
            file.name=NA,
            title=t,
            option.names=paste(periodList, 'p/kVArh'),
            legend.digit=1,
            number.format='%1.2f'
        ));
    }
}
graphics.off();

for (t in tariffList[order(tariffList)]) {
    if (t == "HV HH Metered" || t == "LV Sub HH Metered" || t == "LV HH Metered") {
        l <- list();
        for (o in 1:numPeriods) {
            l[[o]] <- v[tariff==t&period==periodList[o]];
        }
        try(plot.dno.map(
            l,
            file.name=paste(t, 'reactive'),
            file.type=1200,
            title=paste(t, 'Excess reactive power charges', sep=' \U{2014} '),
            option.names=paste(periodList, 'p/kVArh'),
            legend.digit=1,
            number.format='%1.2f'
        ));
    }
}

EOR
}

sub maps4202ts {
    my ( $self, $script ) = @_;
    SpreadsheetModel::Data::RCode::AreaMaps->rCode($script) . <<'EOR';
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t4202 <- dbGetQuery(db, paste(
    'select company, period, option, a.v as tariff, b.v as peryear, c.v as permwh',
    'from data as a inner join data as b using (bid, tab, row) inner join data as c using (bid, tab, row)',
    'left join books using (bid) where a.tab=4202 and a.row>0 and a.col=0 and b.col=1 and c.col=2',
    'and c.v+0 > 0'
    )
);
period <- factor(t4202$period);
periodList <- levels(period);
numPeriods <- length(periodList);
company <- factor(t4202$company);
peryear <- as.numeric(t4202$peryear);
names(peryear) <- company;
ppu <- as.numeric(t4202$permwh)*0.1;
names(ppu) <- company;
tariff <- factor(t4202$tariff);
testkey <- factor(paste(tariff, period, company));
if (length(testkey) > length(levels(testkey))) {
	tariff <- factor(paste(tariff, t4202$option));
}
tariffList <- levels(tariff);

pdf('Illustrative charges over time.pdf', width=11.69, height=8.27);
for (t in tariffList[order(tariffList)]) {
    l <- list();
    for (o in 1:numPeriods) {
        l[[o]] <- ppu[tariff==t&period==periodList[o]];
    }
    try(plot.dno.map(
        l,
        file.name=NA,
        title=t,
        option.names=paste(periodList, 'p/kWh'),
        legend.digit=1,
        number.format='%1.2f'
    ));
}
graphics.off();
EOR
}

sub maps4202cs {
    my ( $self, $script ) = @_;
    SpreadsheetModel::Data::RCode::AreaMaps->rCode($script) . <<'EOR';
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t4202 <- dbGetQuery(db, paste(
    'select company, period, option, a.v as tariff, b.v as peryear, c.v as permwh',
    'from data as a inner join data as b using (bid, tab, row) inner join data as c using (bid, tab, row)',
    'left join books using (bid) where a.tab=4202 and a.row>0 and a.col=0 and b.col=1 and c.col=2',
    'and c.v+0 > 0'
	)
);
option <- factor(t4202$option);
optionList <- levels(option);
numOptions <- length(optionList);
company <- factor(t4202$company);
peryear <- as.numeric(t4202$peryear);
names(peryear) <- company;
ppu <- as.numeric(t4202$permwh)*0.1;
names(ppu) <- company;
tariff <- factor(t4202$tariff);
testkey <- factor(paste(tariff, option, company));
if (length(testkey) > length(levels(testkey))) {
	tariff <- factor(paste(tariff, t4202$period));
}
tariffList <- levels(tariff);

pdf('Illustrative charge comparison.pdf', width=11.69, height=8.27);
for (t in tariffList[order(sapply(tariffList, function (t) { 1/max(ppu[tariff==t]); }))]) {
    l <- list();
    for (o in 1:numOptions) {
        l[[o]] <- ppu[tariff==t&option==optionList[o]];
    }
    try(plot.dno.map(
        l,
        file.name=NA,
        title=t,
        option.names=paste(optionList, 'p/kWh'),
        legend.digit=1,
        number.format='%1.2f'
    ));
}
graphics.off();
EOR
}

sub margins4203 {
    my ( $self, $script ) = @_;
    SpreadsheetModel::Data::RCode::AreaMaps->rCode($script) . <<'EOR';
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t4203 <- dbGetQuery(db, paste(
    'select company, period, option, a.v as tariff, b.v as margin, c.v as boundary',
    'from data as a inner join data as b using (bid, tab, row) inner join data as c using (bid, tab)',
    'left join books using (bid) where a.tab=4203 and a.row>0 and a.col=0 and b.col>0',
    'and c.row=0 and c.col=b.col',
    'and b.v+0 > 0'
    )
);
boundary <- factor(t4203$boundary);
boundaryList <- levels(boundary);
numBoundaries <- length(boundaryList);
company <- factor(t4203$company);
value <- as.numeric(t4203$margin);
names(value) <- company;
tariff <- factor(t4203$tariff);
testkey <- factor(paste(tariff, boundary, company));
if (length(testkey) > length(levels(testkey))) {
	tariff2 <- factor(paste(tariff, t4203$period));
	testkey2 <- factor(paste(tariff2, boundary, company));
	if (length(levels(testkey2)) > length(levels(testkey))) {
		tariff <- tariff2;
		testkey <- testkey2;
	}
}
if (length(testkey) > length(levels(testkey))) {
	tariff <- factor(paste(tariff, t4203$option));
}

pdf('Illustrative margins.pdf', width=11.69, height=8.27);
for (t in levels(tariff)) {
    l <- list();
    scaling <- 1;
    units <- '£';
    format <- '%1.0f';
    digit <- 0;
    mx <- max(value[tariff==t]);
    if (mx > 5000) {
    	scaling <- 0.001;
	    units <- '£k';
		if (mx < 35000) {
			format <- '%1.1f';
			digit <- 1;
		}
    }
    for (o in 1:numBoundaries) {
        l[[o]] <- scaling*value[tariff==t&boundary==boundaryList[o]];
    }
    try(plot.dno.map(
        l,
        file.name = NA,
        title = t,
        option.names = paste(boundaryList, units),
        legend.digit = digit,
        number.format = format,
		colour.maker = function (i) { hsv(0.25+0.004*i, 1-0.006*i, 1); },
		mincol.maker = function (v) { min(v[2:length(v)], na.rm=T); },
		maxcol.maker = function (v) { max(v[2:length(v)], na.rm=T); }
    ));
}
graphics.off();
EOR
}

sub mapCdcmEdcmHardCoded {
    my ( $self, $script ) = @_;
    SpreadsheetModel::Data::RCode::AreaMaps->rCode($script) . <<'EOR';
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
