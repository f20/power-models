package SpreadsheetModel::Data::RCode::MapsFromDatabase;

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
periodList <- gsub(' 02', '', fixed=T, levels(period));
levels(period) <- periodList;
numPeriods <- length(periodList);
company <- factor(t4202$company);
peryear <- as.numeric(t4202$peryear);
names(peryear) <- company;
ppu <- as.numeric(t4202$permwh)*0.1;
names(ppu) <- company;
tariff <- factor(gsub('([0-9])([0-9][0-9][0-9])','\\1,\\2kWh',t4202$tariff));
testkey <- factor(paste(tariff, period, company));
if (length(testkey) > length(levels(testkey))) {
	tariff <- factor(paste(tariff, t4202$option));
}
tariffList <- levels(tariff);

myTariffList <- tariffList[order(tariffList)];

pdf('Illustrative charges over time.pdf', width=11.69, height=8.27);
for (t in myTariffList) {
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

myTariffList <- myTariffList[grep(':', myTariffList, invert=T)];
for (t in myTariffList) {
    minimum <- min(ppu[tariff==t]);
    maximum <- max(ppu[tariff==t]);
    for (o in 2:numPeriods) {
        oo <- 1;
        l <- list();
        for (ooo in (o-oo):o) {
            l[[ooo-o+oo+1]] <- ppu[tariff==t&period==periodList[ooo]];
        }
        try(plot.dno.map(
            l,
            mincol=minimum,
            maxcol=maximum,
            file.name=paste(t,
                ifelse(oo, paste(periodList[o-oo], periodList[o], sep='-'), periodList[o])
            ),
            file.type='1080p',
            title=t,
            option.names=paste(periodList[(o-oo):o], 'p/kWh')
        ));
    }
}

myTariffList <- myTariffList[grep('[0-9],[0-9][0-9][0-9]', myTariffList)];
for (t in myTariffList) {
    minimum <- min(peryear[tariff==t]);
    maximum <- max(peryear[tariff==t]);
    for (o in 1:numPeriods) {
        oomax <- min(o-1, 2);
        for (oo in 0:oomax) {
            l <- list();
            for (ooo in (o-oo):o) {
                l[[ooo-o+oo+1]] <- peryear[tariff==t&period==periodList[ooo]];
            }
            try(plot.dno.map(
                l,
                mincol=minimum,
                maxcol=maximum,
                file.name=paste(t, 'annual',
                    ifelse(oo, paste(periodList[o-oo], periodList[o], sep='-'), periodList[o])
                ),
                file.type='1080p',
                title=t,
                option.names=paste(periodList[(o-oo):o], '\U{a3}/year'),
                legend.digit=0,
                number.format='%.0f'
            ));
        }
    }
}

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
));
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

1;
