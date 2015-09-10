package Compilation::RCodeGenerator;

=head Copyright licence and disclaimer

Copyright 2014-2015 Franck Latrémolière, Reckon LLP and others.

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

sub convert {
    my $self = shift;
    foreach (@_) {
        $_ = $self->$_() if $self->can($_);
    }
}

sub maps4202ts {
    <<'EOR';
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t4202 <- dbGetQuery(db, paste(
    'select company, period, option, a.v as tariff, b.v as peryear, c.v as permwh',
    'from data as a inner join data as b using (bid, tab, row) inner join data as c using (bid, tab, row)',
    'left join books using (bid) where a.tab=4202 and a.row>0 and a.col=0 and b.col=1 and c.col=2',
    'and c.v+0 > 0'
    # order by company, period, option, tariff
    )
);
company <- factor(t4202$company);
period <- factor(t4202$period);
option <- factor(t4202$option);
tariff <- factor(t4202$tariff);
peryear <- as.numeric(t4202$peryear);
names(peryear) <- company;
ppu <- as.numeric(t4202$permwh)*0.1;
names(ppu) <- company;

tariffList <- levels(tariff);
periodList <- levels(period);
numPeriods <- length(periodList);

pdf('Graphs.pdf', width=11.69, height=8.27);
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
    <<'EOR';
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t4202 <- dbGetQuery(db, paste(
    'select company, period, option, a.v as tariff, b.v as peryear, c.v as permwh',
    'from data as a inner join data as b using (bid, tab, row) inner join data as c using (bid, tab, row)',
    'left join books using (bid) where a.tab=4202 and a.row>0 and a.col=0 and b.col=1 and c.col=2',
    'and c.v+0 > 0'
    # order by company, period, option, tariff
    )
);
company <- factor(t4202$company);
period <- factor(t4202$period);
option <- factor(t4202$option);
tariff <- factor(t4202$tariff);
peryear <- as.numeric(t4202$peryear);
names(peryear) <- company;
ppu <- as.numeric(t4202$permwh)*0.1;
names(ppu) <- company;

tariffList <- levels(tariff);
optionList <- levels(option);
numOptions <- length(optionList);

pdf('Graphs.pdf', width=11.69, height=8.27);
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

1;
