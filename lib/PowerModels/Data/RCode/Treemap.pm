package PowerModels::Data::RCode::Treemap;

=head Copyright licence and disclaimer

Copyright 2015 Franck Latrémolière, Reckon LLP and others.

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

sub treemapWithCategories {
    my ( $self, $byCategory ) = @_;
    (
        $byCategory
        ? <<'EOR'
columnIndex <- c('category', 'tariff');
fileName <- 'Treemaps by category';
EOR
        : <<'EOR'
columnIndex <- c('tariff', 'category');
fileName <- 'Treemaps with categories';
EOR
      ) . <<'EOR'
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t3901 <- dbGetQuery(db, paste(
    'select company, period, option,',
    'a.v as tariff, b.v as category, c.v as amount,',
    'a.row as tariffid, b.col as categoryid',
    'from data as a',
    'inner join data as b using (bid, tab)',
    'inner join data as c using (bid, tab)',
    'left join books using (bid)',
    'where a.tab=3901',
    'and a.row=c.row and a.col=0',
    'and b.row=0 and b.col=c.col',
    'and c.col>0 and c.row>0 and c.v+0<>0'
));
company <- factor(t3901$company);
period <- factor(t3901$period);
option <- factor(t3901$option);
for (i in 1:length(t3901$amount)) {
    if (t3901$amount[i] < 0) {
        t3901$amount[i] <- -t3901$amount[i];
        t3901$category[i] <- paste('[negative]', t3901$category[i]);
    }
}
t3901$category<-factor(sub(" \\(.+\\)", "", t3901$category, perl=TRUE));
t3901$tariff<-factor(t3901$tariff);
library(treemap);
pdf(paste(fileName, 'pdf', sep='.'), width=11.69, height=8.27);
for (c in levels(company)) {
    for (p in levels(period)) {
        for (o in levels(option)) {
            name <- paste(c, p, o);
            filter <- company==c&period==p&option==o&t3901$category!='Total net revenue by tariff';
            if (sum(filter)>0) {
                tryCatch(treemap(
                    t3901[filter, ],
                    index=columnIndex,
                    vSize='amount',
EOR
      . (
        $byCategory
        ? ''
        : <<'EOR'
                    vColor='tariffid',
                    type='manual',
                    palette=rep(rainbow(25),max(t3901$tariffid)/25),
                    range=c(1,max(t3901$tariffid)),
EOR
      ) . <<'EOR';
                    title=name,
                    position.legend='none'
                ), error=function(e) e);
            }
        }
    }
}
graphics.off();
EOR
}

sub treemapWithComponents {
    my ( $self, $byCategory ) = @_;
    (
        $byCategory
        ? <<'EOR'
columnIndex <- c('category', 'tariff');
fileName <- 'Treemaps by component';
EOR
        : <<'EOR'
columnIndex <- c('tariff', 'category');
fileName <- 'Treemaps with components';
EOR
      ) . <<'EOR'
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t3801 <- dbGetQuery(db, paste(
    'select company, period, option,',
    'a.v as tariff, b.v as category, c.v as amount,',
    'a.row as tariffid, b.col as categoryid',
    'from data as a',
    'inner join data as b using (bid, tab)',
    'inner join data as c using (bid, tab)',
    'left join books using (bid)',
    'where a.tab=3801',
    'and a.row=c.row and a.col=0',
    'and b.row=0 and b.col=c.col',
    'and c.col>3 and c.col<8 and c.row>0 and c.v+0<>0'
));
company <- factor(t3801$company);
period <- factor(t3801$period);
option <- factor(t3801$option);
for (i in 1:length(t3801$amount)) {
    if (t3801$amount[i] < 0) {
        t3801$amount[i] <- -t3801$amount[i];
        t3801$category[i] <- paste('[negative]', t3801$category[i]);
    }
}
t3801$category<-factor(sub(" \\(.+\\)", "", sub("Revenues from ", "", t3801$category), perl=TRUE));
t3801$tariff<-factor(t3801$tariff);
library(treemap);
pdf(paste(fileName, 'pdf', sep='.'), width=11.69, height=8.27);
for (c in levels(company)) {
    for (p in levels(period)) {
        for (o in levels(option)) {
            name <- paste(c, p, o);
            filter <- company==c&period==p&option==o&t3801$category!='Total net revenue by tariff';
            if (sum(filter)>0) {
                treemap(
                    t3801[filter, ],
                    index=columnIndex,
                    vSize='amount',
EOR
      . (
        $byCategory
        ? ''
        : <<'EOR'
                    vColor='tariffid',
                    type='manual',
                    palette=rainbow(max(t3801$tariffid)),
                    range=c(1,max(t3801$tariffid)),
EOR
      ) . <<'EOR';
                    title=name,
                    position.legend='none'
                );
            }
        }
    }
}
graphics.off();
EOR
}

sub treemap1020 {
    my ( $self, $byCompany ) = @_;
    (
        $byCompany
        ? <<'EOR'
columnIndex <- c('company', 'level');
fileName <- 'Treemaps 500MW by company';
EOR
        : <<'EOR'
columnIndex <- c('level', 'company');
fileName <- 'Treemaps 500MW with companies';
EOR
      ) . <<'EOR'
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t <- dbGetQuery(db, paste(
    'select company, period, option,',
    'a.v as level, b.v as amount',
    'from data as a',
    'inner join data as b using (bid, tab)',
    'left join books using (bid)',
    'where a.tab=1020',
    'and a.row=b.row and a.col=0',
    'and b.col=1 and b.row>0 and b.v+0<>0'
));
company <- factor(t$company);
period <- factor(t$period);
option <- factor(t$option);
level <- factor(t$level);
library(treemap);
pdf(paste(fileName, 'pdf', sep='.'), width=11.69, height=8.27);
for (p in levels(period)) {
	for (o in levels(option)) {
		name <- p;
		filter <- period==p&option==o;
		if (sum(filter)>0) {
			tryCatch(treemap(
				t[filter, ],
				index=columnIndex,
				vSize='amount',
				title=name,
				position.legend='none'
			), error=function(e) e);
		}
	}
}
graphics.off();
EOR
}

sub treemap2706 {
    my ( $self, $byCompany ) = @_;
    (
        $byCompany
        ? <<'EOR'
columnIndex <- c('company', 'level');
fileName <- 'Treemaps assets by company';
EOR
        : <<'EOR'
columnIndex <- c('level', 'company');
fileName <- 'Treemaps assets with companies';
EOR
      ) . <<'EOR'
library(DBI);
library(RSQLite);
drv <- dbDriver('SQLite');
db <- dbConnect(drv, dbname = '~$database.sqlite');
t <- dbGetQuery(db, paste(
    'select company, period, option,',
    'a.v as level, b.v as amount',
    'from data as a',
    'inner join data as b using (bid, tab)',
    'left join books using (bid)',
    'where a.tab=2706',
    'and a.col=b.col and a.row=0',
    'and b.row=1 and b.col>0 and b.v+0<>0'
));
company <- factor(t$company);
period <- factor(t$period);
option <- factor(t$option);
level <- factor(t$level);
library(treemap);
pdf(paste(fileName, 'pdf', sep='.'), width=11.69, height=8.27);
for (p in levels(period)) {
	for (o in levels(option)) {
		name <- p;
		filter <- period==p&option==o&level!="Denominator for allocation of operating expenditure";
		if (sum(filter)>0) {
			tryCatch(treemap(
				t[filter, ],
				index=columnIndex,
				vSize='amount',
				title=name,
				position.legend='none'
			), error=function(e) e);
		}
	}
}
graphics.off();
EOR
}

1;
