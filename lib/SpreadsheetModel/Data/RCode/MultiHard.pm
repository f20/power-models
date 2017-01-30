package SpreadsheetModel::Data::RCode;

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

require SpreadsheetModel::Data::RCode::ExportGraph;

sub bandConsumption {
    my ( $self, $script ) = @_;
    SpreadsheetModel::Data::RCode::ExportGraph->rCode($script) . <<'EOR';
dset2<-read.csv(textConnection('
Customer type,DNO area,Time red,Time amber,Time green,Power red,Power amber,Power green
Domestic,ENWL,0.089480874,0.28904827,0.621470856,1.413854729,1.138344214,0.876067819
Domestic,NPG Northeast,0.104394353,0.31318306,0.582422587,1.431681015,1.113331115,0.861683857
Domestic,NPG Yorkshire,0.104394353,0.31318306,0.582422587,1.4290665,1.113159411,0.862244816
Domestic,SPEN SPD,0.089480874,0.390368852,0.520150273,1.417440562,1.172351624,0.798839571
Domestic,SPEN SPM,0.089480874,0.390368852,0.520150273,1.419957823,1.172077826,0.798612013
Domestic,SSEPD SEPD,0.074567395,0.268442623,0.656989982,1.457076152,1.094721942,0.909419646
Domestic,SSEPD SHEPD,0.193875228,0.276980874,0.529143898,1.320083048,1.113227277,0.823454518
Domestic,UKPN EPN,0.089480874,0.387750455,0.52276867,1.406580119,1.125263801,0.83749554
Domestic,UKPN LPN,0.178961749,0.298269581,0.52276867,1.202450183,1.166994161,0.835414549
Domestic,UKPN SPN,0.089480874,0.387750455,0.52276867,1.392319755,1.133838683,0.833576245
Domestic,WPD EastM,0.089480874,0.31318306,0.597336066,1.690868421,1.327494178,0.724802923
Domestic,WPD SWales,0.074567395,0.428961749,0.496470856,1.515183785,1.170822387,0.775027716
Domestic,WPD SWest,0.059653916,0.393442623,0.546903461,1.487377245,1.168802567,0.825402301
Domestic,WPD WestM,0.089480874,0.31318306,0.597336066,1.689169704,1.329399588,0.724058385
Business,ENWL,0.089480874,0.28904827,0.621470856,1.229737707,1.559846505,0.706535242
Business,NPG Northeast,0.104394353,0.31318306,0.582422587,1.159086422,1.546158448,0.677802162
Business,NPG Yorkshire,0.104394353,0.31318306,0.582422587,1.163447867,1.547018023,0.676558195
Business,SPEN SPD,0.089480874,0.390368852,0.520150273,1.220309415,1.394453301,0.666066192
Business,SPEN SPM,0.089480874,0.390368852,0.520150273,1.230328287,1.397915555,0.661744261
Business,SSEPD SEPD,0.074567395,0.268442623,0.656989982,1.130974277,1.637147144,0.724799668
Business,SSEPD SHEPD,0.193875228,0.276980874,0.529143898,1.193978415,1.443648402,0.696699268
Business,UKPN EPN,0.089480874,0.387750455,0.52276867,1.167908077,1.349223862,0.712231659
Business,UKPN LPN,0.178961749,0.298269581,0.52276867,1.553962706,1.230891822,0.678622398
Business,UKPN SPN,0.089480874,0.387750455,0.52276867,1.229184614,1.393541572,0.668871581
Business,WPD EastM,0.089480874,0.31318306,0.597336066,1.483153381,1.882933646,0.464701922
Business,WPD SWales,0.074567395,0.428961749,0.496470856,0.954893842,1.36264193,0.693444105
Business,WPD SWest,0.059653916,0.393442623,0.546903461,1.045426606,1.455460854,0.667386292
Business,WPD WestM,0.089480874,0.31318306,0.597336066,1.479451032,1.87531141,0.469252869
GSP,ENWL,0.089480874,0.28904827,0.621470856,9.706678641,0.390659284,0.029797887
GSP,NPG Northeast,0.104394353,0.31318306,0.582422587,8.881105051,0.209397111,0.012504929
GSP,NPG Yorkshire,0.104394353,0.31318306,0.582422587,8.40200802,0.368810905,0.012658297
GSP,SPEN SPD,0.089480874,0.390368852,0.520150273,8.947913591,0.482364005,0.02121115
GSP,SPEN SPM,0.089480874,0.390368852,0.520150273,9.190344497,0.323816307,0.09849488
GSP,SSEPD SEPD,0.074567395,0.268442623,0.656989982,12.1238424,0.35745684,0
GSP,SSEPD SHEPD,0.193875228,0.276980874,0.529143898,2.773536395,1.053454205,0.322457969
GSP,UKPN EPN,0.089480874,0.387750455,0.52276867,11.06447552,0.023980449,0.001229275
GSP,UKPN LPN,0.178961749,0.298269581,0.52276867,5.498110264,0.053805597,0
GSP,UKPN SPN,0.089480874,0.387750455,0.52276867,10.92418787,0.037770565,0.015013456
GSP,WPD EastM,0.089480874,0.31318306,0.597336066,10.98114215,0,0.029125646
GSP,WPD SWales,0.074567395,0.428961749,0.496470856,8.43252191,0.85048253,0.01285943
GSP,WPD SWest,0.059653916,0.393442623,0.546903461,16.57105314,0.029157452,0
GSP,WPD WestM,0.089480874,0.31318306,0.597336066,9.817143396,0.388122606,0
'));

loadPlot <- function (d) {
    library(grid);
    library(ggplot2);
    p <- ggplot();
    p <- p + theme_minimal(base_size=16);
    p <- p + theme(legend.position="none");
    m <- max(2, d[4:6]);
    p <- p + scale_y_continuous(limits=c(0, m));
    p <- p + ylab('Consumption relative to average consumption');
    p <- p + xlab('Proportion of year');
    p <- p + geom_rect(aes(xmin=x1, xmax=x2, ymin=0, ymax=y), data=data.frame(x1=c(0, d[1], d[1]+d[2]), x2=c(d[1], d[1]+d[2], 1), y=d[4:6]), linetype=0, fill=c('#FF0000', '#CC9900', '#009966'));
    p;
};

for (i in 1:length(dset2[, 1])) {
    exportGraph(loadPlot(as.matrix(dset2[i, 3:8])), file.name=paste(dset2[i, 2], dset2[i, 1], 'bands'), file.type=1800);
}
EOR
}

sub coincidenceWaterfall {
    my ( $self, $script ) = @_;
    SpreadsheetModel::Data::RCode::ExportGraph->rCode($script) . <<'EOR';
dset1<-read.csv(textConnection('
Customer type,DNO area,HV PP,P PP,BSP PP,GSP PP,Red,X,1/LF
Domestic,ENWL,1.285189769,1.285189769,1.32458814,1.372785233,1.413854729,1.974167296,2.26396621
Domestic,NPG Northeast,1.258427245,1.258427245,1.214707991,1.40665237,1.431681015,1.923646742,2.328771315
Domestic,NPG Yorkshire,1.277088459,1.277088459,1.310332245,1.388398668,1.4290665,1.938801589,2.313456269
Domestic,SPEN SPD,1.20496471,1.20496471,1,1.364465328,1.417440562,1.951577403,2.201027146
Domestic,SPEN SPM,1.274026222,1.274026222,1.259735588,1.356790983,1.419957823,1.887144993,2.242152466
Domestic,SSEPD SEPD,1.182663899,1.317027491,1.395587435,1.422305856,1.457076152,2.193442504,2.412075546
Domestic,SSEPD SHEPD,1.194065183,1.194065183,1,1.175006566,1.320083048,1.797982056,2.433323339
Domestic,UKPN EPN,1.364262314,1.364262314,1.38886677,1.40359861,1.406580119,2.066115761,2.255471338
Domestic,UKPN LPN,1.192741918,1.192741918,1.1978442,1.201881165,1.202450183,1.891073187,2.26065836
Domestic,UKPN SPN,1.369095538,1.369095538,1.380753774,1.384148822,1.392319755,2.104176277,2.261406359
Domestic,WPD EastM,1.553058604,1.553058604,1.651007839,1.674061008,1.690868421,2.147187516,2.397959434
Domestic,WPD SWales,1.345888427,1.345888427,1.355056116,1.384826917,1.515183785,1.948921587,2.272542661
Domestic,WPD SWest,1.298143284,1.298143284,1.40871352,1.483722625,1.487377245,2.208312706,2.405783168
Domestic,WPD WestM,1.565082407,1.565082407,1.5815105,1.645438414,1.689169704,2.079464382,2.329877536
Business,ENWL,1.282060023,1.282060023,1.264808181,1.257324456,1.229737707,1.709986567,2.532780512
Business,NPG Northeast,1.251878198,1.251878198,1.312109755,1.180965194,1.159086422,1.954716554,2.509812019
Business,NPG Yorkshire,1.272948273,1.272948273,1.278716058,1.20416268,1.163447867,1.890118196,2.62387782
Business,SPEN SPD,1.189430873,1.189430873,1,1.246985731,1.220309415,1.703640982,2.540220152
Business,SPEN SPM,1.287678622,1.287678622,1.226988716,1.222382849,1.230328287,1.705387205,2.525252525
Business,SSEPD SEPD,1.297230242,1.259240912,1.207173013,1.179544931,1.130974277,1.422073613,2.499456892
Business,SSEPD SHEPD,1.204392913,1.204392913,1,1.181981305,1.193978415,1.632634358,2.512245226
Business,UKPN EPN,1.164404704,1.164404704,1.163587448,1.1693012,1.167908077,1.462227139,2.470283687
Business,UKPN LPN,1.502844954,1.502844954,1.517453118,1.548777879,1.553962706,1.307430525,2.45320649
Business,UKPN SPN,1.22807871,1.22807871,1.230744786,1.227194059,1.229184614,1.167376826,2.564444898
Business,WPD EastM,1.550065486,1.550065486,1.521836541,1.465434567,1.483153381,1.627386795,2.630692694
Business,WPD SWales,1.092014293,1.092014293,1.058616976,1.101981142,0.954893842,1.532818417,2.430610669
Business,WPD SWest,1.140313409,1.140313409,1.035511188,1.050130431,1.045426606,1.560731287,2.530425994
Business,WPD WestM,1.498961457,1.498961457,1.545708087,1.527569217,1.479451032,1.7347437,2.578137201
'));

coincidencePlot <- function (d) {
    library(grid);
    library(ggplot2);
    nms<-c('HV peaking probabilities', 'Primary peaking probabilities', 'BSP peaking probabilities', 'GSP peaking probabilities', 'Red time band', 'GSP Group peak', 'Customer group peak', '(load factor)', '(coincidence factor)', 'Non-DCP 227 charging basis', 'HV charging basis', 'Primary charging basis', 'BSP charging basis', 'GSP charging basis');
    nms<-factor(nms, levels=rev(nms));
    if (abs(d[3]-1)<1e-8) d[3] <- d[2]; # Scotland
    db <- 10*log10(c(d));
    p <- ggplot();
    pl <- list(p, p, p);
    p <- p + theme_minimal(base_size=16);
    p <- p + xlab('Peak-time consumption relative to average consumption (log scale)');
    p <- p + scale_x_continuous(limits=c(-0.5, 4.5), breaks=c(0, 10*log10(2)), labels=c('1x', '2x'));
    p <- p + theme(panel.grid.major.x=element_line(colour='blue', size=0.25, linetype=2));
    p <- p + theme(axis.text.y=element_text(size=16), legend.position="none");
    p <- p + theme(axis.ticks.y=element_blank(), axis.title.y=element_blank());
    p <- p + geom_vline(xintercept=0);
    db0<-c(0, db[1:6]);
    db0[db0==db[1:7]] <- NA;
    p <- p + geom_segment(mapping=aes(x=x, y=y, xend=x, yend=y), data=data.frame(y=nms, x=0));
    p <- p + geom_segment(mapping=aes(y=i, x=coin0, yend=i, xend=coin), data=data.frame(i=nms[1:5], coin=db[1:5], coin0=db0[1:5]), colour=c(rep('#000000', 4), '#FF0000'), arrow=arrow(length=unit(0.125, "in")), size=1.25);
    p <- p + geom_segment(mapping=aes(y=i, x=coin, yend=i1, xend=coin), data=data.frame(i=nms[1:4], coin=db[1:4], i1=nms[2:5]), size=0.6, colour=rep('#000000', 4));
    pl[[1]] <- p + scale_y_discrete(limits=nms[5:1]);

    p <- p + geom_segment(mapping=aes(y=i, x=coin, yend=i1, xend=coin), data=data.frame(i=nms[5], coin=db[5], i1=nms[6]), size=0.6, colour='#FF0000');
    p <- p + geom_segment(mapping=aes(x=0, y=i, xend=x, yend=i), data=data.frame(i=nms[6], x=db[6]), arrow=arrow(length=unit(0.125, "in")), size=0.25, colour='#009966');
    p <- p + geom_segment(mapping=aes(y=i, x=coin0, yend=i, xend=coin), data=data.frame(i=nms[6], coin=db[6], coin0=db0[6]), colour='#0000FF', arrow=arrow(length=unit(0.125, "in")), size=1.25);

    p <- p + geom_segment(mapping=aes(x=x, y=y, xend=xend, yend=y), data=data.frame(y=nms[8:9], xend=c(0, db[6]), x=rep(db[7], 2)), colour='#CCCCCC', size=1.25, arrow=arrow(length=unit(0.125, "in")));
    p <- p + geom_segment(mapping=aes(y=i, x=coin0, yend=i, xend=coin), data=data.frame(i=nms[7], coin=db[7], coin0=db0[7]), colour='#CCCCCC', arrow=arrow(length=unit(0.125, "in")), size=1.25);
    p <- p + geom_segment(mapping=aes(y=i, x=coin, yend=i1, xend=coin), data=data.frame(i=nms[7], coin=db[7], i1=nms[9]), size=0.6, colour='#CCCCCC');
    p <- p + geom_segment(mapping=aes(y=i, x=coin, yend=i1, xend=coin), data=data.frame(i=nms[6], coin=db[6], i1=nms[9]), size=0.6, colour='#0000FF');
    pl[[2]] <- p + scale_y_discrete(limits=nms[9:1]);

    p <- p + geom_segment(mapping=aes(y=i, x=coin, yend=i1, xend=coin), data=data.frame(i=nms[9], coin=db[6], i1=nms[10]), size=0.6, colour='#0000FF');
    p <- p + geom_segment(mapping=aes(x=x, y=i1, xend=x, yend=as.numeric(i2)+0.2), data=data.frame(i1=nms[1:4], i2=nms[11:14], x=db[1:4]), size=0.25, colour='#FF00FF');
    p <- p + geom_segment(mapping=aes(x=x, y=i1, xend=x, yend=i2), data=data.frame(i1=nms[6], i2=nms[10], x=db[6]), size=0.25, colour='#009966');
    p <- p + geom_segment(mapping=aes(x=x, y=i1, xend=x, yend=i2), data=data.frame(i1=rep(nms[6], 2), i2=rep(nms[14], 2), x=c(db[5], db[6])), size=0.25, colour='#0000FF');

    p <- p + geom_segment(mapping=aes(x=x, y=y, xend=xend, yend=y), data=data.frame(y=nms[10], x=0, xend=db[6]), arrow=arrow(length=unit(0.125, "in")), size=1.25, colour='#009966');
    p <- p + geom_segment(mapping=aes(x=x, y=y, xend=xend, yend=y), data=data.frame(y=as.numeric(nms[11:14])+0.2, xend=db[1:4], x=rep(0, 4)), arrow=arrow(length=unit(0.125, "in")), size=1.25, colour='#FF00FF');
    p <- p + geom_segment(mapping=aes(x=x, y=y, xend=xend, yend=y), data=data.frame(y=nms[11:14], x=rep(db[5], 4), xend=rep(db[6], 4)), arrow=arrow(length=unit(0.125, "in")), size=1.25, colour='#0000FF');
    pl[[3]] <- p;

    pl;
};

for (i in 1:length(dset1[, 1])) {
    pl <- coincidencePlot(as.matrix(dset1[i, 3:9]));
    png(paste(dset1[i, 2], dset1[i, 1], 'coincidences-1.png'), width=1800, height=1800/1.5*7/16, res=1800/10.0);
    print(pl[[1]]);
    graphics.off();
    png(paste(dset1[i, 2], dset1[i, 1], 'coincidences-2.png'), width=1800, height=1800/1.5*11/16, res=1800/10.0);
    print(pl[[2]]);
    graphics.off();
    exportGraph(pl[[3]], file.name=paste(dset1[i, 2], dset1[i, 1], 'coincidences-3'), file.type=1800);
}
EOR
}

1;
