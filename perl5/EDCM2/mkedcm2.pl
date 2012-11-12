#!/usr/bin/perl

=head Copyright licence and disclaimer

Copyright 2011-2012 Reckon LLP and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted
provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of
conditions and the following disclaimer in the documentation and/or other materials provided
with the distribution.

THIS SOFTWARE IS PROVIDED BY RECKON LLP AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED
WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL RECKON LLP OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

# Suggestion for debugging:
# perl perl5/EDCM2/mkedcm2.pl -xlsx -single  2>&1 | more -F

use warnings;
use strict;
use utf8;
require Carp;
$SIG{__DIE__} = \&Carp::confess;
use File::Spec::Functions qw(rel2abs abs2rel catfile updir catdir);
use File::Basename 'dirname';
my $perl5dir;

BEGIN {  # double dirname - this script is expected to be perl5/EDCM2/mkedcm2.pl
    $perl5dir =
      dirname dirname( rel2abs( -l $0 ? ( readlink $0, dirname $0) : $0 ) );
}

use lib ( $perl5dir, catdir( dirname($perl5dir), 'cpan' ) );
use Ancillary::Manufacturing;
my $maker = Ancillary::Manufacturing->factory;

my ($illustrative) = grep { /^-+illustrative/i } @ARGV;    # -illustrative
my ($randomise)    = grep { /^-+random/i } @ARGV;          # -random, -randomcut
my ($large)  = grep { /^-+large/i } @ARGV;     # -large (even when data are not)
my ($small)  = grep { /^-+small/i } @ARGV;     # -small
my ($rev)    = grep { /^-+rev/i } @ARGV;       # -rev
my ($huge)   = grep { /^-+huge/i } @ARGV;      # -huge
my ($medium) = grep { /^-+medium/i } @ARGV;    # -medium

# -ldnoyes, -ldnotar, -ldnoyes5, -ldnotar5
my ($ldno) = grep { /^-+ldno/i } @ARGV;
$ldno = 'ldnorev5' unless $illustrative || $randomise || $ldno;

my ($dcp130)      = grep { /-+DCP130/i } @ARGV;
my ($dcp139)      = grep { /-+DCP139/i } @ARGV;
my ($dgcondition) = grep { /-+DGCONDITION/i } @ARGV;

my $run = ( grep { /^-+single/i } @ARGV ) ? 'run' : 'runParallel';
$maker->{useXLSX}->() if grep { /^-+xlsx/i } @ARGV;

my @companies = grep { /\.yml$/ } @ARGV;
@companies = catfile( dirname($perl5dir), qw(data extra Blank.yml) )
  unless @companies;

foreach my $company (@companies) {
    foreach my $power qw(FCP LRIC) {
        next
          if $company =~ /(^|\/)(CN|WPD.*M|SSE|SP)/
          and $power  =~ /LRIC/;
        next
          if $company =~ /(^|\/)(CE|NP|ENW|UKPN|WPD[- ](?:S.*W|Wales|West\.))/
          and $power  =~ /FCP/;
        open my $dh, '<', $company;
        $maker->{processStream}->( $dh, abs2rel($company) );
        $maker->{processRuleset}->(
            {
                template   => "%-$power",
                PerlModule => 'EDCM2',
                $large || $company =~ /Blank/
                ? (
                    $small
                    ? (
                        numTariffs   => 1,
                        numLocations => 1,
                        illustrative => 1,
                      )
                    : $medium ? (
                        numTariffs   => 200,
                        numLocations => 1200,
                      )
                    : $huge ? (
                        numTariffs => $power =~ /LRIC/ ? 650 : 800,
                        numLocations => 5000,
                      )
                    : (
                        numTariffs   => 300,
                        numLocations => 1200,
                    ),
                  )
                : $small ? ( small => $small )
                : (),
                method     => $power,
                ldnoRev    => $ldno,
                noNegative => 1,
                protect    => 1,
                summaries  => 1,
                validation => 'lenientnomsg',
                $randomise || $illustrative
                ? ( illustrative => 1 )
                : (),
                $randomise   ? ( randomise   => $randomise ) : (),
                $dcp130      ? ( DCP130      => 1 )          : (),
                $dcp139      ? ( DCP139      => 1 )          : (),
                $dgcondition ? ( DGCONDITION => 1 )          : (),
            }
        );

    }
}

$maker->{validate}->(
    $perl5dir,
    $rev
    ? 'DBI:SQLite:dbname=/Reckon/projects/EDCM general model development/O_Drafting/EDCM revision number database.sqlite'
    : undef
);
$maker->{$run}->( $maker->{prepare}->( $maker->{listMatching}->() ) );

