#!/usr/bin/perl -w
use warnings;
use strict;
use YAML;

open my $f, '<', shift;
local undef $/;

my $d = Load <$f>;

$d->{1000}[0]{_note} =
  'Illustrative data based on February 2012 CDCM model 100.';
$d->{1000}[1]{'Company charging year data version'} ||= '#VALUE!';
$d->{1000}[2]{'Company charging year data version'} = 'Illustrative';
$d->{1000}[3]{'Company charging year data version'} =
  'Spare capacity prototype';

foreach ( 1 .. 8 ) {
    my $col = $d->{1025}[$_];
    $col->{'LV Agg WC Domestic'}     = $col->{'Domestic Unrestricted'};
    $col->{'LV Agg WC Non-Domestic'} =
      $col->{'Small Non Domestic Unrestricted'};
    $col->{'LV Agg CT Non-Domestic'}    = $col->{'LV Medium Non-Domestic'};
    $col->{'LV Medium Non-Domestic WC'} =
      $col->{'Small Non Domestic Unrestricted'};
    $col->{'LV HH WC'} = $col->{'Small Non Domestic Unrestricted'};
}

foreach ( 1 .. 2 ) {
    my $col = $d->{1041}[$_];
    $col->{'LV Agg WC Domestic'}     = $col->{'Domestic Unrestricted'};
    $col->{'LV Agg WC Non-Domestic'} =
      $col->{'Small Non Domestic Unrestricted'};
    $col->{'LV Agg CT Non-Domestic'}    = $col->{'LV Medium Non-Domestic'};
    $col->{'LV Medium Non-Domestic WC'} = $col->{'LV Medium Non-Domestic'};
    $col->{'LV HH WC'}                  = $col->{'LV HH Metered'};
}

foreach ( 1 .. 6 ) {
    my $col = $d->{1053}[$_];
    foreach my $prefix ( '', 'LDNO LV ', 'LDNO HV ' ) {
        $col->{ $prefix . 'LV Agg WC Domestic' }        = '';
        $col->{ $prefix . 'LV Agg WC Non-Domestic' }    = '';
        $col->{ $prefix . 'LV Agg CT Non-Domestic' }    = '';
        $col->{ $prefix . 'LV Medium Non-Domestic WC' } = '';
        $col->{ $prefix . 'LV HH WC' }                  = '';
    }
}

foreach ( 1 .. 3 ) {
    my $col = $d->{1061}[$_];
    $col->{'Domestic Unrestricted'}           ||= '#VALUE!';
    $col->{'Small Non Domestic Unrestricted'} ||= '#VALUE!';
    $col->{'NHH UMS'}                         ||= '#VALUE!';
    $col->{'LV Medium Non-Domestic WC'} = $col->{'LV Medium Non-Domestic'};
}

foreach ( 1 .. 3 ) {
    my $col = $d->{1062}[$_];
    $col->{'LV Medium Non-Domestic WC'} = $col->{'LV Medium Non-Domestic'};
}

print Dump $d;
