#!/usr/bin/env perl
use warnings;
use strict;
use utf8;

my %rules = ( PerlModule => 'TopDown' );

my $volumetricUsage = [ 'Volumetric charge £/m3', 'Flow m3/year' ];

my $oxygenUsage = [ 'Oxygen charge £/kg', 'Oxygen demand kg/year' ];

my $sludgeUsage = [ 'Sludge charge £/kg', 'Suspended solids kg/year' ];

$rules{usageList} = [
    [
        'Sewers 300mm or more',
        $volumetricUsage,
        [ 'Operating expenses', 'Depreciation (IRC)', 'Return on capital', ]
    ],
    [
        'Combined sewers below 300mm',
        $volumetricUsage,
        [ 'Operating expenses', 'Depreciation (IRC)', 'Return on capital' ]
    ],
    [
        'Foul sewers below 300mm',
        $volumetricUsage,
        [ 'Operating expenses', 'Depreciation (IRC)', 'Return on capital' ]
    ],
    [
        'Drains below 300mm',
        $volumetricUsage,
        [ 'Operating expenses', 'Depreciation (IRC)', 'Return on capital' ]
    ],
    [
        'Reception', $volumetricUsage,
        [ 'Operating expenses', 'Depreciation', 'Return on capital' ]
    ],
    [
        'Wastewater settlement',
        $volumetricUsage,
        [ 'Depreciation', 'Return on capital' ]
    ],
    [ 'Biological oxidation', $oxygenUsage, ['Operating expenses'] ],
    [
        'Sludge transport',
        $sludgeUsage,
        [ 'Operating expenses', 'Depreciation', 'Return on capital' ]
    ],
    [
        'Sludge treatment',
        $sludgeUsage,
        [ 'Operating expenses', 'Depreciation', 'Return on capital' ]
    ],
    [
        'Sludge disposal',
        $sludgeUsage,
        [ 'Operating expenses', 'Depreciation', 'Return on capital' ]
    ],
];

my $property = [ 'Fixed charge £/property/year', 'Properties' ];

my $sewageVolume = [ 'Sewage charge £/m3', 'Sewage m3/year' ];

my $surfaceArea = [ 'Drainage charge £/m2/year', 'Drained area m2' ];

my $effluentVolume = [ 'Effluent volume charge £/m3', 'Effluent m3/year' ];
my $effluentOxygen = [ 'Effluent oxygen charge £/kg', 'Oxygen demand kg/year' ];
my $effluentSludge =
  [ 'Effluent sludge charge £/kg', 'Suspended solids kg/year' ];

$rules{tariffList} = [
    [ 'Domestic metered with surface water',    $sewageVolume, $property ],
    [ 'Domestic metered without surface water', $sewageVolume, $property ],
    [
        'Trade effluent (up to 225mm)', $effluentVolume,
        $effluentOxygen,                $effluentSludge
    ],
    [
        'Trade effluent (more than 225mm)',
        $effluentVolume, $effluentOxygen, $effluentSludge
    ],
    [ 'Surface water drainage', $surfaceArea ],
];

$rules{exemptDemandList} = [ [ 'Drained highway', $surfaceArea ], ];

$rules{routeingList} = [
    [ 'Flow m3 from effluent m3',          $volumetricUsage, $effluentVolume ],
    [ 'Oxygen kg from effluent oxygen kg', $oxygenUsage,     $effluentOxygen ],
    [ 'Sludge kg from effluent sludge kg', $sludgeUsage,     $effluentSludge ],
    [ 'Flow m3 from sewage m3',            $volumetricUsage, $sewageVolume ],
    [ 'Oxygen kg from sewage m3',          $oxygenUsage,     $sewageVolume ],
    [ 'Sludge kg from sewage m3',          $sludgeUsage,     $sewageVolume ],
    [ 'Flow m3/year from drained area m2', $volumetricUsage, $surfaceArea ],
    [ 'Oxygen kg/year from drained area m2', $oxygenUsage,     $surfaceArea ],
    [ 'Sludge kg/year from drained area m2', $sludgeUsage,     $surfaceArea ],
    [ 'Flow m3/year from property',          $volumetricUsage, $property ],
    [ 'Oxygen kg/year from property',        $oxygenUsage,     $property ],
    [ 'Sludge kg/year from property',        $sludgeUsage,     $property ],
];

use YAML;
YAML::DumpFile( "%-wastewater.yml.$$", \%rules );
rename "%-wastewater.yml.$$", '%-wastewater.yml';
