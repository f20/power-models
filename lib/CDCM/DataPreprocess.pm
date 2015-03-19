package CDCM;

=head Copyright licence and disclaimer

Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.

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

# This file is out of the scope of revision numbers.

use warnings;
use strict;
use utf8;

sub _override {
    my ( $table, $override ) = @_;
    return unless ref $table eq 'ARRAY' && ref $override eq 'ARRAY';
    for ( my $c = 0 ; $c < @$override ; ++$c ) {
        next unless ref $override->[$c] eq 'HASH';
        while ( my ( $k, $v ) = each %{ $override->[$c] } ) {
            $table->[$c]{$k} = $v;
        }
    }
}

sub preprocessDataset {

    my ($model) = @_;

    my $d = $model->{dataset};

    $d->{1000}[3]{'Company charging year data version'} = $model->{version}
      if $model->{version};

    if ( $model->{dcp133} ) {
        $d->{1000}[3]{'Company charging year data version'} .= ' (DCP 133)';
        _override( $d->{$_}, $d->{ $_ . 'dcp133' } )
          foreach $model->{dcp133} =~ /nodivers/i ? () : qw(1017),
          qw(1018 1020);
    }

    if ( $model->{unauth} && $model->{unauth} =~ /day/ ) {
        my $vd = $d->{1053};
        if ( $vd and !$vd->[6]{_column} || $vd->[6]{_column} !~ /exceed/i ) {
            splice @$vd, 6, 0, { map { ( $_ => '' ); } keys %{ $vd->[5] } };
            my $add = $model->{unauth} =~ /add/i;
            while ( my ( $t, $p ) = each %{ $vd->[5] } ) {
                next unless $t =~ s/ exceedprop$//s;
                $vd->[6]{$t} = $vd->[5]{$t} * $p;
                $vd->[5]{$t} -= $vd->[6]{$t} unless $add;
            }
        }
    }

    if ( $model->{addVolumes} && $model->{addVolumes} =~ /matching/i ) {
        $d->{1054} ||= $d->{1053};
    }

    if (   $model->{targetRevenue}
        && $model->{targetRevenue} =~ /single/i
        && $d->{1076}[4] )
    {
        foreach ( grep { !/^_/ } keys %{ $d->{1076}[1] } ) {
            $d->{1076}[1]{$_} += $d->{1076}[2]{$_} if $d->{1076}[2]{$_};
            $d->{1076}[2]{$_} = '';
            $d->{1076}[1]{$_} += $d->{1076}[3]{$_} if $d->{1076}[3]{$_};
            $d->{1076}[3]{$_} = '';
            $d->{1076}[1]{$_} -= $d->{1076}[4]{$_} if $d->{1076}[4]{$_};
            $d->{1076}[4]{$_} = '';
        }
    }

    if (   $model->{targetRevenue}
        && $model->{targetRevenue} =~ /DCP132/i
        && !$d->{1001}
        && $d->{1076}[1]
        && ( my ($row1076) = grep { !/^_/ } keys %{ $d->{1076}[1] } ) )
    {
        require YAML;
        $d->{1001} = YAML::Load(<<EOY);
---
- _table: 1001. CDCM target revenue data and calculations
- nothing to see: here
- 1 Revenue raised outside CDCM EDCM and Certain Interconnector Revenue: ''
  2 Voluntary under-recovery: ''
  3 Revenue raised outside CDCM please provide description if used: ''
  4 Revenue raised outside CDCM please provide description if used: ''
  Allowed Pass-Through Items: PTt
  Base Demand Revenue: BRt
  Base Demand Revenue Before Inflation: PUt
  Connection Guaranteed Standards Systems Processes penalty: 'CGSRAt, CGSSPt & AUMt'
  Correction Factor: -Kt
  Incentive Revenue and Other Adjustments: ''
  Incentive Revenue for Distributed Generation: IGt
  Innovation Funding Incentive Adjustment: IFIt
  Latest Forecast of CDCM Revenue: ''
  Losses Incentive 1: UILt
  Losses Incentive 2: PCOLt
  Losses Incentive 3: -COLt
  Losses Incentive 4: PPLt
  Low Carbon Network Fund 1: LCN1t
  Low Carbon Network Fund 2: LCN2t
  Low Carbon Network Fund 3: LCN3t
  Merger Adjustment: MGt
  Other 1 Excluded services Top-up standby and enhanced system security: ES4
  Other 2 Excluded services Revenue protection services: ES5
  Other 3 Excluded services Miscellaneous: ES7
  Other 4 please provide description if used: ''
  Other 5 please provide description if used: ''
  Pass-Through Business Rates: RBt
  Pass-Through Licence Fees: LFt
  Pass-Through Others: 'MPTt, HBt, IEDt'
  Pass-Through Price Control Reopener: UNCt
  Pass-Through Transmission Exit: TBt
  Quality of Service Incentive Adjustment: IQt
  RPI Indexation Factor: PIADt
  Tax Trigger Mechanism Adjustment: CTRAt
  Total Allowed Revenue: ARt
  Total Other Revenue to be Recovered by Use of System Charges: ''
  Total Revenue for Use of System Charges: ''
  Total Revenue to be raised outside the CDCM: ''
  Transmission Connection Point Charges Incentive Adjustment: ITt
  _column: Term
- 1 Revenue raised outside CDCM EDCM and Certain Interconnector Revenue: ''
  2 Voluntary under-recovery: ''
  3 Revenue raised outside CDCM please provide description if used: ''
  4 Revenue raised outside CDCM please provide description if used: ''
  Allowed Pass-Through Items: CRC3
  Base Demand Revenue: CRC3
  Base Demand Revenue Before Inflation: CRC3
  Connection Guaranteed Standards Systems Processes penalty: CRC12
  Correction Factor: CRC3
  Incentive Revenue and Other Adjustments: ''
  Incentive Revenue for Distributed Generation: CRC11
  Innovation Funding Incentive Adjustment: CRC10
  Latest Forecast of CDCM Revenue: ''
  Losses Incentive 1: CRC7
  Losses Incentive 2: CRC7
  Losses Incentive 3: CRC7
  Losses Incentive 4: CRC7
  Low Carbon Network Fund 1: CRC13
  Low Carbon Network Fund 2: CRC13
  Low Carbon Network Fund 3: CRC13
  Merger Adjustment: CRC3
  Other 1 Excluded services Top-up standby and enhanced system security: CRC15
  Other 2 Excluded services Revenue protection services: CRC15
  Other 3 Excluded services Miscellaneous: CRC15
  Other 4 please provide description if used: ''
  Other 5 please provide description if used: ''
  Pass-Through Business Rates: CRC4
  Pass-Through Licence Fees: CRC4
  Pass-Through Others: CRC4
  Pass-Through Price Control Reopener: CRC4
  Pass-Through Transmission Exit: CRC4
  Quality of Service Incentive Adjustment: CRC8
  RPI Indexation Factor: CRC3
  Tax Trigger Mechanism Adjustment: CRC3
  Total Allowed Revenue: ''
  Total Other Revenue to be Recovered by Use of System Charges: ''
  Total Revenue for Use of System Charges: ''
  Total Revenue to be raised outside the CDCM: ''
  Transmission Connection Point Charges Incentive Adjustment: CRC9
  _column: CRC
- 1 Revenue raised outside CDCM EDCM and Certain Interconnector Revenue: ''
  2 Voluntary under-recovery: ''
  3 Revenue raised outside CDCM please provide description if used: ''
  4 Revenue raised outside CDCM please provide description if used: ''
  Allowed Pass-Through Items: ''
  Base Demand Revenue: ''
  Base Demand Revenue Before Inflation: ''
  Connection Guaranteed Standards Systems Processes penalty: ''
  Correction Factor: ''
  Incentive Revenue and Other Adjustments: ''
  Incentive Revenue for Distributed Generation: ''
  Innovation Funding Incentive Adjustment: ''
  Latest Forecast of CDCM Revenue: ''
  Losses Incentive 1: ''
  Losses Incentive 2: ''
  Losses Incentive 3: ''
  Losses Incentive 4: ''
  Low Carbon Network Fund 1: ''
  Low Carbon Network Fund 2: ''
  Low Carbon Network Fund 3: ''
  Merger Adjustment: ''
  Other 1 Excluded services Top-up standby and enhanced system security: ''
  Other 2 Excluded services Revenue protection services: ''
  Other 3 Excluded services Miscellaneous: ''
  Other 4 please provide description if used: ''
  Other 5 please provide description if used: ''
  Pass-Through Business Rates: ''
  Pass-Through Licence Fees: ''
  Pass-Through Others: ''
  Pass-Through Price Control Reopener: ''
  Pass-Through Transmission Exit: ''
  Quality of Service Incentive Adjustment: ''
  RPI Indexation Factor: 1
  Tax Trigger Mechanism Adjustment: ''
  Total Allowed Revenue: ''
  Total Other Revenue to be Recovered by Use of System Charges: ''
  Total Revenue for Use of System Charges: ''
  Total Revenue to be raised outside the CDCM: ''
  Transmission Connection Point Charges Incentive Adjustment: ''
  _column: Value
- _column: Revenue elements and subtotals (£/year)
EOY
        $d->{1001}[4]{'Base Demand Revenue Before Inflation'} =
          $d->{1076}[1]{$row1076} || '';
        $d->{1001}[4]{'Pass-Through Others'} = $d->{1076}[2]{$row1076} || '';
        $d->{1001}[4]{'Correction Factor'}   = $d->{1076}[3]{$row1076} || '';
        $d->{1001}[4]{
'1 Revenue raised outside CDCM EDCM and Certain Interconnector Revenue'
          } = $d->{1076}[4]{$row1076}
          || '';
    }

    if (   $d->{1001}
        && $model->{targetRevenue}
        && $model->{targetRevenue} =~ /DCP132/i )
    {
        $d->{1001}[0] = sub {
            my ($k) = @_;
            $k =~ s/ see note 1//i;
            $k =~ s/( [A-Z][0-9]?)+$//;
            grep { substr( $_, 0, length $k ) eq $k; } keys %{ $d->{1001}[2] };
        };
    }

    if ( $d->{1001} && $model->{revenueAdj} ) {
        my $adj = $model->{revenueAdj};
        if ( $adj < 100 && $d->{1053} ) {
            my $units = 0;
            foreach my $cust ( keys %{ $d->{1053}[1] } ) {
                next if $cust eq '_column';
                my $u = 0;
                $u += $_
                  foreach grep { $_; } map { $d->{1053}[$_]{$cust} } 1 .. 3;
                $cust =~ /gener/i ? $units -= $u : $units += $u;
            }
            $adj *= 10 * $units;
        }
        $d->{1001}[4]{'Pass-Through Others'} ||= 0;
        $d->{1001}[4]{'Pass-Through Others'} += $adj;
    }

    if ( $d->{1001} ) {
        foreach my $root ( '3', '4', 'Other 4', 'Other 5' ) {
            if ( my ($k) = grep { /^$root/ } keys %{ $d->{1001}[4] } ) {
                my $nk = $root;
                $nk .= ' Revenue raised outside CDCM' if $root =~ /^[34]/;
                $d->{1001}[$_]{$nk} ||= $d->{1001}[$_]{$k}
                  foreach grep { $d->{1001}[$_]; } 1 .. 6;
                $k =~ s/^$root[. -]*//;
                $k =~ s/^Revenue raised outside CDCM[. -]*//
                  if $root =~ /^[34]/;
                $d->{1001}[2]{$nk} =
                  ( $d->{1001}[2]{$nk} || '' ) . ucfirst($k);
            }
        }
    }

    if ( $model->{tariffs} =~ /dcp137/i ) {
        $d->{1028}[0] = sub {
            my ($key) = @_;
            $key =~ /(.*) \S+ GDA$/i;
          }
          if $d->{1028};
        $d->{1053}[0] = sub {
            local $_ = $_[0];
            / \S+ GDA$/i ? '' : ();
          }
          if $d->{1053};
    }

# Below is mostly for DCP 179 but applied to all cases to help multi-model manufacturing.

    if (  !exists $d->{1025}[1]{'LV Network Non-Domestic Non-CT'}
        && exists $d->{1025}[1]{'LV HH Metered'} )
    {
        foreach ( 1 .. 8 ) {
            my $col = $d->{1025}[$_];
            $col->{'LV Network Domestic'} = $col->{'Domestic Unrestricted'};
            $col->{'LV Network Non-Domestic Non-CT'} = $col->{
                $model->{tariffs} =~ /dcp179bare/i
                ? 'LV Medium Non-Domestic'
                : 'Small Non Domestic Unrestricted'
            };
            $col->{'LV Network Non-Domestic CT'} = $col->{'LV HH Metered'};
            $col->{'LV Sub Non-CT'} = $col->{'LV Sub Medium Non-Domestic'};
            $col->{'LV Sub CT'}     = $col->{'LV Sub HH Metered'};
        }
    }

    if (  !exists $d->{1025}[1]{'LV Network Domestic'}
        && exists $d->{1025}[1]{'Domestic Unrestricted'} )
    {
        foreach ( 1 .. 8 ) {
            my $col = $d->{1025}[$_];
            $col->{'LV Network Domestic'} = $col->{'Domestic Unrestricted'};
        }
    }

    if (  !exists $d->{1025}[1]{'LV Network Non-Domestic Non-CT'}
        && exists $d->{1025}[1]{'Small Non Domestic Unrestricted'} )
    {
        foreach ( 1 .. 8 ) {
            my $col = $d->{1025}[$_];
            $col->{'LV Network Non-Domestic Non-CT'} =
              $col->{'Small Non Domestic Unrestricted'};
        }
    }

    $d->{1042}[1] = $d->{1041}[2]
      if $model->{impliedLoadFactors}
      && $model->{impliedLoadFactors} =~ /input/i
      && !$d->{1042}[1]
      && $d->{1041}
      && $d->{1041}[2];

    if (  !exists $d->{1041}[1]{'LV Network Non-Domestic Non-CT'}
        && exists $d->{1041}[1]{'Domestic Unrestricted'} )
    {
        foreach ( 1 .. 2 ) {
            my $col = $d->{1041}[$_];
            $col->{'LV Network Domestic'} ||= $col->{'Domestic Unrestricted'};
            $col->{'LV Network Non-Domestic Non-CT'} ||=
              $col->{'Small Non Domestic Unrestricted'};
            $col->{'LV Network Non-Domestic CT'} ||= $col->{'LV HH Metered'};
            $col->{'LV Sub Non-CT'} ||= $col->{'LV Sub Medium Non-Domestic'};
            $col->{'LV Sub CT'}     ||= $col->{'LV Sub HH Metered'};
            $col->{'HV Network Non-CT'} ||= $col->{'HV Medium Non-Domestic'};
            $col->{'HV Network CT'}     ||= $col->{'HV HH Metered'};
        }
        foreach ( 1 .. 6 ) {
            my $col = $d->{1053}[$_];
            foreach my $prefix ( '', 'LDNO LV ', 'LDNO HV ' ) {
                $col->{ $prefix . 'LV Network Domestic' }            ||= '';
                $col->{ $prefix . 'LV Network Non-Domestic Non-CT' } ||= '';
                $col->{ $prefix . 'LV Network Non-Domestic CT' } ||=
                  $col->{ $prefix . 'LV HH Metered' };
                $col->{ $prefix . 'LV Sub Non-CT' } ||= '';
                $col->{ $prefix . 'LV Sub CT' } ||=
                  $col->{ $prefix . 'LV Sub HH Metered' };
                $col->{ $prefix . 'HV Network Non-CT' } ||= '';
                $col->{ $prefix . 'HV Network CT' } ||=
                  $col->{ $prefix . 'HV HH Metered' };
            }
        }
    }

    if (  !exists $d->{1061}[1]{'Domestic Unrestricted'}
        && exists $d->{1061}[1]{'Domestic Two Rate'} )
    {
        0 and $model->addModifiedWarning;
        foreach my $t ( 'Domestic', 'Small Non Domestic' ) {
            foreach ( 1, 2, 3 ) {
                $d->{1061}[$_]{ $t . ' Unrestricted' } =
                  0.85 * $d->{1061}[$_]{ $t . ' Two Rate' } +
                  0.15 * ( $d->{1062}[$_]{ $t . ' Two Rate' } || 0 );
            }
        }
    }

    if (  !exists $d->{1025}[1]{'LV Generation NHH or Aggregate HH'}
        && exists $d->{1025}[1]{'LV Generation NHH'} )
    {
        foreach ( 1 .. 8 ) {
            my $col = $d->{1025}[$_];
            $col->{'LV Generation NHH or Aggregate HH'} ||=
              $col->{'LV Generation NHH'};
        }
    }
    elsif ( !exists $d->{1025}[1]{'LV Generation NHH'}
        && exists $d->{1025}[1]{'LV Generation NHH or Aggregate HH'} )
    {
        foreach ( 1 .. 8 ) {
            my $col = $d->{1025}[$_];
            $col->{'LV Generation NHH'} ||=
              $col->{'LV Generation NHH or Aggregate HH'};
        }
    }

    if (  !exists $d->{1053}[1]{'LV Generation NHH or Aggregate HH'}
        && exists $d->{1053}[1]{'LV Generation NHH'} )
    {
        foreach ( 1 .. 6 ) {
            my $col = $d->{1053}[$_];
            foreach my $prefix ( '', 'LDNO LV ', 'LDNO HV ' ) {
                $col->{ $prefix . 'LV Generation NHH or Aggregate HH' } ||=
                  $col->{ $prefix . 'LV Generation NHH' };
            }
        }
    }

    if ( $d->{p300} ) {
        foreach ( 'Domestic', 'Small Non Domestic' ) {
            $model->addModifiedWarning;
            my $hhTariffName =
              $_ eq 'Domestic'
              ? 'LV Network Domestic'
              : 'LV Network Non-Domestic Non-CT';
            my $units1 = $d->{1053}[1]{"$_ Unrestricted"};
            my $units0 = $d->{1053}[1]{"$_ Off Peak related MPAN"};
            my $units2 =
              $d->{1053}[1]{"$_ Two Rate"} + $d->{1053}[2]{"$_ Two Rate"};
            $d->{1041}[2]{$hhTariffName} =
              ( $units1 + $units2 + $units0 ) /
              ( $units1 / $d->{1041}[2]{"$_ Unrestricted"} +
                  $units2 / $d->{1041}[2]{"$_ Two Rate"} );
            $d->{1041}[1]{$hhTariffName} =
              $d->{1041}[2]{$hhTariffName} *
              ( $units1 * $d->{1041}[1]{"$_ Unrestricted"} /
                  $d->{1041}[2]{"$_ Unrestricted"} +
                  $units2 * $d->{1041}[1]{"$_ Two Rate"} /
                  $d->{1041}[2]{"$_ Two Rate"} ) /
              ( $units1 + $units2 + $units0 );
            if ( my $prop = $d->{p300}[1]{$_} ) {

                foreach my $rate ( 1 .. 3 ) {
                    $d->{1053}[$rate]{$hhTariffName} =
                      $prop *
                      ( $d->{1053}[1]{"$_ Unrestricted"} *
                          $d->{1061}[$rate]{"$_ Unrestricted"} +
                          $d->{1053}[1]{"$_ Off Peak related MPAN"} *
                          $d->{1061}[$rate]{"$_ Off Peak related MPAN"} +
                          $d->{1053}[1]{"$_ Two Rate"} *
                          $d->{1061}[$rate]{"$_ Two Rate"} +
                          $d->{1053}[2]{"$_ Two Rate"} *
                          ( $d->{1062}[$rate]{"$_ Two Rate"} || 0 ) );
                }
                $d->{1053}[4]{$hhTariffName} =
                  $prop *
                  ( $d->{1053}[4]{"$_ Unrestricted"} +
                      $d->{1053}[4]{"$_ Two Rate"} );
                foreach my $c ( 1, 4 ) {
                    $d->{1053}[$c]{$_} *= 1 - $prop
                      foreach grep { $d->{1053}[$c]{$_} } "$_ Unrestricted",
                      "$_ Off Peak related MPAN";
                }
                foreach my $c ( 1, 2, 4 ) {
                    $d->{1053}[$c]{"$_ Two Rate"} *= 1 - $prop;
                }
            }
        }
    }

}

sub addModifiedWarning {
    my ($model) = @_;
    return if $model->{datasetModifiedWarning};
    $model->{datasetModifiedWarning} = 1;
    $model->{dataset}{1000}[3]{'Company charging year data version'} .=
      ' (modified)';
}

1;
