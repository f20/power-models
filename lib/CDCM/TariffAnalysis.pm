package CDCM;

# Copyright 2016 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub applyVolumesToTariff {
    my ( $model, $components, $tariff, $volumeData, $daysInYear, $name ) = @_;
    my @termsNoDays;
    my @termsWithDays;
    my %args = ( A900 => $daysInYear );
    my $i = 0;
    foreach (@$components) {
        ++$i;
        if (m#/day#) {
            push @termsWithDays, "A90$i*A$i";
        }
        else {
            push @termsNoDays, "A90$i*A$i";
        }
        $args{"A90$i"} = $tariff->{$_};
        $args{"A$i"}   = $volumeData->{$_};
    }
    Arithmetic(
        name => $name || 'Revenue (£)',
        arithmetic => '='
          . join( '+',
            @termsWithDays
            ? ( '0.01*A900*(' . join( '+', @termsWithDays ) . ')' )
            : (),
            @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
            : ('0'),
          ),
        arguments     => \%args,
        defaultFormat => '0soft',
    );
}

sub unroundedTariffAnalysis {

    my ( $model, $allComponents, $allTariffsByEndUser,
        $componentLabelset, $daysInYear, $tariffsExMatching, @matchingTables, )
      = @_;

    my @utaTables;
    foreach my $cat ( 'Asset', 'Transmission exit', 'Other expenditure' ) {
        my $cols = $tariffsExMatching->{ $allComponents->[0] }{source}{cols};
        my $con  = Constant(
            name => 'Levels containing ' . lc($cat) . ' charges',
            cols => $cols,
            data => $cat =~ /asset/i
            ? [ map { /asset/i ? 1 : 0 } @{ $cols->{list} } ]
            : $cat =~ /exit/i
            ? [ map { /exit/i ? 1 : 0 } @{ $cols->{list} } ]
            : [ map { /operating/i ? 1 : 0 } @{ $cols->{list} } ],
        );
        my %hash = ( name => $cat );
        $hash{$_} = SumProduct(
            name   => "$cat contributions to " . lcfirst($_),
            vector => $con,
            matrix => $tariffsExMatching->{$_}{source},
        ) foreach @$allComponents;
        push @utaTables, \%hash;
        Columnset(
            name    => "Unrounded tariff analysis: $cat charges",
            columns => [ @hash{@$allComponents} ],
        );
    }

    {
        my %hash = ( name => 'Matching' );
        foreach my $tariffComponent (@$allComponents) {
            my @tables;
            push @tables, $_ foreach grep { $_ }
              map { $_->{$tariffComponent} } @matchingTables;
            $hash{$tariffComponent} =
              @tables
              ? Arithmetic(
                name => 'Matching contributions to '
                  . lcfirst($tariffComponent),
                arithmetic => '=' . join( '+', map { "A$_" } 1 .. @tables ),
                arguments =>
                  { map { ( "A$_" => $tables[ $_ - 1 ] ); } 1 .. @tables },
              )
              : Constant(
                name => 'Matching contributions to '
                  . lcfirst($tariffComponent),
                rows => $allTariffsByEndUser,
                cols => $componentLabelset->{$tariffComponent},
                data => [ [] ]
              );
        }
        push @utaTables, \%hash;
        Columnset(
            name    => 'Unrounded tariff analysis: Matching charges',
            columns => [ @hash{@$allComponents} ],
        );
    }

    @utaTables;

}

1;
