﻿package Elec::CustomersTyped;

# Copyright 2016-2019 Franck Latrémolière, Reckon LLP and others.
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
use base 'Elec::Customers';
use SpreadsheetModel::Shortcuts ':all';

sub volumeDataColumn {
    my ( $self, $component ) = @_;
    [
        map {
            $component =~ m#kWh#    # active units
              ? ( /\[standing\]/ ? undef : 0 )
              : $component =~ m#kVA# ? (    # capacity or reactive
                /Medium|Non-CT|Unrestricted|Two Rate|UMS|Unmetered|Aggregated/
                ? undef
                : 0
              )
              : (                           # fixed charges
                /\[.*units\]|UMS|Unmetered/ ? undef : 0
              );
        } @{ $self->userLabelset->{list} }
    ];
}

sub userLabelsetRegrouped {
    my ($self) = @_;
    return $self->{userLabelsetRegrouped}
      if $self->{userLabelsetRegrouped};
    my $userLabelset = $self->userLabelset;
    my @groupNameList;
    my %groupMembers;
    foreach ( @{ $userLabelset->{groups} } ) {
        my $groupName = $_->{name};
        push @groupNameList, $groupName;
        $groupMembers{$groupName} ||= [];
        foreach ( @{ $_->{list} } ) {
            my $group =
              /^(.*?)[^a-zA-Z0-9]*\[.*\][^a-zA-Z0-9]*$/s ? $1 : $groupName;
            push @groupNameList, $group unless $groupMembers{$group};
            push @{ $groupMembers{$group} }, $_;
        }
    }
    $self->{userLabelsetRegrouped} = Labelset(
        name   => 'Tariffs regrouped',
        groups => [
            map { Labelset( name => $_, list => $groupMembers{$_} ); }
              @groupNameList
        ],
    );
}

1;
