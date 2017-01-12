package Financial::Inputs;

=head Copyright licence and disclaimer

Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.

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

sub new {
    my ( $class, %args ) = @_;
    die __PACKAGE__ . ' needs a model attribute' unless $args{model};
    bless \%args, $class;
}

sub sales {
    my ($input) = @_;
    (
        lines  => $input->{model}{numSales},
        name   => 'Sales',
        number => 1430,
    );
}

sub costSales {
    my ($input) = @_;
    return unless $input->{model}{numCostSales};
    (
        lines  => $input->{model}{numCostSales},
        name   => 'Cost of sales',
        number => 1440,
    );
}

sub adminExp {
    my ($input) = @_;
    (
        lines  => $input->{model}{numAdminExp},
        name   => 'Administrative expenses',
        number => 1442,
    );
}

sub exceptional {
    my ($input) = @_;
    return unless $input->{model}{numExceptional};
    (
        lines  => $input->{model}{numExceptional},
        name   => 'Exceptional costs',
        number => 1444,
    );
}

sub capitalExp {
    my ($input) = @_;
    return unless $input->{model}{numCapitalExp};
    (
        lines  => $input->{model}{numCapitalExp},
        name   => 'Capital expenditure',
        number => 1447,
    );
}

sub assets {
    my ($input) = @_;
    (
        lines  => $input->{model}{numAssets},
        name   => 'Fixed assets',
        number => 1450,
    );
}

sub debt {
    my ($input) = @_;
    (
        lines  => $input->{model}{numDebt},
        name   => 'Borrowings',
        number => 1460,
    );
}

1;
