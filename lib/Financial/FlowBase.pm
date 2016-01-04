package Financial::FlowBase;

=head Copyright licence and disclaimer

Copyright 2015, 2016 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my ( $class, %args ) = @_;
    die __PACKAGE__ . ' needs a model attribute' unless $args{model};
    $args{lines}           ||= 4;
    $args{name}            ||= 'XYZ';
    $args{show_balance}    ||= 'xyz balance (£)';
    $args{show_buffer}     ||= 'xyz cash buffer (£)';
    $args{show_flow}       ||= 'xyz flow (£)';
    $args{show_formatBase} ||= '0';
    $args{show_item}       ||= 'Item ';
    bless \%args, $class;
}

sub prefix_calc {
    my ($flow) = @_;
    $flow->{is_cost} ? '=-1*' : '=';
}

sub finish {
    die 'Not implemented in ' . __PACKAGE__;
}

sub labelset {
    my ($flow) = @_;
    $flow->{labelset} ||= Labelset(
        editable => (
            $flow->{database}{names} ||= Dataset(
                name          => 'Item name',
                defaultFormat => 'texthard',
                rows          => $flow->labelsetNoNames,
                data          => [ map { '' } $flow->labelsetNoNames->indices ],
            )
        ),
    );
}

sub labelsetNoNames {
    my ($flow) = @_;
    $flow->{labelsetNoNames} ||= Labelset(
        name          => 'Items without names',
        defaultFormat => 'thitem',
        list          => [ 1 .. $flow->{lines} ]
    );
}

sub startDate {
    my ($flow) = @_;
    $flow->{database}{startDate} ||= Dataset(
        name          => 'Start date',
        defaultFormat => 'datehard',
        rows          => $flow->labelsetNoNames,
        data          => [ map { '' } @{ $flow->labelsetNoNames->{list} } ],
    );
}

sub endDate {
    my ($flow) = @_;
    $flow->{database}{endDate} ||= Dataset(
        name          => 'End date',
        defaultFormat => 'datehard',
        rows          => $flow->labelsetNoNames,
        data          => [ map { '' } @{ $flow->labelsetNoNames->{list} } ],
    );
}

sub averageDays {
    my ($flow) = @_;
    $flow->{database}{averageDays} ||= Dataset(
        name          => 'Average credit and inventory days',
        defaultFormat => '0.0hard',
        rows          => $flow->labelsetNoNames,
        data          => [ map { 0; } @{ $flow->labelsetNoNames->{list} } ],
    );
}

sub worstDays {
    my ($flow) = @_;
    my $worstDays  = $flow->{is_cost} ? 'minDays' : 'maxDays';
    my $worstLabel = $flow->{is_cost} ? 'Lowest'  : 'Highest';
    $flow->{database}{$worstDays} ||= Dataset(
        name          => $worstLabel . ' credit and inventory days',
        defaultFormat => '0.0hard',
        rows          => $flow->labelsetNoNames,
        data          => [ map { 0; } @{ $flow->labelsetNoNames->{list} } ],
    );
}

sub stream {
    die "$_[0]->stream not implemented in " . __PACKAGE__;
}

sub balance {
    die 'Not implemented in ' . __PACKAGE__;
}

sub buffer {
    die 'Not implemented in ' . __PACKAGE__;
}

1;

