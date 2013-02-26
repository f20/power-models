package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2013 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Shortcuts 'Notes';

sub generalNotes {
    my ($model) = @_;
    Notes(
        name  => 'Overview',
        lines => [
            <<'EOL',

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2013 Franck Latrémolière, Reckon LLP and others.

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
EOL
            $model->{noLinks} ? () : <<EOL,

This workbook is structured as a series of named and numbered tables. There
is a list of tables below, with hyperlinks.  Above each calculation table,
there is a description of the calculations made, and a hyperlinked list of
the tables or parts of tables from which data are used in the calculation.

Hyperlinks point to the first column heading of the relevant table, or to
the first column heading of the relevant part of the table in the case of
references to a particular set of columns within a composite data table.
Scrolling up or down is usually required after clicking a hyperlink in order
to bring the relevant data and/or headings into view.

Some versions of Microsoft Excel can display a "Back" button, which can be
useful when using hyperlinks to navigate around the workbook.
EOL
            <<EOL,

UNLESS STATED OTHERWISE, THIS WORKBOOK IS ONLY A PROTOTYPE FOR TESTING
PURPOSES AND ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
        ]
    );
}

sub configNotes {
    Notes(
        lines => <<'EOL'
Model configuration

This sheet enables some names and labels to be configured.  It does not affect calculations.

The list of tariffs, number of timebands and structure of network levels can only be configured when the model is built.

Voltage and network levels are defined as follows:
* 132kV means voltages of at least 132kV. For the purposes of this workbook, 132kV is not included within EHV.
* EHV, in this workbook, means voltages of at least 22kV and less than 132kV.
* HV means voltages of at least 1kV and less than 22kV.
* LV means voltages below 1kV.
EOL
    );
}

sub inputDataNotes {
    Notes(
        lines => <<'EOL'
Input data

This sheet contains all the input data (except LLFCs which can be entered directly into the Tariff sheet).
EOL
    );
}

sub networkModelNotes {
    Notes(
        lines => <<'EOL'
Network model

This sheet collects data from a network model and calculates aggregated annuitised unit costs from these data.
EOL
    );
}

sub serviceModelNotes {
    Notes(
        lines => <<'EOL'
Service models

This sheet collects and processes data from the service models.
EOL
    );
}

sub lafNotes {
    Notes(
        lines => <<'EOL'
Loss adjustment factors and network use matrices

This sheet calculates matrices of loss adjustment factors and of network use factors.
These matrices map out the extent to which each type of user uses each level of the network, and are used throughout the workbook.
EOL
    );
}

sub loadsNotes {
    my ($model) = @_;
    Notes(
        lines => [
            <<'EOL'
Load characteristics

This sheet compiles information about the assumed characteristics of network users.

A load factor represents the average load of a user or user group, relative to the maximum load level of that user or
user group. Load factors are numbers between 0 and 1.

A coincidence factor represents the expectation value of the load of a user or user group at the time of system maximum load,
relative to the maximum load level of that user or user group.  Coincidence factors are numbers between 0 and 1.

EOL
            , $model->{hasGenerationCapacity}
            ? <<'EOL'
An F factor, for a generator, is the expectation value of output at the time of system maximum load, relative to installed generation capacity. 
F factors are user inputs in respect of generators which are credited on the basis of their installed capacity.

EOL
            : (),
            <<'EOL'
A load coefficient is the expectation value of the load of a user or user group at the time of system maximum load, relative to the average load level of that user or user group.
For demand users, the load coefficient is a demand coefficient and can be calculated as the ratio of the coincidence factor to the load factor.
EOL
        ]
    );
}

sub useNotes {
    Notes(
        lines => <<'EOL'
Network use

This sheet combines the volume forecasts and network use matrices in order to estimate the extent to which the network will be used in the charging year.
EOL
    );
}

sub smlNotes {
    Notes(
        lines => <<'EOL'
Forecast simultaneous maximum load
EOL
    );
}

sub amlNotes {
    Notes(
        lines => <<'EOL'
Forecast aggregate maximum load
EOL
    );
}

sub operatingNotes {
    my ($model) = @_;
    Notes(
        lines => [
            $model->{opAlloc}
            ? (
                'Operating expenditure',
                '',
'This sheet calculates elements of tariff components that recover operating expenditure excluding network rates.'
              )
            : 'Other expenditure'
        ]
    );
}

sub contributionNotes {
    Notes( lines => <<'EOL');
Customer contributions

This sheet calculates factors used to take account of the costs deemed to be covered by connection charges.
EOL
}

sub yardstickNotes {
    Notes(
        lines => <<'EOT'
Yardsticks

This sheet calculates average p/kWh and p/kW/day charges that would apply if no costs were recovered through capacity or fixed charges.
EOT
    );
}

sub multiNotes {
    Notes(
        lines => <<'EOL'
Load characteristics for multiple unit rates
EOL
    );
}

sub standingNotes {
    Notes(
        lines => <<'EOL'
Allocation to standing charges

This sheet reallocates some costs from unit charges to fixed or capacity charges, for demand users only.
EOL
    );
}

sub standingNhhNotes {
    Notes(
        lines => <<'EOL'
Standing charges as fixed charges

This sheet allocates standing charges to fixed charges for non half hourly settled demand users.
EOL
    );
}

sub reactiveNotes {
    my ($model) = @_;
    Notes(
        name  => 'Reactive power unit charges',
        lines => [
            $model->{reactive} && $model->{reactive} =~ /band/i
            ? ( 'The calculations in this sheet are '
                  . 'based on steps 1-6 (Ofgem). '
                  . 'This gives banded reactive power unit charges.' )
            : ()
        ]
    );
}

sub aggregationNotes {
    Notes(
        lines => <<'EOL'
Aggregation

This sheet aggregates elements of tariffs excluding revenue matching and final adjustments and rounding.
EOL
    );
}

sub revenueNotes {
    Notes(
        lines => <<'EOL'
Revenue shortfall or surplus
EOL
    );

=head Development note

Matching starts with a summary of charges which has reactive unit charges.

It does not really need to, but doing it that way enables the table of reactive unit
charges against $allTariffsByEndUsers to look good (rather than look like a silly duplicate
at the bottom of the reactive sheet).

=cut

}

sub scalerNotes {
    Notes(
        lines => <<'EOL'
Revenue matching

This sheet modifies tariffs so that the total expected net revenues matches the target.
EOL
    );
}

sub adderNotes {
    Notes(
        lines => <<'EOL'
Adder
EOL
    );
}

sub roundingNotes {
    Notes( lines => <<'EOL');
Tariff component adjustment and rounding
EOL
}

sub componentNotes {
    Notes(
        name  => 'Tariff components and rules',
        lines => [ split /\n/, <<'EOL']
This sheet is for user information only.  It summarises the rules coded in this model to calculate each components of each tariff, and allows some names and labels to be configured.

The following shorthand is used:
PAYG: higher unit rates applicable to tariffs with no standing charges (as opposed to "Standard").
Yardstick: single unit rates (as opposed to "1", "2" etc.)

EOL
    );
}

1;
