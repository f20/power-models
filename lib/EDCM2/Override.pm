package EDCM2;

# Copyright 2009-2012 Energy Networks Association Limited and others.
# Copyright 2020 Franck Latrémolière and others.
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

sub tariffOverride {

    my ($model) = @_;

    my @columns = (
        Dataset(
            defaultFormat => '0.000hard',
            name          => 'Import super-red unit rate (p/kWh)',
            rows          => $model->{tariffSet},
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
        ),
        Dataset(
            defaultFormat => '0.00hard',
            name          => 'Import fixed charge (p/day)',
            rows          => $model->{tariffSet},
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
        ),
        Dataset(
            defaultFormat => '0.00hard',
            name          => 'Import capacity rate (p/kVA/day)',
            rows          => $model->{tariffSet},
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
        ),
        Dataset(
            defaultFormat => '0.00hard',
            name          => 'Import exceeded capacity rate (p/kVA/day)',
            rows          => $model->{tariffSet},
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
        ),
        Dataset(
            defaultFormat => '0.000hard',
            name          => 'Export super-red unit rate (p/kWh)',
            rows          => $model->{tariffSet},
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
        ),
        Dataset(
            defaultFormat => '0.00hard',
            name          => 'Export fixed charge p/day',
            rows          => $model->{tariffSet},
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
        ),
        Dataset(
            defaultFormat => '0.00hard',
            name          => 'Export capacity rate (p/kVA/day)',
            rows          => $model->{tariffSet},
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
        ),
        Dataset(
            defaultFormat => '0.00hard',
            name          => 'Export exceeded capacity rate (p/kVA/day)',
            rows          => $model->{tariffSet},
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
        ),
    );

    $model->{table999} = Columnset(
        name     => 'Tariff overrides',
        columns  => \@columns,
        number   => 999,
        location => 999,
        appendTo => $model->{inputTables},
    );

    @columns;

}

1;
