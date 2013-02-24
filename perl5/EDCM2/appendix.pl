#!/usr/bin/perl

=head Copyright licence and disclaimer

Copyright 2012 Reckon LLP and others.

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
use File::Spec::Functions qw(rel2abs catdir);
use File::Basename 'dirname';
use YAML;
my $perl5dir;

BEGIN {
    $perl5dir =
      dirname( dirname( rel2abs( -l $0 ? ( readlink $0, dirname $0) : $0 ) ) );
}
use lib $perl5dir;

my $workbookModule = 'SpreadsheetModel::Workbook';
my $fileExtension  = '.xls';
require SpreadsheetModel::Workbook;
if ( grep { /xlsx/i } @ARGV ) {
    lib->import( catdir( dirname($perl5dir), 'cpan' ) );
    $workbookModule = 'SpreadsheetModel::WorkbookXLSX';
    $fileExtension .= 'x';
    require SpreadsheetModel::WorkbookXLSX;
}

require Ancillary::DatabaseExport;

my $db = Ancillary::DatabaseExport->new;

$db->summariesByCompany( $workbookModule, $fileExtension,
    appendix => Load <<'EOY' );
---
Scenario 1:
  - "Scenario 1: Excluding all exempt generators, other than those that have already opted in"
  - t4601c6-1
  - t4601c7-1
  - t4601c8-1
  - t4601c9-1
  - t4601c20-1: Total export charge (£/year)
  - t4601c21-1
  - t4601c22-1
  - t4601c23-1
---
Scenario 2:
  - "Scenario 2: Including all generators, including exempt ones"
  - t4601c6-2
  - t4601c7-2
  - t4601c8-2
  - t4601c9-2
  - t4601c20-2: Total export charge (£/year)
  - t4601c21-2
  - t4601c22-2
  - t4601c23-2
---
Scenario 3:
  - "Scenario 3: Excluding all exempt generators, other than those that have already opted in and those who are forecast to receive net credits"
  - t4601c6-3
  - t4601c7-3
  - t4601c8-3
  - t4601c9-3
  - t4601c20-3: Total export charge (£/year)
  - t4601c21-3
  - t4601c22-3
  - t4601c23-3
EOY
