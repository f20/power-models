﻿package SpreadsheetModel::CLI::CommandRunner;

=head Copyright licence and disclaimer

Copyright 2011-2016 Franck Latrémolière and others. All rights reserved.

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

sub useDatabase {

    my $self = shift;

    if ( grep { /extract1076from1001/i } @_ ) {
        require SpreadsheetModel::Data::Database;
        SpreadsheetModel::Data::Database->extract1076from1001;
    }

    if ( grep { /chedam/i } @_ ) {
        require SpreadsheetModel::Data::ChedamMaster;
        SpreadsheetModel::Data::ChedamMaster->runFromDatabase;
    }

    require SpreadsheetModel::Data::Database;
    my $db = SpreadsheetModel::Data::Database->new;

    if ( grep { /\bcsv\b/i } @_ ) {
        require SpreadsheetModel::Data::ExportCsv;
        if ( grep { /^tall$/i } @_ ) {
            $db->dumpTallCsv(@_);
            return;
        }
        $db->dumpEdcmCsv( grep { /^all$/i } @_ );
        return;
    }

    my $workbookModule =
      ( grep { /^-+xls$/i } @_ )
      ? 'SpreadsheetModel::Book::WorkbookXLS'
      : 'SpreadsheetModel::Book::WorkbookXLSX';
    eval "require $workbookModule" or die $@;

    my $options = ( grep { /right/i } @_ ) ? { alignment => 'right' } : {};

    foreach
      my $modelsMatching ( map { /^all$/i ? '.' : m#^/(.+)/$# ? $1 : (); } @_ )
    {
        require SpreadsheetModel::Data::ExportTabs;
        my $tablesMatching = join '|', map { /^([0-9]+)$/ ? "^$1" : (); } @_;
        $tablesMatching ||= '.';
        local $_ = "Compilation $modelsMatching$tablesMatching";
        s/ *[^ a-zA-Z0-9-^]//g;
        $db->tableCompilations( $workbookModule, $options, $_, $modelsMatching,
            $tablesMatching );
        return;
    }

    if ( grep { /\btscs/i } @_ ) {
        require SpreadsheetModel::Data::ExportTscs;
        my @tablesMatching = map { /^([0-9]+)$/ ? "^$1" : (); } @_;
        @tablesMatching = ('.') unless @tablesMatching;
        $options->{tablesMatching} = \@tablesMatching;
        $options->{singleSheet} = grep { /single/i } @_;
        $db->tscsCreateIntermediateTables($options)
          unless grep { /norebuild/i } @_;
        $db->tscsCompilation( $workbookModule, $options, );
        return;
    }

    if (
        my ($change) =
        map { /^(change\S*|dcp\S*)/i ? $1 : /-(?:change|dcp)=(.+)/i ? $1 : (); }
        @_
      )
    {
        my $options = {};
        if ( my (@yml) = grep { /\.ya?ml$/is && -s $_; } @_ ) {
            require YAML;
            @yml = map { YAML::LoadFile($_); } @yml;
            $options = @yml == 1 ? $yml[0] : { map { %$_; } @yml };
        }
        $options->{colour} = 'orange' if grep { /^-*orange$/ } @_;
        my ($name) = map { /-+name=(.*)/i ? $1 : (); } @_;
        my ($base) = map { /-+base=(.+)/i ? $1 : (); } @_;
        if ($base) { $name ||= $change . ' v ' . $base; }
        else {
            $name ||= $change;
            $base = qr/original|clean|after|master|mini|F201|L201|F600|L600/i;
        }
        require SpreadsheetModel::Data::ExportImpact;
        my @arguments = (
            $workbookModule,
            name      => $name,
            basematch => $base,
            newmatch  => "[-+]$change",
            %$options,
        );
        my @outputs =
          map { /^-+(cdcm\S*|edcm\S*|modelm\S*)/ ? $1 : (); } @_;
        @outputs = qw(cdcm) unless @outputs;
        $db->$_(@arguments) foreach map {
            $_ eq 'cdcm'
              ? qw(cdcmTariffImpact cdcmPpuImpact cdcmRevenueMatrixImpact cdcmUserImpact)
              : $_ eq 'edcm' ? qw(edcmTariffImpact edcmRevenueMatrixImpact)
              :                $_;
        } @outputs;
        $db->$_( @arguments, tall => 1 ) foreach map {
                $_ eq 'cdcm' ? qw(cdcmTariffImpact)
              : $_ eq 'edcm' ? qw()
              :                ();
        } @outputs;
        return;
    }

}

1;
