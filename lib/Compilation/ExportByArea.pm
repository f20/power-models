package Compilation;

=head Copyright licence and disclaimer

Copyright 2009-2012 Reckon LLP and others.

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

use Encode 'decode_utf8';

sub edcmGenerationOptions {
    require YAML;
    $db->summariesByCompany( $workbookModule,
        appendix => YAML::Load( <<'EOY' ) );
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
}

sub summariesByCompany {
    my ( $self, $workbookModule, $name, @sheets ) = @_;
    my %bidMap;
    foreach (
        @{ $$self->selectall_arrayref('select bid, filename from books') } )
    {
        ( my $bid, local $_ ) = @$_;
        s/\.xlsx?$//is;
        s/-r[0-9]+$//is;
        s/-(?:FCP|LRIC[a-z]*)//is;
        next unless s/-([^-]+)$//s;
        $bidMap{$_}{$1} = $bid;
    }

    my $getTitle =
      $$self->prepare(
        'select v from data where bid=? and tab=? and col=? and row=0');

    my $getData =
      $$self->prepare(
        'select row, v from data where bid=? and tab=? and col=? and row>0');

    while ( my ( $company, $optionhr ) = each %bidMap ) {
        warn "Making $company-$name";
        my $wb =
          $workbookModule->new(
            $company . '-' . $name . $workbookModule->fileExtension );
        $wb->setFormats;
        my $thcFormat   = $wb->getFormat('thc');
        my $titleFormat = $wb->getFormat('notes');
        my $thtarFormat = $wb->getFormat('thtar');

        foreach (@sheets) {
            my ( $sheetName, $columnsar ) = %$_;
            my @columns = @$columnsar or die $sheetName;
            my $ws = $wb->add_worksheet($sheetName);
            $ws->write_string( 0, 0, ( shift @columns ), $titleFormat );
            $ws->set_column( 0, 0, 18 );
            $ws->hide_gridlines(2);
            $ws->freeze_panes( 1, 1 );

            for ( my $c = 0 ; $c < @columns ; ++$c ) {
                local $_ = $columns[$c];
                my $title;
                ( $_, $title ) = %$_ if ref $_;
                my ( $tab, $col, $opt ) = /t([0-9]+)c([0-9]+)-(.*)/;
                next unless $tab;
                unless ($title) {
                    $getTitle->execute( $optionhr->{$opt}, $tab, $col );
                    ($title) = $getTitle->fetchrow_array;
                    eval { $title = decode_utf8 $title; $title =~ s/&amp;/&/g; };
                }
                $title ||= 'No title';
                my $format = $wb->getFormat(
                      $title =~ /kVArh|kWh/ ? '0.000copynz'
                    : $title =~ /\bp\//     ? '0.00copynz'
                    : $title =~ /%/         ? '%softpm'
                    : $title =~ /change/i   ? '0softpm'
                    :                         '0softnz'
                );
                $ws->set_column( 1 + $c, 1 + $c, $title =~ /name/i ? 54 : 18 );
                unless ($c) {
                    $getData->execute( $optionhr->{$opt}, $tab, 0 );
                    while ( my ( $r, $v ) = $getData->fetchrow_array ) {
                        $ws->write( $r + 2, 0, $v, $thtarFormat );
                    }
                }
                $ws->write_string( 2, $c + 1, $title, $thcFormat );
                $getData->execute( $optionhr->{$opt}, $tab, $col );
                while ( my ( $r, $v ) = $getData->fetchrow_array ) {
                    eval { $v = decode_utf8 $v; $v =~ s/&amp;/&/g; };
                    $ws->write( $r + 2, $c + 1, $v, $format );
                }

            }
        }
    }
}

1;
