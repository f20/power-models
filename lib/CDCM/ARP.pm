package CDCM::ARP;

=head Copyright licence and disclaimer

Copyright 2014 Franck Latrémolière.

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
use Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Shortcuts ':all';

sub new {
    bless {
        historical => [],
        scenario   => [],
        allmodels  => 0,
      },
      shift;
}

sub worksheetsAndClosuresWithArp {

    my ( $arp, $model, $wbook, @pairs ) = @_;

    push @pairs, 'ARP$' => sub {
        my ($wsheet) = @_;
        push @{ $arp->{finishClosures} }, sub {
            delete $wbook->{logger};
            delete $wbook->{noLinks};
            delete $wbook->{titleAppend};
            $wsheet->set_column( 0, 0,   64 );
            $wsheet->set_column( 1, 255, 12 );
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name  => 'Annual Review Pack',
                lines => [
                         $model->{colour}
                      && $model->{colour} =~ /orange|gold/ ? <<EOL : (),

This document, model or dataset has been prepared by Reckon LLP on the instructions of the DCUSA Panel or
one of its working groups.  Only the DCUSA Panel and its working groups have authority to approve this
material as meeting their requirements.  Reckon LLP makes no representation about the suitability of this
material for the purposes of complying with any licence conditions or furthering any relevant objective.
EOL
                ]
              ),
              $model->licenceNotes;
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name        => 'Historical models',
                sourceLines => [
                    map { [ $_->{nickName}, undef, @{ $_->{sheetLinks} }, ]; }
                      @{ $arp->{historical} }
                ],
              ),
              Notes(
                name        => 'Scenario models',
                sourceLines => [
                    map {
                        [
                            $_->{nickName}, $arp->{assumptions},
                            @{ $_->{sheetLinks} },
                        ];
                    } @{ $arp->{scenario} }
                ],
              );
        };
      }
      unless @{ $arp->{historical} } or @{ $arp->{scenario} };

    push @{ $arp->{models} }, $model;

    push @pairs,
      'Comparisons$' => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 255, 16 );
        push @{ $arp->{finishClosures} }, sub {
            delete $wbook->{logger};
            delete $wbook->{noLinks};
            delete $wbook->{titleAppend};
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name        => 'Model list',
                sourceLines => [
                    map { [ 1 + $_, $arp->{models}[$_]{nickName} ]; }
                      0 .. $#{ $arp->{models} }
                ]
            );
        };
      }
      if @{ $arp->{models} } == 2;

    if ( $model->{dataset}{1000}[2] ) {
        push @{ $arp->{historical} }, $model;
        return @pairs;
    }

    unshift @pairs,
      'Assumptions$' => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 255, 16 );
        push @{ $arp->{finishClosures} }, sub {
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                name        => 'Model list',
                sourceLines => [
                    map { [ 1 + $_, $arp->{models}[$_]{nickName} ]; }
                      0 .. $#{ $arp->{models} }
                ]
            );
        };
        $_->wsWrite( $wbook, $wsheet ) foreach $arp->{assumptions} = Notes(
            name       => '',
            lines      => ['Assumptions'],
            rowFormats => ['caption'],
        );
      }
      unless @{ $arp->{scenario} };

    push @{ $arp->{scenario} }, $model;

    @pairs;

}

sub finish {
    $_->() foreach @{ $_[0]{finishClosures} };
}

1;
