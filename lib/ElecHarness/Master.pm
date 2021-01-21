package ElecHarness;

# Copyright 2021 Franck Latrémolière and others.
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

sub sheetPriority {
    ( undef, local $_ ) = @_;
    /Index/ ? 42 : /Waterfall/ ? 41 : 40;
}

sub requiredModulesForRuleset {
    my ( $class, $ruleset ) = @_;
    $ruleset->{waterfalls} ? 'ElecHarness::Waterfalls' : ();
}

sub new {
    my $class = shift;
    my $me = bless { @_, }, $class;
    ${ $me->{sharingObjectRef} } = $me if ref $me->{sharingObjectRef};
    $me;
}

sub finishMultiModelSharing {
    $_->() foreach @{ $_[0]{finishClosures} };
}

sub worksheetsAndClosures {

    my ( $me, $wbook ) = @_;

    push @{ $me->{finishClosures} }, sub {
        delete $wbook->{logger};
        delete $wbook->{titleWriter};
        $wbook->{noLinks} = 1;
    };

    'Steps$' => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 255, 13 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Step rules' ), @{ $me->{stepRules} };
      },
      $me->{waterfalls}
      ? (
        'Calc$' => sub {
            my ($wsheet) = @_;
            $wsheet->set_column( 0, 255, 13 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Waterfalls' );
            push @{ $me->{finishClosures} }, sub {
                $_->wsWrite( $wbook, $wsheet )
                  foreach @{ $me->waterfallTablesAndCharts->[0] };
            };
        },
        'Waterfalls$' => sub {
            my ($wsheet) = @_;
            $wsheet->set_column( 0, 255, 13 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Waterfalls' );
            push @{ $me->{finishClosures} }, sub {
                $_->wsWrite( $wbook, $wsheet )
                  foreach @{ $me->waterfallTablesAndCharts->[1] };
            };
        },
      )
      : (),
      'Index$' => SpreadsheetModel::Book::FrontSheet->new(
        model     => $me,
        copyright => 'Copyright 2021 Franck Latrémolière and others.'
      )->closure($wbook);

}

1;
