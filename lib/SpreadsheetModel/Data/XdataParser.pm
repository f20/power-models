package SpreadsheetModel::Data::XdataParser;

=head Copyright licence and disclaimer

Copyright 2011-2017 Franck Latrémolière, Reckon LLP and others.

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

use constant {
    XDP_dataOverrides      => 0,
    XDP_setRule            => 1,
    XDP_applyDataOverrides => 2,
    XDP_jsonMachineMaker   => 3,
};

sub new {
    my ( $class, $dataOverrides, $setRule, $jsonMachineMaker ) = @_;
    my $applyDataOverrides = sub {
        foreach ( grep { ref $_ eq 'HASH' } @_ ) {
            while ( my ( $tab, $dat ) = each %$_ ) {
                if ( ref $dat eq 'HASH' ) {
                    while ( my ( $row, $rd ) = each %$dat ) {
                        next unless ref $rd eq 'ARRAY';
                        for ( my $col = 0 ; $col < @$rd ; ++$col ) {
                            $dataOverrides->{$tab}[ $col + 1 ]{$row} =
                              $rd->[$col];
                        }
                    }
                }
                elsif ( ref $dat eq 'ARRAY' ) {
                    for ( my $col = 0 ; $col < @$dat ; ++$col ) {
                        my $cd = $dat->[$col];
                        next unless ref $cd eq 'HASH';
                        while ( my ( $row, $v ) = each %$cd ) {
                            $dataOverrides->{$tab}[$col]{$row} = $v;
                        }
                    }
                }
            }
        }
    };
    my $self =
      bless [ $dataOverrides, $setRule, $applyDataOverrides,
        $jsonMachineMaker, ],
      $class;
    $self;
}

sub doParseXdata {

    my $self = shift;

    foreach (@_) {
        if (/^---\r?\n/s) {
            my @y = eval { Load $_; };
            warn $@ if $@;
            $self->[XDP_applyDataOverrides]->(@y);
            next;
        }
        local $_ = $_;
        if (s/\{(.*)\}//s) {
            foreach ( grep { $_ } split /\}\s*\{/s, $1 ) {
                my $d =
                  $self->[XDP_jsonMachineMaker]->()->decode( '{' . $_ . '}' );
                next unless ref $d eq 'HASH';
                $self->[XDP_setRule]->(
                    map { %$_; } grep { $_; }
                      map  { delete $d->{$_}; }
                      grep { /^rules?$/is; }
                      keys %$d
                );
                $self->[XDP_applyDataOverrides]->($d);
            }
            next;
        }
        while (s/(\S.*\|.*\S)//m) {
            my ( $tab, $col, @more ) = split /\|/, $1, -1;
            next unless $tab;
            if ( $tab =~ /^rules$/is ) {
                $self->[XDP_setRule]->( $col, @more );
            }
            elsif ( @more == 1 ) {
                $self->[XDP_dataOverrides]{$tab}{$col} = $more[0];
            }
            elsif (@more == 2
                && $tab =~ /^[0-9]+$/s
                && $col
                && $col =~ /^[0-9]+$/s )
            {
                $self->[XDP_dataOverrides]{$tab}[$col]{ $more[0] } = $more[1];
            }
        }
        s/^\s*\n//gm;
        $self->[XDP_setRule]->( extraNotice => $_ ) if /\S/;
    }

}

1;
