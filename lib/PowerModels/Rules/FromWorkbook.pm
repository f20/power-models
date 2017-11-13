package PowerModels::Rules::FromWorkbook;

=head Copyright licence and disclaimer

Copyright 2016-2017 Franck Latrémolière and others. All rights reserved.

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
use YAML;

sub extractYaml {
    my ($workbook) = @_;
    my @yamlBlobs;
    my $current;
    for my $worksheet ( $workbook->worksheets() ) {
        my ( $row_min, $row_max ) = $worksheet->row_range();
        my ( $tableNumber, $evenIfLocked, $columnHeadingsRow, $to1, $to2 );
        for my $row ( $row_min .. $row_max ) {
            my $rowName;
            my $cell = $worksheet->get_cell( $row, 0 );
            my $v;
            $v = $cell->unformatted if $cell;
            next unless defined $v;
            eval { $v = Encode::decode( 'UTF-16BE', $v ); }
              if $v =~ m/\x{0}/;
            if ($current) {
                if ( $v eq '' ) {
                    push @yamlBlobs, $current;
                    undef $current;
                }
                else {
                    $current .= "$v\n";
                }
            }
            elsif ( $v eq '---' ) {
                $current = "$v\n";
            }
        }
        if ($current) {
            push @yamlBlobs, $current;
            undef $current;
        }
    }
    @yamlBlobs;
}

sub jbzWriter {

    my $set;
    $set = sub {
        my ( $scalar, $key, $sha1hex ) = @_;
        if ( $key =~ s#^([^/]*)/## ) {
            $set->( $scalar->{$1} ||= {}, $key, $sha1hex );
        }
        else {
            $scalar->{$key} = $sha1hex;
        }
    };

    sub {
        my ( $book, $workbook ) = @_;
        my %scalars;
        for my $worksheet ( $workbook->worksheets() ) {
            my $scalar = {};
            my ( $row_min, $row_max ) = $worksheet->row_range();
            for my $row ( $row_min .. $row_max ) {
                if ( my $cell = $worksheet->get_cell( $row, 0 ) ) {
                    if ( my $v = $cell->unformatted ) {
                        if ( $v =~ /(\S+): ([0-9a-fA-F]{40})/ ) {
                            $set->(
                                $scalar,
                                $1 eq 'validation' ? 'dataset.yml' : $1, $2
                            );
                        }
                    }
                }
            }
            $scalars{ $worksheet->{Name} } = $scalar if %$scalar;
        }
        return unless %scalars;
        my $jsonModule;
        if    ( eval 'require JSON' )     { $jsonModule = 'JSON'; }
        elsif ( eval 'require JSON::PP' ) { $jsonModule = 'JSON::PP'; }
        else { warn 'No JSON module found'; goto FAIL; }
        $book =~ s/\.xl\S+//i;
        $book .= '.jbz';
        $book =~ s/'/'"'"'/g;
        open my $fh, qq%|bzip2>'$book'% or goto FAIL;
        binmode $fh or goto FAIL;
        print {$fh}
          $jsonModule->new->canonical(1)
          ->utf8->pretty->encode(
            keys %scalars > 1 ? \%scalars : values %scalars )
          or goto FAIL;
        return;
      FAIL: warn $!;
        return;
    };

}

sub rulesWriter {
    sub {
        my ( $book, $workbook ) = @_;
        $book =~ s/\.xl\S+//i;
        $book .= '-rules.yml';
        open my $fh, '>', $book . $$;
        binmode $fh;
        print {$fh} extractYaml($workbook);
        close $fh;
        rename $book . $$, $book;
    };
}

1;
