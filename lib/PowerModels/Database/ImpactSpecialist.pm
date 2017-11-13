package PowerModels::Database;

=head Copyright licence and disclaimer

Copyright 2009-2016 Franck Latrémolière, Reckon LLP and others.

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

sub cdcmPpuImpact {

    my ( $self, $wbmodule, %options ) = @_;

    $self->defaultOptions( \%options );
    $options{tableNumber} = 3901;
    $self->findLines( \%options );

    my $wb = $wbmodule->new(
        "Impact pence per unit $options{name}" . $wbmodule->fileExtension );
    unless ($wb) {
        warn 'Could not create ppu impact file';
        return;
    }
    $wb->setFormats(
        {
            $options{colour} ? ( colour => $options{colour} ) : (),
            alignment => 1
        }
    );

    my $linesAfter  = $options{linesAfter};
    my $linesBefore = $options{linesBefore};
    my $thFormat    = $wb->getFormat('th');
    my $thcFormat   = $wb->getFormat('thc');
    my $thcaFormat  = $wb->getFormat('caption');
    my @format1 =
      map { $wb->getFormat($_); } ( map { '0.000copy' } 1 .. 1 );
    my @format2 =
      map { $wb->getFormat($_); } ( map { '0.000softpm' } 1 .. 1 );
    my @format3 =
      map { $wb->getFormat($_); } ( map { '%softpm' } 1 .. 1 );

    my @books   = $self->listModels;
    my $findRow = $self->prepare(
'select row from data where bid=? and tab>3800 and tab<3803 and col=0 and v=?'
    );
    my $q = $self->prepare(
'select v from data where bid=? and tab>3800 and tab<3803 and row=? and col=?'
    );

    foreach my $i ( 0 .. $#{ $options{sheetNames} } ) {
        my $qr = $options{sheetNames}[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/i && $_->[1] =~ /$options{basematch}/i; }
          @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/i && $_->[1] =~ /$options{newmatch}/i; }
          @books;
        next unless $bida;
        $bida = $bida->[0];
        my $ws = $wb->add_worksheet( $options{sheetNames}[$i] );
        $ws->set_column( 0, 0,   48 );
        $ws->set_column( 1, 254, 16 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 1, 1 );
        $ws->write_string(
            0, 0,
            $options{sheetTitles}[$i],
            $wb->getFormat('notes')
        );

        $ws->write_string( 2, 1, 'Baseline average p/kWh',     $thcFormat );
        $ws->write_string( 2, 2, 'Average p/kWh on new basis', $thcFormat );
        $ws->write_string( 2, 3, 'Change (p/kWh)',             $thcFormat );
        $ws->write_string( 2, 4, 'Percentage change',          $thcFormat );

        use Spreadsheet::WriteExcel::Utility;
        my $diff = $ws->store_formula('=A2-A1');
        my $perc = $ws->store_formula('=IF(A1,A3/A2-1,0)');

        for ( my $j = 0 ; $j < @$linesAfter ; ++$j ) {
            $ws->write_string( 3 + $j, 0, $linesAfter->[$j], $thFormat );
            $findRow->execute( $bidb, $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            $q->execute( $bidb, $rowb, 3 );
            my ($rb) = $q->fetchrow_array;
            $q->execute( $bida, $rowa, 3 );
            my ($ra) = $q->fetchrow_array;
            $q->execute( $bidb, $rowb, 1 );
            my ($vb) = $q->fetchrow_array;
            $q->execute( $bida, $rowa, 1 );
            my ($va) = $q->fetchrow_array;
            eval {
                     $va = defined $va
                  && $va !~ /^#/
                  && defined $ra
                  && $ra !~ /^#/ ? 0.1 * $ra / $va : '';
            };
            $va = '' if $@;
            eval {
                     $vb = defined $vb
                  && $vb !~ /^#/
                  && defined $rb
                  && $rb !~ /^#/ ? 0.1 * $rb / $vb : '';
            };
            $vb = '' if $@;
            $ws->write( 3 + $j, 1, $vb, $format1[0] );
            $ws->write( 3 + $j, 2, $va, $format1[0] );
            my $old = xl_rowcol_to_cell( 3 + $j, 1 );
            my $new = xl_rowcol_to_cell( 3 + $j, 2 );
            $ws->repeat_formula(
                3 + $j, 3, $diff, $format2[0],
                A1 => $old,
                A2 => $new,
            );
            $ws->repeat_formula(
                3 + $j, 4, $perc, $format3[0],
                A1 => $old,
                A2 => $old,
                A3 => $new,
            );
        }
    }
}

sub cdcmUserImpact {

    my ( $self, $wbmodule, %options ) = @_;

    $self->defaultOptions( \%options );
    $options{tableNumber} = 4202;
    $self->findLines( \%options );

    my $wb = $wbmodule->new(
        "Impact users $options{name}" . $wbmodule->fileExtension );
    unless ($wb) {
        warn 'Could not create user impact file';
        return;
    }
    $wb->setFormats(
        { $options{colour} ? ( colour => $options{colour} ) : (), } );

    my $linesAfter  = $options{linesAfter};
    my $linesBefore = $options{linesBefore};

    my $scalingFactor = $options{MWh} ? 1      : 0.1;
    my $ppuFormatCore = $options{MWh} ? '0.00' : '0.000';

    my @books   = $self->listModels;
    my $findRow = $self->prepare(
        'select row from data where bid=? and tab=4202 and col=0 and v=?');
    my $q = $self->prepare(
        'select v from data where bid=? and tab=4202 and row=? and col=?');

    foreach my $i ( 0 .. $#{ $options{sheetNames} } ) {
        my $qr = $options{sheetNames}[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/i && $_->[1] =~ /$options{basematch}/i; }
          @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/i && $_->[1] =~ /$options{newmatch}/i; }
          @books;
        next unless $bida;
        $bida = $bida->[0];
        my $ws = $wb->add_worksheet( $options{sheetNames}[$i] );
        $ws->set_column( 0, 0,   48 );
        $ws->set_column( 1, 254, 16 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 1, 1 );
        $ws->write_string(
            0, 0,
            $options{sheetTitles}[$i],
            $wb->getFormat('notes')
        );

        $ws->write_string( 2, 1, 'Baseline £/year',     $wb->getFormat('thc') );
        $ws->write_string( 2, 2, 'Baseline p/kWh',      $wb->getFormat('thc') );
        $ws->write_string( 2, 3, '£/year on new basis', $wb->getFormat('thc') );
        $ws->write_string( 2, 4, 'p/kWh on new basis',  $wb->getFormat('thc') );
        $ws->write_string( 2, 5, 'Change (£/year)',     $wb->getFormat('thc') );
        $ws->write_string( 2, 6, 'Change (p/kWh)',      $wb->getFormat('thc') );
        $ws->write_string( 2, 7, 'Percentage change',   $wb->getFormat('thc') );

        use Spreadsheet::WriteExcel::Utility;
        my $diff = $ws->store_formula('=A2-A1');
        my $perc = $ws->store_formula('=IF(A1,A3/A2-1,0)');

        for ( my $j = 0 ; $j < @$linesAfter ; ++$j ) {
            $ws->write_string( 3 + $j, 0, $linesAfter->[$j],
                $wb->getFormat( $linesAfter->[$j] =~ /\(/ ? 'th' : 'thg' ) );
            $findRow->execute( $bidb, $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            {
                $q->execute( $bidb, $rowb, 1 );
                my ($vb) = $q->fetchrow_array;
                $q->execute( $bida, $rowa, 1 );
                my ($va) = $q->fetchrow_array;
                next unless defined $va && defined $vb;
                $ws->write( 3 + $j, 1, $vb, $wb->getFormat('0copy') );
                $ws->write( 3 + $j, 3, $va, $wb->getFormat('0copy') );
                my $old = xl_rowcol_to_cell( 3 + $j, 1 );
                my $new = xl_rowcol_to_cell( 3 + $j, 3 );
                $ws->repeat_formula(
                    3 + $j, 5, $diff, $wb->getFormat('0softpm'),
                    A1 => $old,
                    A2 => $new,
                );
                $ws->repeat_formula(
                    3 + $j, 7, $perc, $wb->getFormat('%softpm'),
                    A1 => $old,
                    A2 => $old,
                    A3 => $new,
                );
            }
            {
                $q->execute( $bidb, $rowb, 2 );
                my ($vb) = $q->fetchrow_array;
                $q->execute( $bida, $rowa, 2 );
                my ($va) = $q->fetchrow_array;
                next
                  unless defined $va
                  && defined $vb
                  && $va !~ /^#/
                  && $vb !~ /^#/;
                $ws->write(
                    3 + $j, 2,
                    $vb * $scalingFactor,
                    $wb->getFormat( $ppuFormatCore . 'copy' )
                );
                $ws->write(
                    3 + $j, 4,
                    $va * $scalingFactor,
                    $wb->getFormat( $ppuFormatCore . 'copy' )
                );
                my $old = xl_rowcol_to_cell( 3 + $j, 2 );
                my $new = xl_rowcol_to_cell( 3 + $j, 4 );
                $ws->repeat_formula(
                    3 + $j, 6, $diff,
                    $wb->getFormat( $ppuFormatCore . 'softpm' ),
                    A1 => $old,
                    A2 => $new,
                );
            }
        }
    }
}

sub modelmEdcmImpact {

    my ( $self, $wbmodule, %options ) = @_;

    $self->defaultOptions( \%options );

    my $wb = $wbmodule->new(
        "Impact EDCM discounts $options{name}" . $wbmodule->fileExtension );
    unless ($wb) {
        warn 'Could not create EDCM discount impact file';
        return;
    }
    $wb->setFormats(
        { $options{colour} ? ( colour => $options{colour} ) : (), } );

    my @books = $self->listModels;
    my $q     = $self->prepare(
        'select v from data where bid=? and tab=1504 and row=? and col=1');

    foreach my $i ( 0 .. $#{ $options{sheetNames} } ) {
        my $qr = $options{sheetNames}[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/i && $_->[1] =~ /$options{basematch}/i; }
          @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/i && $_->[1] =~ /$options{newmatch}/i; }
          @books;
        next unless $bida;
        $bida = $bida->[0];
        my $ws = $wb->add_worksheet( $options{sheetNames}[$i] );
        $ws->set_column( 0, 254, 16 );
        $ws->hide_gridlines(2);
        0 and $ws->freeze_panes( 1, 1 );
        $ws->write_string(
            0, 0,
            $options{sheetTitles}[$i],
            $wb->getFormat('notes')
        );

        my $ncol = 4;
        my $thcaFormat = $wb->getFormat( 'captionca', 'tlttr' );
        $ws->write_string( 2, 1, 'Baseline discounts', $thcaFormat );
        $ws->write( 2, $_, undef, $thcaFormat ) foreach 2 .. $ncol;
        $ws->write_string( 2, 1 + $ncol, 'Discounts on new basis',
            $thcaFormat );
        $ws->write( 2, $_ + $ncol, undef, $thcaFormat ) foreach 2 .. $ncol;
        $ws->write_string( 2, 1 + $ncol * 2,
            'Change in discount', $thcaFormat );
        $ws->write( 2, $_ + $ncol * 2, undef, $thcaFormat ) foreach 2 .. $ncol;

        my $thcFormat = $wb->getFormat('thc');
        $ws->write_string( 3, $_, 'LV demand', $thcFormat ) foreach 1, 5, 9;
        $ws->write_string( 3, $_, 'LV Sub demand or LV generation', $thcFormat )
          foreach 2, 6, 10;
        $ws->write_string( 3, $_, 'HV demand or LV Sub generation', $thcFormat )
          foreach 3, 7, 11;
        $thcFormat = $wb->getFormat( 'thc', 'tlttr' );
        $ws->write_string( 3, $_, 'HV generation', $thcFormat )
          foreach 4, 8, 12;

        use Spreadsheet::WriteExcel::Utility;
        my $diff = $ws->store_formula('=A2-A1');

        my @rows = (
            'Boundary 0000',
            'Boundary 132kV',
            'Boundary 132kV/EHV',
            'Boundary EHV',
            'Boundary HVplus'
        );
        for ( my $j = 0 ; $j < 5 ; ++$j ) {
            $ws->write_string( 4 + $j, 0, $rows[$j], $wb->getFormat('th') );
            for ( my $k = 1 ; $k < 5 ; ++$k ) {
                my $row = $k + ( 5 - $j ) * 4;
                my @deco;
                push @deco, 'tlttr' if $k == 4;
                {
                    $q->execute( $bidb, $row );
                    my ($vb) = $q->fetchrow_array;
                    $q->execute( $bida, $row );
                    my ($va) = $q->fetchrow_array;
                    next
                      unless defined $va
                      && defined $vb
                      && $va !~ /^#/
                      && $vb !~ /^#/;
                    $ws->write( 4 + $j, $k, $vb,
                        $wb->getFormat( '%copy', @deco ) );
                    $ws->write( 4 + $j, 4 + $k, $va,
                        $wb->getFormat( '%copy', @deco ) );
                    my $old = xl_rowcol_to_cell( 4 + $j, $k );
                    my $new = xl_rowcol_to_cell( 4 + $j, 4 + $k );
                    $ws->repeat_formula(
                        4 + $j, 8 + $k, $diff,
                        $wb->getFormat( '%softpm', @deco ),
                        A1 => $old,
                        A2 => $new,
                    );
                }
            }
        }
    }
}

1;
