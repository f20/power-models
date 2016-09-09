package SpreadsheetModel::Data::Database;

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

sub defaultOptions {
    my ( $self, $options ) = @_;
    $options->{name}       ||= 'Comparison';
    $options->{basematch}  ||= sub { $_[0] !~ /change|dcp/i };
    $options->{newmatch}   ||= sub { $_[0] =~ /change|dcp/i };
    $options->{sheetNames} ||= [ split /\n/, <<EOL ];
ENWL
NPG Northeast
NPG Yorkshire
SPEN SPD
SPEN SPM
SSEN SEPD
SSEN SHEPD
UKPN EPN
UKPN LPN
UKPN SPN
WPD EastM
WPD SWales
WPD SWest
WPD WestM
SimplePower
DuckPower
GoosePower
TestPowerF
TestPowerL
EOL
    $options->{sheetTitles} ||= [
        map { "$_: illustrative impact of $options->{name}" } split /\n/,
        <<EOL ];
Electricity North West
Northern Powergrid Northeast
Northern Powergrid Yorkshire
SP Distribution
SP Manweb
Southern Electric Power Distribution
Scottish Hydro Electric Power Distribution
Eastern Power Networks
London Power Networks
South Eastern Power Networks
WPD East Midlands
WPD South Wales
WPD South West
WPD West Midlands
Simple Power Networks
Duck Power Networks
Goose Power Networks
Test Power Networks FCP
Test Power Networks LRIC
EOL
}

sub findLines {

    my ( $self, $options, $linesUnionFlag ) = @_;

    return if $options->{linesAfter} && $options->{linesBefore};

    my @after =
      map { $_->[0] } @{
        $self->selectall_arrayref(
            'select v from data as d1 inner join books using (bid)'
              . ' where tab=? and'
              . ' col=0 and row>0'
              . ' and filename regexp ?'
              . ' and (select 1 from data as d2 where d1.bid=d2.bid'
              . ' and d1.tab=d2.tab and d1.row=d2.row and d2.col>0)'
              . ' group by v order by min(row)',
            undef, $options->{tableNumber}, $options->{newmatch},
        )
      };


    my @before = map { $_->[0] } @{
        $self->selectall_arrayref(
            'select v from data as d1 inner join books using (bid)'
              . ' where tab=? and'
              . ' col=0 and row>0'
              . ' and filename regexp ?'
              . ' and (select 1 from data as d2 where d1.bid=d2.bid'
              . ' and d1.tab=d2.tab and d1.row=d2.row and d2.col>0)'
              . ' group by v order by min(row)',
            undef, $options->{tableNumber}, $options->{basematch},
        )
    };

    if ( join( '|', @after ) eq join( '|', @before ) ) {
        $options->{linesBefore} = $options->{linesAfter} = \@after;
        return;
    }

    my $mapping = $options->{linesMapping} || {};
    my %beforeMap = map { ( $_ => undef ); } @before;
    my ( @newBefore, @newAfter );

    foreach (@after) {
        my $bef = $mapping->{$_};
        $bef = $_ unless defined $bef;
        if ( exists $beforeMap{$bef} ) {
            delete $beforeMap{$bef} if $linesUnionFlag;
            push @newBefore, $bef;
            push @newAfter,  $_;
        }
        elsif ($linesUnionFlag) {
            push @newAfter,  $_;
            push @newBefore, $_;
        }
    }
    if ($linesUnionFlag) {
        foreach ( grep { exists $beforeMap{$_} } @before ) {
            push @newBefore, $_;
            push @newAfter,  $_;
        }
    }
    $options->{linesAfter}  = \@newAfter;
    $options->{linesBefore} = \@newBefore;

}

sub findComponents {
    my ( $self, $options ) = @_;
    $options->{firstColumnBefore} ||= $self->selectall_arrayref(
        'select min(col) from data inner join books using (bid)'
          . ' where tab=? and row=0 and'
          . ' v<>"" and v not like "%LLF%" and v not like "PC%" and'
          . ' v not like "%checksum%"'
          . ' and filename regexp ?',
        undef, $options->{tableNumber}, $options->{basematch},
    )->[0][0];
    $options->{firstColumnAfter} ||= $self->selectall_arrayref(
        'select min(col) from data inner join books using (bid)'
          . ' where tab=? and row=0 and'
          . ' v<>"" and v not like "%LLF%" and v not like "PC%" and'
          . ' v not like "%checksum%"'
          . ' and filename regexp ?',
        undef, $options->{tableNumber}, $options->{newmatch},
    )->[0][0];
    $options->{components} ||= [
        map { $_->[0] } @{
            $self->selectall_arrayref(
                'select v from data inner join books using (bid)'
                  . ' where tab=? and row=0 and col>=? and'
                  . ' v<>"" and v not like "%LLF%" and v not like "PC%" and'
                  . ' v not like "%checksum%"'
                  . ' and filename regexp ?'
                  . ' group by v order by min(col)',
                undef,
                $options->{tableNumber},
                $options->{firstColumnAfter},
                $options->{newmatch},
            )
        }
    ];
    die 'No tariff components found for table '
      . "$options->{tableNumber} in models matching $options->{newmatch}"
      unless @{ $options->{components} };
    $options->{componentsBefore} ||= [
        map { $_->[0] } @{
            $self->selectall_arrayref(
                'select v from data inner join books using (bid)'
                  . ' where tab=? and row=0 and col>=? and'
                  . ' v<>"" and v not like "%LLF%" and v not like "PC%" and'
                  . ' v not like "%checksum%"'
                  . ' and filename regexp ?'
                  . ' group by v order by min(col)',
                undef,
                $options->{tableNumber},
                $options->{firstColumnBefore},
                $options->{basematch},
            )
        }
    ];
    my $componentsString       = join( '|', @{ $options->{components} } );
    my $componentsBeforeString = join( '|', @{ $options->{componentsBefore} } );
    die <<EOE unless $componentsString eq $componentsBeforeString;
Mismatch in tariff components:
$componentsString
$componentsBeforeString
EOE
}

sub genericTariffImpact {

    my ( $self, $wbmodule, %options ) = @_;

    $self->defaultOptions( \%options );
    $self->findLines( \%options );

    my $wb =
      $wbmodule->new( ( $options{tall} ? 'Impact (tall)' : 'Impact' )
        . ' tariffs '
          . $options{name}
          . $wbmodule->fileExtension );
    unless ($wb) {
        warn 'Could not create tariff impact file';
        return;
    }
    $wb->setFormats(
        {
            $options{colour} ? ( colour => $options{colour} ) : (),
            alignment => 1
        }
    );

    my $ncol = @{ $options{components} };
    my @format1 =
      $options{format1}
      ? @{ $options{format1} }
      : map { /k(W|VAr)h/ ? '0.000copy' : '0.00copy'; }
      @{ $options{components} };
    my @format2 =
      $options{format2}
      ? @{ $options{format1} }
      : map { local $_ = $_; s/copy/softpm/; $_; } @format1;
    $_ = $wb->getFormat($_)
      foreach @format1[ 0 .. ( @format1 - 2 ) ],
      @format2[ 0 .. ( @format2 - 2 ) ];
    $_ = $wb->getFormat( $_, 'tlttr' )
      foreach $format1[$#format1], $format2[$#format2];
    my $pcFormat = $wb->getFormat('%softpm');
    my @format3 = map { $pcFormat; } 2 .. $ncol;
    push @format3, $wb->getFormat( '%softpm', 'tlttr' );

    my @books = $self->listModels;
    my $findRow =
      $self->prepare(
        'select row from data where bid=? and tab=? and col=0 and v=?');
    my $q =
      $self->prepare(
        'select v from data where bid=? and tab=? and row=? and col=?');

    my ( $ws, $diff, $perc, $topShift );
    my $hShift     = $options{tall} ? 2 : 0;
    my $vShift     = 0;
    my $sheetSetup = sub {

        my ( $name, $title ) = @_;

        $ws = $wb->add_worksheet($name);
        $ws->hide_gridlines(2);
        if ($title) {
            $topShift = 2;
            $ws->write_string( 0, 0, $title, $wb->getFormat('notes') );
        }
        else { $topShift = 0; }

        if ( $options{tall} ) {
            $ws->set_column( 0, 0,   14 );
            $ws->set_column( 1, 1,   9 );
            $ws->set_column( 2, 2,   48 );
            $ws->set_column( 3, 254, 12 );
            $ws->freeze_panes( 2 + $topShift, 3 );
            if (undef) {
                $ws->write_string( 3, 0, 'DNO area',    $wb->getFormat('thc') );
                $ws->write_string( 3, 1, 'Tariff ID',   $wb->getFormat('thc') );
                $ws->write_string( 3, 2, 'Tariff name', $wb->getFormat('thc') );
            }
        }
        else {
            $ws->set_column( 0, 0,   48 );
            $ws->set_column( 1, 254, 12 );
            $ws->freeze_panes( 2 + $topShift, 1 );
        }

        my $thcaFormat = $wb->getFormat( 'captionca', 'tlttr' );
        $ws->write_string(
            $topShift,
            1 + $hShift,
            'Baseline prices', $thcaFormat
        );
        $ws->write( $topShift, $_ + $hShift, undef, $thcaFormat )
          foreach 2 .. $ncol;

        0 and $thcaFormat = $wb->getFormat( 'captionca', 'tlttr', 'red' );
        $ws->write_string(
            $topShift,
            1 + $ncol + $hShift,
            'Prices on new basis', $thcaFormat
        );
        $ws->write( $topShift, $_ + $ncol + $hShift, undef, $thcaFormat )
          foreach 2 .. $ncol;

        0 and $thcaFormat = $wb->getFormat( 'captionca', 'tlttr', 'blue' );
        $ws->write_string( $topShift, 1 + $ncol * 2 + $hShift,
            'Price change', $thcaFormat );
        $ws->write( $topShift, $_ + $ncol * 2 + $hShift, undef, $thcaFormat )
          foreach 2 .. $ncol;

        0 and $thcaFormat = $wb->getFormat( 'captionca', 'tlttr', 'red' );
        $ws->write_string(
            $topShift,
            1 + $ncol * 3 + $hShift,
            'Percentage change', $thcaFormat
        );
        $ws->write( $topShift, $_ + $ncol * 3 + $hShift, undef, $thcaFormat )
          foreach 2 .. $ncol;

        $diff = $ws->store_formula('=A2-A1');
        $perc = $ws->store_formula('=IF(A1,A3/A2-1,"")');

        my @list = @{ $options{components} };
        for ( my $j = 1 ; $j < 2 + 3 * $ncol ; $j += $ncol ) {
            for ( my $k = 0 ; $k < @list ; ++$k ) {
                $ws->write_string(
                    $topShift + 1,
                    $j + $k + $hShift,
                    $list[$k],
                    $wb->getFormat( 'thc', $k == $#list ? 'tlttr' : () )
                );
            }
        }
    };

    $sheetSetup->( 'All', "Illustrative impact of $options{name}" )
      if $options{tall};

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

        my $linesAfter = $options{linesAfter}
          || [
            map { @$_ } @{
                $self->selectall_arrayref(
                    'select v from data where bid=? and tab=?'
                      . ' and col=0 and row>0',
                    undef, $bida, $options{tableNumber}
                )
            }
          ];

        my $linesBefore = $options{linesBefore} || $linesAfter;

        $sheetSetup->( $options{sheetNames}[$i], $options{sheetTitles}[$i] )
          unless $options{tall};

        my $thFormat = $wb->getFormat('th');
        for ( my $j = 0 ; $j < @$linesAfter ; ++$j ) {
            $findRow->execute( $bidb, $options{tableNumber},
                $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $options{tableNumber},
                $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            if ( $options{tall} ) {
                $ws->write_string(
                    2 + $topShift + $j + $vShift, 0,
                    $options{sheetNames}[$i],     $thFormat
                );
                $ws->write( 2 + $topShift + $j + $vShift,
                    1, 1 + $j, $wb->getFormat('thtar') );
            }
            $ws->write_string(
                2 + $topShift + $j + $vShift,
                $hShift,
                $options{nameExtraColumn}
                ? eval {
                    $q->execute(
                        $bida, $options{tableNumber},
                        $rowa, $options{nameExtraColumn}
                    );
                    my ($x) = $q->fetchrow_array;
                    $x =~ s/\[.*\]//;
                    "Tariff $linesAfter->[$j]: $x";
                }
                : $linesAfter->[$j],
                $thFormat
            );
            for ( my $k = 0 ; $k < $ncol ; ++$k ) {
                $q->execute(
                    $bidb, $options{tableNumber},
                    $rowb, $k + $options{firstColumnBefore}
                );
                my ($vb) = $q->fetchrow_array;
                $q->execute(
                    $bida, $options{tableNumber},
                    $rowa, $k + $options{firstColumnAfter}
                );
                my ($va) = $q->fetchrow_array;

                $ws->write(
                    2 + $topShift + $j + $vShift,
                    $k + 1 + $hShift,
                    $vb, $format1[$k]
                );
                $ws->write(
                    2 + $topShift + $j + $vShift,
                    $k + 1 + $ncol + $hShift,
                    $va, $format1[$k]
                );

                use Spreadsheet::WriteExcel::Utility;
                my $old = xl_rowcol_to_cell( 2 + $topShift + $j + $vShift,
                    $k + 1 + $hShift );
                my $new = xl_rowcol_to_cell( 2 + $topShift + $j + $vShift,
                    $k + 1 + $ncol + $hShift );
                $ws->repeat_formula(
                    2 + $topShift + $j + $vShift,
                    $k + 1 + 2 * $ncol + $hShift,
                    $diff,
                    $format2[$k],
                    A1 => $old,
                    A2 => $new,
                );
                $ws->repeat_formula(
                    2 + $topShift + $j + $vShift,
                    $k + 1 + 3 * $ncol + $hShift,
                    $perc,
                    $format3[$k],
                    A1 => $old,
                    A2 => $old,
                    A3 => $new,
                );
            }
        }
        $vShift += @$linesAfter if $options{tall};
    }

    $ws->autofilter( 1 + $topShift, 0, 1 + $topShift + $vShift, 2 + 4 * $ncol )
      if $options{tall};

}

sub revenueMatrixImpact {

    my ( $self, $wbmodule, %options ) = @_;

    $self->defaultOptions( \%options );
    $self->findLines( \%options, 1 );

    my $wb =
      $wbmodule->new(
        "Impact revenue $options{name}" . $wbmodule->fileExtension );
    unless ($wb) {
        warn 'Could not create revenue impact file';
        return;
    }
    $wb->setFormats(
        { $options{colour} ? ( colour => $options{colour} ) : () } );

    my $thFormat   = $wb->getFormat('th');
    my $thcFormat  = $wb->getFormat('thc');
    my $thcaFormat = $wb->getFormat('caption');
    my $format1    = $wb->getFormat('0copy');
    my $format2    = $wb->getFormat('0softpm');
    my $format3    = $wb->getFormat('%softpm');

    my @books = $self->listModels;
    my $findRow =
      $self->prepare(
        'select row from data where bid=? and tab=? and col=0 and v=?');
    my $q =
      $self->prepare(
        'select v from data where bid=? and tab=? and row=? and col=?');

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

        my $linesAfter = $options{linesAfter}
          || [
            map { @$_ } @{
                $self->selectall_arrayref(
                    'select v from data where bid=? and tab=?'
                      . ' and col=0 and row>0',
                    undef, $bida, $options{tableNumber}
                )
            }
          ];

        my $linesBefore = $options{linesBefore} || $linesAfter;

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

        $ws->write_string( 2, 1, 'Baseline revenue (£/year)',     $thcFormat );
        $ws->write_string( 2, 2, 'Revenue on new basis (£/year)', $thcFormat );
        $ws->write_string( 2, 3, 'Change (£/year)',               $thcFormat );
        $ws->write_string( 2, 4, 'Percentage change',             $thcFormat );

        $ws->write_string( 4 + @$linesAfter,
            0, 'Baseline revenue matrix (£/year)', $thcaFormat );
        $ws->write_string( 7 + 2 * @$linesAfter,
            0, 'Revenue matrix on new basis (£/year)', $thcaFormat );
        $ws->write_string( 10 + 3 * @$linesAfter,
            0, 'Change (£/year)', $thcaFormat );
        $ws->write_string( 13 + 4 * @$linesAfter,
            0, 'Percentage change', $thcaFormat );

        my @list = map { local $_ = $_->[0]; s/[\r\n]+/\n/g; $_; } @{
            $self->selectall_arrayref(
                'select v from data where bid=? and tab='
                  . $options{tableNumber}
                  . ' and row=0 and col>'
                  . $options{col1}
                  . ' and col<'
                  . $options{col2}
                  . ' order by col',
                undef, $bidb
            )
        };

        foreach my $r (
            4 + @$linesAfter,
            7 + 2 * @$linesAfter,
            10 + 3 * @$linesAfter,
            13 + 4 * @$linesAfter
          )
        {
            for ( my $k = 0 ; $k < @list ; ++$k ) {
                $ws->write_string( $r + 1, 1 + $k, $list[$k], $thcFormat );
            }
        }

        use Spreadsheet::WriteExcel::Utility;
        my $diff = $ws->store_formula('=A2-A1');
        my $perc = $ws->store_formula('=IF(A1,A3/A2-1,0)');

        for ( my $j = 0 ; $j < @$linesAfter ; ++$j ) {
            $findRow->execute( $bidb, $options{tableNumber},
                $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $options{tableNumber},
                $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            my $tariffName = $options{nameExtraColumn}
              ? eval {
                $q->execute(
                    $bida, $options{tableNumber},
                    $rowa, $options{nameExtraColumn}
                );
                my ($x) = $q->fetchrow_array;
                $x =~ s/\[.*\]//;
                "Tariff $linesAfter->[$j]: $x";
              }
              : $linesAfter->[$j];
            foreach my $r (
                3,
                6 + @$linesAfter,
                9 + 2 * @$linesAfter,
                12 + 3 * @$linesAfter,
                15 + 4 * @$linesAfter
              )
            {
                $ws->write( $r + $j, 0, $tariffName, $thFormat );
            }
            my $tota = 0;
            my $totb = 0;
            for ( my $k = 1 ; $k < $options{col2} - $options{col1} ; ++$k ) {
                $q->execute(
                    $bidb, $options{tableNumber},
                    $rowb, $k + $options{col1}
                );
                my ($vb) = $q->fetchrow_array;
                $q->execute(
                    $bida, $options{tableNumber},
                    $rowa, $k + $options{col1}
                );
                my ($va) = $q->fetchrow_array;
                $ws->write( 6 + @$linesAfter + $j,     $k, $vb, $format1 );
                $ws->write( 9 + 2 * @$linesAfter + $j, $k, $va, $format1 );
                my $old = xl_rowcol_to_cell( 6 + @$linesAfter + $j,     $k );
                my $new = xl_rowcol_to_cell( 9 + 2 * @$linesAfter + $j, $k );
                $ws->repeat_formula(
                    12 + 3 * @$linesAfter + $j, $k, $diff, $format2,
                    A1 => $old,
                    A2 => $new,
                );
                $ws->repeat_formula(
                    15 + 4 * @$linesAfter + $j, $k, $perc, $format3,
                    A1 => $old,
                    A2 => $old,
                    A3 => $new,
                );

                if ( grep { $k + $options{col1} == $_ } @{ $options{columns} } )
                {
                    $tota += $va if defined $va && $va !~ /^#/;
                    $totb += $vb if defined $vb && $vb !~ /^#/;
                }
            }

            $ws->write( 3 + $j, 1, $totb, $format1 );
            $ws->write( 3 + $j, 2, $tota, $format1 );
            my $old = xl_rowcol_to_cell( 3 + $j, 1 );
            my $new = xl_rowcol_to_cell( 3 + $j, 2 );
            $ws->repeat_formula(
                3 + $j, 3, $diff, $format2,
                A1 => $old,
                A2 => $new,
            );
            $ws->repeat_formula(
                3 + $j, 4, $perc, $format3,
                A1 => $old,
                A2 => $old,
                A3 => $new,
            );
        }

    }

}

1;
