package SpreadsheetModel::Dataset;

=head Copyright licence and disclaimer

Copyright 2008-2013 Franck Latrémolière, Reckon LLP and others.

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

use SpreadsheetModel::Object ':_util';
our @ISA = qw(SpreadsheetModel::Object);

use SpreadsheetModel::Stack;
use Spreadsheet::WriteExcel::Utility;

use constant {
    OLD_STYLE_SCRIBBLES  => undef,
    NUM_SCRIBBLE_COLUMNS => 1,
    BLANK_LINE           => 1,
};

sub objectType {
    'Input data';
}

sub populateCore {
    my ($self) = @_;
    $self->{core}{$_} = $self->{$_}
      foreach grep { exists $self->{$_}; } qw(arithmetic data);
}

sub wsUrl {
    my ( $self, $wb ) = @_;
    return unless $self->{$wb} && $self->lastRow > -1;
    my ( $wo, $ro, $co ) = @{ $self->{$wb} }{qw(worksheet row col)};
    my $ce = xl_rowcol_to_cell( ( $ro || 1 ) - 1, $co );
    my $wn =
        $wo
      ? $wo->get_name
      : die
"No worksheet for $self->{name} $self->{debug} $self->{rows} x $self->{cols}";
    "internal:'$wn'!$ce";
}

sub htmlDescribe {
    my ( $self, $hb, $hs ) = @_;
    my $formula;
    my $sourceRef = [];
    if ( $self->{arithmetic} && $self->{arguments} ) {
        my @formula = $self->{arithmetic};
        $sourceRef =
          [ _rewriteFormulas( \@formula, [ $self->{arguments} ] ) ];
        ($formula) = @formula;
        $self->{formulaLines} = [ $self->objectType . " $formula" ];
    }
    else {
        $formula = $self->{arithmetic}
          || $self->objectType
          unless ref $self eq __PACKAGE__
          && $hb->{Inputs}
          && $hs == $hb->{Inputs};
    }
    my @arglist;
    my @forlist = $formula ? split( "\n", $formula ) : ();
    @forlist = ( shift @forlist, map { [ br => undef ], $_ } @forlist )
      if @forlist > 1;
    foreach my $i ( 1 .. @$sourceRef ) {
        my $ar = $sourceRef->[ $i - 1 ];
        my $href = join '#', @{ $ar->htmlWrite( $hb, $hs ) };
        push @arglist,
          [ div =>
              [ [ '' => "x$i = " ], [ a => "$ar->{name}", href => $href ] ] ];
    }

    my $hsa = $hb->{Ancillary} || $hs;

    [
        p => [
            [
                  $self->{rows}
                ? $self->{rows}->htmlLink( $hb, $hsa )
                : [ '' => 'Single row' ]
            ],
            [ 'strong' => ' × ' ],
            [
                  $self->{cols}
                ? $self->{cols}->htmlLink( $hb, $hsa )
                : [ '' => 'Single column' ]
            ]
        ]
    ],
      @forlist
      ? [ p => [ strong => [ map { ref $_ ? $_ : [ '' => $_ ] } @forlist ] ] ]
      : (), @arglist,;
}

sub htmlWrite {
    my ( $self, $hb, $hs ) = @_;
    $hs = $hb->{Inputs} if $hb->{Inputs} && ref $self eq __PACKAGE__;
    $self->SUPER::htmlWrite( $hb, $hs );
}

sub check {
    my ($self) = @_;
    $self->{rows} ||= 0;
    $self->{cols} ||= 0;
    $self->{lines} = [ SpreadsheetModel::Object::splitLines( $self->{lines} ) ]
      if $self->{lines};
    if ( $SpreadsheetModel::ShowDimensions && $self->{name} !~ /\]$/ ) {
        $self->{name} = new SpreadsheetModel::Label( $self->{name},
                "$self->{name} ["
              . ( $_[0]{rows} ? @{ $_[0]{rows}{list} } : 1 ) . '×'
              . ( $_[0]{cols} ? @{ $_[0]{cols}{list} } : 1 )
              . ']' );
    }
    if ( ref $self->{data} eq 'CODE' ) {
        my $d = $self->{data};
        my @rows = $self->{rows} ? @{ $self->{rows}{list} } : 0;
        $self->{data} = [
            map {
                my $c = $_;
                [ map { $d->( $_, $c ); } @rows ];
            } $self->{cols} ? @{ $self->{cols}{list} } : 0
        ];
    }
    return;
}

sub lastCol {
    $_[0]{cols} ? $#{ $_[0]{cols}{list} } : 0;
}

sub lastRow {
    $_[0]{rows} ? $#{ $_[0]{rows}{list} } : 0;
}

sub dataset {
    my ( $self, $wb, $ws ) = @_;
    return unless $self->{number} && $self->{dataset};
    my $d = $self->{dataset}{ $self->{number} };
    $d = $self->{dataset}{defaultClosure}->( $self->{number}, $wb, $ws )
      if !$d && $self->{dataset}{defaultClosure};
    ref $d eq 'CODE' ? $d->( $self->{number}, $wb, $ws ) : $d;
}

sub wsPrepare {
    my ( $self, $wb, $ws ) = @_;
    my $noData =
      $wb->{noData} && ref $self eq __PACKAGE__ && !$self->{useIllustrative};
    my ( @overrideColumns, @rowKeys );
    if ( my $dataset = $self->dataset( $wb, $ws ) ) {
        my $fc = $self->{colOffset} || 0;
        ++$fc unless $self->{noRowLabels};
        my $lc = $fc + $self->lastCol;
        @overrideColumns = @{$dataset}[ $fc .. $lc ];
        unless ( ref $dataset->[0] eq 'ARRAY' ) {
            @rowKeys = map {
                local $_ = $_;
                s/.*\n//s;
                s/[^A-Za-z0-9 -]/ /g;
                s/- / /g;
                s/ +/ /g;
                s/^ //;
                s/ $//;
                $_;
              } $self->{rows} ? @{ $self->{rows}{list} }
              : $self->{location}
              && ref $self->{location} eq 'SpreadsheetModel::Columnset'
              ? ( $self->{location}{singleRowName}
                  || _shortNameRow( $self->{location}{name} ) )
              : ( $self->{singleRowName}
                  || _shortNameRow( $self->{name} ) );
            if ( !$self->{rows} && !grep { exists $_->{ $rowKeys[0] } }
                @overrideColumns )
            {
                my @rowKeys2 =
                  grep { !/_column/ } keys %{ $overrideColumns[0] };
                @rowKeys = @rowKeys2 if @rowKeys2;
            }
            elsif ( ref $dataset->[0] eq 'CODE' ) {
                foreach ( grep { !exists $overrideColumns[0]{$_}; } @rowKeys ) {
                    foreach my $trial ( $dataset->[0]->($_) ) {
                        if ( exists $overrideColumns[0]{$trial} ) {
                            $_ = $trial;
                            last;
                        }
                        if ( $trial eq '' ) {
                            foreach my $col (@overrideColumns) {
                                $col->{$_} = '';
                            }
                            last;
                        }
                    }
                }
            }
            $self->{rowKeys} = \@rowKeys;
        }
    }
    my $format = $wb->getFormat( $self->{defaultFormat} || '0.000hard' );
    my $missingFormat =
      $wb->getFormat( $self->{defaultMissingFormat} || 'unused' );
    if ( ref $self->{data}[0] ) {
        my $data = $noData
          ? [
            map {
                [ map { defined $_ ? '#VALUE!' : undef } @$_ ]
            } @{ $self->{data} }
          ]
          : $self->{data};
        $self->{byrow}
          ? sub {
            my ( $x, $y ) = @_;
            my $d;
            if ( defined $data->[$y][$x] ) {
                $d =
                    @rowKeys
                  ? $overrideColumns[$x]{ $rowKeys[$y] }
                  : $overrideColumns[$x][ $y + 1 ]
                  if @overrideColumns;
                $d = $data->[$y][$x] unless defined $d;
            }
            defined $d
              ? (
                $d,
                $self->{rowFormats} && $self->{rowFormats}[$y]
                ? $wb->getFormat( $self->{rowFormats}[$y] )
                : $format
              )
              : ( '', $missingFormat );
          }
          : sub {
            my ( $x, $y ) = @_;
            my $d;
            if ( defined $data->[$x][$y] ) {
                $d =
                    @rowKeys
                  ? $overrideColumns[$x]{ $rowKeys[$y] }
                  : $overrideColumns[$x][ $y + 1 ]
                  if @overrideColumns;
                $d = $data->[$x][$y] unless defined $d;
            }
            defined $d
              ? (
                $d,
                $self->{rowFormats} && $self->{rowFormats}[$y]
                ? $wb->getFormat( $self->{rowFormats}[$y] )
                : $format
              )
              : ( '', $missingFormat );
          };
    }
    else {
        my $data =
          $noData
          ? [ map { defined $_ ? '#VALUE!' : undef } @{ $self->{data} } ]
          : $self->{data};
        $self->lastCol
          ? sub {
            my ( $x, $y ) = @_;
            my $d;
            if ( defined $data->[$x] ) {
                $d =
                    @rowKeys
                  ? $overrideColumns[$x]{ $rowKeys[$y] }
                  : $overrideColumns[$x][ $y + 1 ]
                  if @overrideColumns;
                $d = $data->[$x] unless defined $d;
            }
            defined $d
              ? (
                $d,
                $self->{rowFormats} && $self->{rowFormats}[$y]
                ? $wb->getFormat( $self->{rowFormats}[$y] )
                : $format
              )
              : ( '', $missingFormat );
          }
          : sub {
            my ( $x, $y ) = @_;
            my $d;
            if ( defined $data->[$y] ) {
                $d =
                    @rowKeys
                  ? $overrideColumns[$x]{ $rowKeys[$y] }
                  : $overrideColumns[$x][ $y + 1 ]
                  if @overrideColumns;
                $d = $data->[$y] unless defined $d;
            }
            defined $d
              ? (
                $d,
                $self->{rowFormats} && $self->{rowFormats}[$y]
                ? $wb->getFormat( $self->{rowFormats}[$y] )
                : $format
              )
              : ( '', $missingFormat );
          }
    }
}

sub wsWrite {

    my ( $self, $wb, $ws, $row, $col, $noCopy ) = @_;

    if ( $self->{$wb} ) {

        return @{ $self->{$wb} }{qw(worksheet row col)}
          if !$wb->{copy} || $noCopy || $self->{$wb}{worksheet} == $ws;

        return @{ $self->{$wb}{$ws}{$wb} }{qw(worksheet row col)}
          if $self->{$wb}{$ws};

    }

    if ( $self->{location}
        && ref $self->{location} eq 'SpreadsheetModel::Columnset' )
    {
        $self->{location}->wsWrite( $wb, $ws, $row, $col, $noCopy );
    }
    elsif ($self->{location}
        && $wb->{ $self->{location} } )
    {
        return $self->wsWrite( $wb, $wb->{ $self->{location} }, undef, undef,
            1 )
          if $wb->{ $self->{location} } != $ws;
    }
    elsif (ref $self eq __PACKAGE__
        && $wb->{dataSheet}
        && $wb->{dataSheet} ne $ws )
    {
        return $self->wsWrite( $wb, $wb->{dataSheet}, undef, undef, 1 );
    }

    if ( $self->{$wb} ) {

        return @{ $self->{$wb} }{qw(worksheet row col)}
          if !$wb->{copy} || $noCopy || $self->{$wb}{worksheet} == $ws;
        return @{ $self->{$wb}{$ws}{$wb} }{qw(worksheet row col)}
          if $self->{$wb}{$ws};

        return (
            (
                $self->{$wb}{$ws} =
                  new SpreadsheetModel::Stack( sources => [$self] )
            )->wsWrite( $wb, $ws, $row, $col )
        );

    }

    foreach (qw(rows cols)) {
        $self->{$_}->wsPrepare( $wb, $ws ) if $self->{$_};
    }

    my $cell = $self->wsPrepare( $wb, $ws );

    if ( $wb->{logger} ) {
        $self->addTableNumber( $wb, $ws );
        $wb->{logger}->log($self);
    }

    if ( $self->{arithmetic} && $self->{arguments} ) {
        my @formula = $self->{arithmetic};
        $self->{sourceLines} =
          [ _rewriteFormulas( \@formula, [ $self->{arguments} ] ) ];
        $self->{formulaLines} = [ $self->objectType . " @formula" ];
    }

    elsif ( $self->{sourceLines} ) {
        my %z;
        my @z;
        foreach ( @{ $self->{sourceLines} } ) {
            push @z, $_ unless exists $z{ 0 + $_ };
            undef $z{ 0 + $_ };
        }
        $self->{sourceLines} =
          [ sort { _numsort( $a->{name} ) cmp _numsort( $b->{name} ) } @z ];
    }

    $self->{name} .= " ($self->{debug})"
      if $wb->{debug} && 0 > index $self->{name}, $self->{debug};

    return ( $ws, undef, undef )
      if $self->{rows} && !@{ $self->{rows}{list} }
      || $self->{cols} && !@{ $self->{cols}{list} };

    unless ( defined $row && defined $col ) {
        ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 );
    }

    $ws->set_row( $row, 21 );
    $ws->write( $row++, $col, "$self->{name}", $wb->getFormat('caption') );

    my $dataset;
    $dataset = $self->{dataset}{ $self->{number} } if $self->{number};
    undef $dataset unless ref $dataset eq 'ARRAY';

    if (OLD_STYLE_SCRIBBLES) {
        my $note;
        $note = $dataset->[0]{_note} if $dataset;
        $ws->write_string(
            $row++, $col,
            $note || '',
            $wb->getFormat('scribbles')
        );
    }

    if ( $self->{lines}
        or !$wb->{noLinks}
        and $self->{formulaLines} || $self->{name} && $self->{sourceLines} )
    {
        my $textFormat = $wb->getFormat('text');
        my $linkFormat = $wb->getFormat('link');
        my $xc         = 0;
        foreach (
            $self->{lines} ? @{ $self->{lines} } : (),
            !$wb->{noLinks} && $self->{sourceLines}
            ? ( 'Data sources:', @{ $self->{sourceLines} } )
            : (),
            !$wb->{noLinks}
            && $self->{formulaLines} ? @{ $self->{formulaLines} }
            : ()
          )
        {

            # $ws->set_row( $row, undef, $f );

            if ( ref($_) =~ /^SpreadsheetModel::/ ) {
                my $na = 'x' . ( ++$xc ) . " = $_->{name}";
                if ( my $url = $_->wsUrl($wb) ) {
                    $ws->write_url( $row++, $col, $url, $na, $linkFormat );
                    (
                        $_->{location}
                          && ref $_->{location} eq 'SpreadsheetModel::Columnset'
                        ? $_->{location}
                        : $_
                      )->addForwardLink($self)
                      if $wb->{findForwardLinks};
                }
                else {
                    $ws->write_string( $row++, $col, $na, $textFormat );
                }
            }
            elsif (/^(https?|mailto:)/) {
                $ws->write_url( $row++, $col, "$_", "$_", $linkFormat );
            }
            else {
                $ws->write_string( $row++, $col, "$_", $textFormat );
            }
        }
    }

    ++$row if BLANK_LINE;

    my $lastCol = $self->lastCol;
    my $lastRow = $self->lastRow;

    ++$row
      if $self->{cols}
      || !exists $self->{singleColName}
      || $self->{singleColName};
    ++$col
      if !$self->{noRowLabels}
      and $self->{rows}
      || !exists $self->{singleRowName}
      || $self->{singleRowName};

    my @dataAreHere;
    @dataAreHere = $wb->{ 0 + $self->{rows} }->($lastCol)
      if $self->{rows} && $wb->{ 0 + $self->{rows} };

    if ( $self->{cols} ) {
        $ws->write(
            $row - 1,
            $col + $_,
            _shortNameCol( $self->{cols}{list}[$_], $wb, $ws ),
            $wb->getFormat(
                !$self->{cols}{groups} || defined $self->{cols}{groupid}[$_]
                ? ( $self->{cols}{defaultFormat} || 'thc' )
                : 'thg'
            )
        ) for 0 .. $lastCol;
    }
    elsif ( !exists $self->{singleColName} ) {
        my $srn = _shortNameRow $self->{name};
        $srn =~ s/^[0-9]+[a-z]*\.\s+//i;
        $srn =~ s/\s*\(copy\)$//i;    # hacky - should use Label shortName?
        $ws->write( $row - 1, $col, $srn, $wb->getFormat('thc') );
    }
    elsif ( $self->{singleColName} ) {
        $ws->write(
            $row - 1, $col,
            _shortNameRow( $self->{singleColName} ),
            $wb->getFormat('th')
        );
    }

    unless ( @dataAreHere || $self->{noRowLabels} ) {
        if ( $self->{rows} ) {
            my $thFormat =
              $wb->getFormat( $self->{rows}{defaultFormat} || 'th' );
            my $thgFormat = $wb->getFormat('thg');
            $ws->write(
                $row + $_,
                $col - 1,
                _shortNameRow( $self->{rows}{list}[$_] ),
                !$self->{rows}{groups} || defined $self->{rows}{groupid}[$_]
                ? $thFormat
                : $thgFormat
            ) for 0 .. $lastRow;
        }
        elsif ( !exists $self->{singleRowName} ) {
            my $srn = _shortNameRow $self->{name};
            $srn =~ s/^[0-9]+[a-z]*\.\s+//i;
            $srn =~ s/\s*\(copy\)$//i;    # hacky - should use Label shortName?
            $ws->write( $row, $col - 1, $srn, $wb->getFormat('th') );
        }
        elsif ( $self->{singleRowName} ) {
            $ws->write(
                $row, $col - 1,
                _shortNameRow( $self->{singleRowName} ),
                $wb->getFormat('th')
            );
        }
    }

    if (@dataAreHere) {
        $ws->write_url(
            $row++,
            $col,
            'internal:\''
              . ( $dataAreHere[0] ? $dataAreHere[0]->get_name : 'BROKEN LINK' )
              . '\'!'
              . xl_rowcol_to_cell( @dataAreHere[ 1, 2 ] ),
            'Data',
            $wb->getFormat('link')
        );
        $ws->{nextFree} = $row + 1
          unless $ws->{nextFree} > $row;
        ( $ws, $row, $col ) = @dataAreHere;
    }
    else {
        @dataAreHere = ( $ws, $row, $col );
        my $scribbleFormat = $wb->getFormat('scribbles');
        foreach ( 1 .. NUM_SCRIBBLE_COLUMNS ) {
            my $c2 = $col + $_ + $lastCol;
            my @note;
            if ( $dataset && $self->{rowKeys} ) {
                my $nd = $dataset->[ $lastCol + $_ + 1 ];
                @note = map { $nd->{ $self->{rowKeys}[$_] } } 0 .. $lastRow;
                @note = ( $dataset->[0]{_note} )
                  if !$#note && $dataset->[0] && $dataset->[0]{_note};
            }
            foreach my $y ( 0 .. $lastRow ) {
                $ws->write( $row + $y, $c2, $note[$y], $scribbleFormat );
            }
        }
    }

    foreach my $x ( $self->colIndices ) {
        foreach my $y ( $self->rowIndices ) {
            my ( $value, $format, $formula, @more ) = $cell->( $x, $y );
            if (@more) {

                # 'Not calculated' not working with .xlsx
                $ws->repeat_formula( $row + $y, $col + $x, $formula, $format,
                    @more, 1 ? () : ( result => 'Not calculated' ) );
            }
            elsif ($formula) {

                # 'Not calculated' not working with .xlsx
                $ws->write_formula( $row + $y, $col + $x, $formula, $format,
                    1 ? () : 'Not calculated' );
            }
            else {
                $ws->write( $row + $y, $col + $x, $value, $format );
            }
        }
    }

    $self->dataValidation(
        $wb, $ws, $row, $col,
        $row + $lastRow,
        $col + $lastCol
    ) if $self->{validation};

    @{ $self->{$wb} }{qw(worksheet row col)} = @dataAreHere;

    $row += $lastRow;
    $self->requestForwardLinks( $wb, $ws, \$row, $col ) if $wb->{forwardLinks};
    ++$row;
    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;

    if ( $self->{postWriteCalls}{$wb} ) {
        $_->($self) foreach @{ $self->{postWriteCalls}{$wb} };
    }

    @dataAreHere;

}

sub dataValidation {

=item $dataset->dataValidation($wb, $ws, $row, $col, $rowEnd, $colEnd)

Rules:

* Only to be called if $self->{validation} is true.

* $rowEnd and $colEnd are optional but must be true if supplied.

* Implements $wb->{validation} =~ /lenient/i for blank rows.

=cut

    my ( $self, $wb, $ws, $row, $col, $rowEnd, $colEnd ) = @_;

    return if $wb->{validation} && $wb->{validation} =~ /noval/i;

    if ( $wb->{validation} && $wb->{validation} =~ /nomsg/i ) {
        delete $self->{validation}{$_} foreach qw(input_title input_message);
    }

    $rowEnd ||= $row + $self->lastRow;
    $colEnd ||= $col + $self->lastCol;

    if (   $self->{rows}
        && $self->{rows}{noCollapse}
        && $wb->{validation}
        && $wb->{validation} =~ /lenient/i )
    {
        my @rowArray =
          map  { $row + $_ }
          grep { !defined $self->{rows}{groupid}[$_] }
          0 .. $#{ $self->{rows}{list} };
        push @rowArray, $rowEnd + 1;
        foreach ( 0 .. ( @rowArray - 2 ) ) {
            my $r = $rowArray[$_];
            $ws->write( $r, $_, '', $wb->getFormat('hard') )
              foreach $col .. $colEnd;
            $ws->data_validation(
                $r, $col, $r, $colEnd,
                {
                    validate      => 'length',
                    criteria      => '=',
                    value         => '0',
                    error_message => 'This cell should remain blank.'
                }
            );
            $ws->data_validation( $r + 1, $col, $rowArray[ $_ + 1 ] - 1,
                $colEnd, { %{ $self->{validation} } } );
        }
    }
    else {
        # The duplication of %{ $self->{validation} } is necessary as
        # this function will modify the content of its validation argument.
        $ws->data_validation( $row, $col, $rowEnd, $colEnd,
            { %{ $self->{validation} } } );
    }

}

sub _indices {
    my ($rows) = @_;
       !$rows            ? (0)
      : $rows->{indices} ? @{ $rows->{indices} }
      :                    0 .. $#{ $rows->{list} };
}

sub rowIndices {
    @_ = $_[0]{rows};
    goto \&_indices;
}

sub colIndices {
    @_ = $_[0]{cols};
    goto \&_indices;
}

1;
