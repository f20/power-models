package SpreadsheetModel::Dataset;

=head Copyright licence and disclaimer

Copyright 2008-2016 Franck Latrémolière, Reckon LLP and others.

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
    return $self->{$wb}{seeOther}->wsUrl($wb) if $self->{$wb}{seeOther};
    my ( $wo, $ro, $co ) = @{ $self->{$wb} }{qw(worksheet row col)};
    my $ce = xl_rowcol_to_cell(
        UNIVERSAL::isa( $self->{location}, 'SpreadsheetModel::CalcBlock' )
        ? ( $ro, ( $co || 1 ) - 1 )
        : ( ( $ro || 1 ) - 1, $co )
    );
    my $wn =
        $wo
      ? $wo->get_name
      : die "No worksheet for $self->{name}"
      . " ($self->{debug} $self->{rows} x $self->{cols})";
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
    my $noPlaceholderData =
         ref $self eq __PACKAGE__
      && !$self->{usePlaceholderData}
      && !( $self->{dataset} && $self->{dataset}{usePlaceholderData} );
    my ( @overrideColumns, @rowKeys );
    if ( my $dataset = $self->dataset( $wb, $ws ) ) {
        my $fc = $self->{colOffset} || 0;
        my $lc = ++$fc + $self->lastCol;
        @overrideColumns = @{$dataset}[ $fc .. $lc ];
        @rowKeys         = map {
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
          && UNIVERSAL::isa( $self->{location}, 'SpreadsheetModel::Columnset' )
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
    my $format = $wb->getFormat( $self->{defaultFormat} || '0.000hard' );
    my $missingFormat =
      $wb->getFormat( $self->{defaultMissingFormat} || 'unused' );
    if ( ref $self->{data}[0] ) {
        my $data = $noPlaceholderData
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
          $noPlaceholderData
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
        return $self->{$wb}{seeOther}->wsWrite( $wb, $ws, $row, $col, $noCopy )
          if $self->{$wb}{seeOther};
        return @{ $self->{$wb} }{qw(worksheet row col)}
          if !$wb->{copy} || $noCopy || $self->{$wb}{worksheet} == $ws;
        return @{ $self->{$wb}{$ws}{$wb} }{qw(worksheet row col)}
          if $self->{$wb}{$ws};
    }

    if ( $self->{location} && UNIVERSAL::can( $self->{location}, 'wsWrite' ) ) {
        $self->{location}->wsWrite( $wb, $ws, $row, $col, $noCopy );
    }
    else {
        my $wsWanted;
        $wsWanted = $wb->{ $self->{location} } if $self->{location};
        $wsWanted = $wb->{dataSheet}
          if !$wsWanted
          && !$self->{ignoreDatasheet}
          && ref $self eq __PACKAGE__;
        return $self->wsWrite( $wb, $wsWanted, undef, undef, 1 )
          if $wsWanted && $wsWanted != $ws;
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

    elsif ( $self->{sourceLines} && @{ $self->{sourceLines} } ) {
        warn
          "$self $self->{name} $self->{debug} is using a deprecated feature; "
          . "sourceLines = @{$self->{sourceLines}}";
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

    my @dataAreHere = ($ws);

    $ws->set_row( $row, 21 );
    $ws->write( $row++, $col, "$self->{name}", $wb->getFormat('caption') );

    my $dataset;
    $dataset = $self->{dataset}{ $self->{number} } if $self->{number};
    undef $dataset unless ref $dataset eq 'ARRAY';

    if ( $self->{lines}
        or !( $wb->{noLinks} && $wb->{noLinks} == 1 )
        and $self->{formulaLines} || $self->{name} && $self->{sourceLines} )
    {
        my $hideFormulas = $wb->{noLinks} && $self->{sourceLines};
        my $textFormat   = $wb->getFormat('text');
        my $linkFormat   = $wb->getFormat('link');
        my $xc           = 0;
        foreach (
            $self->{lines} ? @{ $self->{lines} } : (),
            !( $wb->{noLinks} && $wb->{noLinks} == 1 )
            && $self->{sourceLines} && @{ $self->{sourceLines} }
            ? ( 'Data sources:', @{ $self->{sourceLines} } )
            : (),
            !( $wb->{noLinks} && $wb->{noLinks} == 1 )
            && $self->{formulaLines} ? @{ $self->{formulaLines} }
            : ()
          )
        {
            if ( UNIVERSAL::isa( $_, 'SpreadsheetModel::Object' ) ) {
                my $na = 'x' . ( ++$xc ) . " = $_->{name}";
                if ( my $url = $_->wsUrl($wb) ) {
                    $ws->set_row( $row, undef, undef, 1, 1 )
                      if $hideFormulas;
                    $ws->write_url( $row++, $col, $url, $na, $linkFormat );
                    (
                        $_->{location}
                          && UNIVERSAL::isa( $_->{location},
                            'SpreadsheetModel::Columnset' )
                        ? $_->{location}
                        : $_
                      )->addForwardLink($self)
                      if $wb->{findForwardLinks};
                }
                else {
                    $ws->set_row( $row, undef, undef, 1, 1 )
                      if $hideFormulas;
                    $ws->write_string( $row++, $col, $na, $textFormat );
                }
            }
            elsif (/^(https?|mailto:)/) {
                $ws->set_row( $row, undef, undef, 1, 1 ) if $hideFormulas;
                $ws->write_url( $row++, $col, "$_", "$_", $linkFormat );
            }
            else {
                $ws->set_row( $row, undef, undef, 1, 1 ) if $hideFormulas;
                $ws->write_string( $row++, $col, "$_", $textFormat );
            }
        }
        $ws->set_row( $row, undef, undef, undef, 0, 0, 1 )
          if $hideFormulas;
    }

    ++$row;    # Blank line

    my $lastCol = $self->lastCol;
    my $lastRow = $self->lastRow;

    ++$row
      if $self->{cols}
      || !exists $self->{singleColName}
      || $self->{singleColName};
    ++$col
      if $self->{rows}
      || !exists $self->{singleRowName}
      || $self->{singleRowName};

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
        $ws->write(
            $row - 1, $col,
            _shortNameRow( $self->{name} ),
            $wb->getFormat('thc')
        );
    }
    elsif ( $self->{singleColName} ) {
        $ws->write(
            $row - 1, $col,
            _shortNameRow( $self->{singleColName} ),
            $wb->getFormat('th')
        );
    }

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
        $ws->write(
            $row, $col - 1,
            _shortNameRow( $self->{name} ),
            $wb->getFormat('th')
        );
    }
    elsif ( $self->{singleRowName} ) {
        $ws->write(
            $row, $col - 1,
            _shortNameRow( $self->{singleRowName} ),
            $wb->getFormat('th')
        );
    }

    @dataAreHere[ 1, 2 ] = ( $row, $col );
    @{ $self->{$wb} }{qw(worksheet row col)} = @dataAreHere;

    my $scribbleFormat = $wb->getFormat('scribbles');
    foreach ( 1 .. 1 ) {    # Scribble columns
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

    my $comment = $self->{comment};
    $comment = $self->{lines}
      if !defined $comment
      && $wb->{linesAsComment}
      && $self->{lines};
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
                $value = "=$value"
                  if $value
                  and $value eq '#VALUE!' || $value eq '#N/A'
                  and $wb->formulaHashValues;
                $ws->write( $row + $y, $col + $x, $value, $format );
                if ($comment) {
                    $ws->write_comment(
                        $row + $y,
                        $col + $x,
                        (
                            map { ref $_ eq 'ARRAY' ? join "\n", @$_ : $_; }
                              ref $comment eq 'HASH'
                            ? $comment->{text}
                            : $comment
                        ),
                        x_scale => ref $comment eq 'HASH'
                          && $comment->{x_scale} ? $comment->{x_scale} : 3,
                    );
                    undef $comment;
                }
            }
        }
    }

    $self->dataValidation(
        $wb, $ws, $row, $col,
        $row + $lastRow,
        $col + $lastCol
    ) if $self->{validation};

    if ( $self->{conditionalFormatting} ) {
        foreach (
            ref $self->{conditionalFormatting} eq 'ARRAY'
            ? @{ $self->{conditionalFormatting} }
            : $self->{conditionalFormatting}
          )
        {
            $_->{format} = $wb->getFormat( $_->{format} )
              if $_->{format} && ( ref $_->{format} ) !~ /ormat/;
            eval {
                $ws->conditional_formatting(
                    $row, $col,
                    $row + $lastRow,
                    $col + $lastCol, $_
                );
            };
            if ($@) {
                warn "Omitting conditional formatting: $@";
                return;
            }
        }
    }

    $row += $lastRow;
    $self->requestForwardLinks( $wb, $ws, \$row, $col ) if $wb->{forwardLinks};
    ++$row;
    $dataAreHere[0]{nextFree} = $row unless $dataAreHere[0]{nextFree} > $row;

    if ( $self->{postWriteCalls}{$wb} ) {
        $_->($self) foreach @{ $self->{postWriteCalls}{$wb} };
    }

    @dataAreHere;

}

sub dataValidation {

=item $dataset->dataValidation($wb, $ws, $row, $col, $rowEnd, $colEnd)

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

    $self->{validation}{input_message} ||=
      ref $self->{lines} eq 'ARRAY'
      ? join "\n", @{ $self->{lines} }
      : $self->{lines}
      if $self->{validation}
      && $self->{lines}
      && $wb->{validation}
      && $wb->{validation} =~ /withlinesmsg/i;

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
            $ws->write( $r, $_, '', $wb->getFormat('textnocolour') )
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

            # The duplication of %{ $self->{validation} } is necessary as
            # this function will modify the content of its validation argument.
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
