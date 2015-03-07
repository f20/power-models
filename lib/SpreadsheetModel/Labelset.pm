package SpreadsheetModel::Labelset;

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

use overload
  '""'     => \&toString,
  '0+'     => sub { $_[0] },
  fallback => 1;

require SpreadsheetModel::Object;
our @ISA = qw(SpreadsheetModel::Object);

use Spreadsheet::WriteExcel::Utility;

sub populateCore {
    my ($self) = @_;
    $self->{core}{$_} = [
        map {
            ref $_ eq 'ARRAY'
              ? [ $_->[0]->getCore, $_->[1], $_->[2] ]
              : "$_"
        } @{ $self->{$_} }
      ]
      foreach grep { exists $self->{$_}; } qw(list groups);
    $self->{core}{$_} = [ map { $_->getCore } @{ $self->{$_} } ]
      foreach grep { exists $self->{$_}; } qw(accepts);
}

sub htmlLink {
    my ( $self, $hb, $hs ) = @_;
    my $href = join '#', @{ $self->htmlWrite( $hb, $hs ) };
    [ a => "$self->{name}", href => $href ];
}

sub htmlDescribe {
    my ( $self, $hb, $hs ) = @_;
    $self->{groups}
      ? (
        map {
            [
                fieldset => [
                    [ legend => $_->{name} ],
                    map { [ div => "$_" ] } @{ $_->{list} }
                ]
            ]
        } @{ $self->{groups} }
      )
      : ( map { [ div => "$_" ] } @{ $self->{list} } ),
      $self->{accepts}
      ? [
        fieldset => [
            [ legend => 'Accepts' ],
            map { [ div => $_->htmlLink( $hb, $hs ) ] } @{ $self->{accepts} }
        ]
      ]
      : ();
}

sub wsPrepare {
    my ( $self, $wb, $ws ) = @_;
    return if $wb->{findForwardLinks} ;
    foreach ( grep { ref $_ eq 'ARRAY' } @{ $self->{list} } ) {
        my ( $sh, $ro, $co ) = $_->[0]->wsWrite( $wb, $ws );
        return unless $sh;
        $_ = q%='%
          .$sh->get_name   . q%'!%
          . xl_rowcol_to_cell( $ro + $_->[1], $co + $_->[2] );
    }
}

sub _checkName {
    my ($self) = @_;
    return unless $self->{name} =~ /^Untitled, /;
    $self->{name} =
      @{ $self->{list} } > 4
      ? join '', 'Labels: ', ( map { "$_, " } @{ $self->{list} }[ 0, 1, 2 ] ),
      '…'
      : 'Labels: ' . join ', ', @{ $self->{list} };
    return;
}

sub check {
    my ($self) = @_;
    return $self->_checkName if 'ARRAY' eq ref $self->{list};
    if ( $self->{editable} ) {
        $self->{list} =
          $self->{editable}{cols}
          ? [ map { [ $self->{editable}, 0, $_ ] }
              0 .. $self->{editable}->lastCol ]
          : $self->{editable}{rows} ? [ map { [ $self->{editable}, $_, 0 ] }
              0 .. $self->{editable}->lastRow ]
          : [ $self->{editable} ];
        push @{ $self->{accepts} },
          $self->{editable}{cols} || $self->{editable}{rows};
        return;
    }
    return 'Broken labelset' unless 'ARRAY' eq ref $self->{groups};
    if ( grep { !ref $_; } @{ $self->{groups} } ) {
        $self->{list} = $self->{groups};
        delete $self->{groups};
        return $self->_checkName;
    }
    my @list;
    my @groupid;
    my @indices;
    my $noCollapse = grep { $#{ $_->{list} } } @{ $self->{groups} };
    for my $gid ( 0 .. $#{ $self->{groups} } ) {
        my $g = $self->{groups}[$gid];
        unless ( ref $g eq 'SpreadsheetModel::Labelset' ) {
            push @list,    $g;
            push @indices, $#list;
            push @groupid, $gid;
            next;
        }
        if ( $noCollapse || $#{ $g->{list} } ) {
            push @list, "$g", @{ $g->{list} };
            push @indices, $#list - $#{ $g->{list} } .. $#list;
            push @groupid, undef, map { $gid } @{ $g->{list} };
        }
        else {
            push @list,    @{ $g->{list} };
            push @indices, $#list;
            push @groupid, map { $gid } @{ $g->{list} };
        }
    }
    $self->{noCollapse} = $noCollapse;
    $self->{list}       = \@list;
    $self->{indices}    = \@indices;
    $self->{groupid}    = \@groupid;

    if ( $self->{accepts} ) {
        foreach ( @{ $self->{accepts} } ) {
            warn "$self should probably not be accepting $_"
              if @{ $self->{list} } % @{ $_->{list} };
        }
    }
    return $self->_checkName;
}

sub supersetIndex {
    my ( $self, $superset ) = @_;

    my $key = 'supersetIndex' . ( 0 + $superset );
    return $self->{$key} if exists $self->{$key};

    return $self->{$key} = -@{ $superset->{list} }
      if $self->{accepts} && grep { $_ == $superset } @{ $self->{accepts} };

    my @sind = $superset->indices;
    my @ind  = $self->indices;
    my @id;

    if ( $self->{groups} ) {
        foreach my $gid ( 0 .. $#{ $self->{groups} } ) {
            if ( ref $self->{groups}[$gid] eq 'SpreadsheetModel::Labelset'
                && $self->{groups}[$gid]{list} == $superset->{list} )
            {
                my $ii = 0;
                foreach my $i (@ind) {
                    next unless $self->{groupid}[$i] == $gid;
                    $id[$i] = -1 - $ii++;
                }
            }
        }
    }

    foreach ( grep { !defined $id[$_] } @ind ) {
        my $name = $self->{list}[$_];
        next if ( $id[$_] ) = grep { $name eq $superset->{list}[$_] } @sind;
        if ( $self->{groupid} ) {
            $name = $self->{groups}[ $self->{groupid}[$_] ];
            next if ( $id[$_] ) =
              grep { $name eq $superset->{list}[$_] } @sind;
        }
        if ( $self->{accepts} ) {
            foreach my $a ( @{ $self->{accepts} } ) {
                $name = $a->{list}[$_];
                last
                  if ( $id[$_] ) =
                  grep { $name eq $superset->{list}[$_] } @sind;
            }
        }
        return $self->{$key} = undef unless defined $id[$_];
    }

    # undef return value means that the matching has failed

    $self->{$key} = \@id;
}

sub indices {
    $_[0]{indices} ? @{ $_[0]{indices} } : 0 .. $#{ $_[0]{list} };
}

sub toString {
    $_[0]{name};
}

sub wsWrite {
    warn "A labelset ($_[0]{name} $_[0]{debug})"
      . ' cannot be written to'
      . ' a spreadsheet by itself';
}

1;
