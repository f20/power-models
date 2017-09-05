package SpreadsheetModel::Rules::RulesTools;

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
use Digest::SHA;

sub combineRulesets {
    my $self = shift;
    my @rulesets;
    foreach (@_) {
        my @a;
        eval { @a = YAML::LoadFile($_); };
        if ($@) {
            warn "$_: $@";
        }
        elsif ( @a == 0 ) {
            warn "$_: no objects found\n";
        }
        else {
            push @rulesets, [ $_, $a[0] ];
        }
    }
    my $kvmap = compileAndIndexRulesets(@rulesets);
    0 and traverse_anonymous($kvmap);
    my @orderedRulesetHtml =
      map {
            qq%<span title="$_">%
          . xmlEscape( $_->[1]{'.'} || $_->[0] )
          . '</span>';
      } @rulesets;
    my ( $fragmentIdsByRuleset, $ruleFragments, $fragmentIdsByMenu ) =
      buildFragmentList( \@orderedRulesetHtml, $kvmap, idIssuerFactory() );
    traverse_named( $ruleFragments, $fragmentIdsByMenu );
    open my $fh, '>', 'index.html' . $$;
    binmode $fh, ':utf8';
    print {$fh}
      toHtml( \@orderedRulesetHtml, $fragmentIdsByRuleset, $ruleFragments );
    close $fh;
    rename 'index.html' . $$, 'index.html';
}

sub compileAndIndexRulesets {

    my (@rulesets) = @_;

    my %kvmap;
    my %keySet;
    my $block = 0;
    my $bit   = 1;
    foreach (@rulesets) {
        my ( $rulesetName, $ruleset ) = @$_;
        unless ( ref $ruleset eq 'HASH' ) {
            die "The ruleset $rulesetName is not a HASH";
        }
        while ( my ( $k, $v ) = each %$ruleset ) {
            next if $k eq '.' || $k eq 'nickName' || $k eq 'template';
            undef $keySet{$k};
            my $fk = "$k?";
            my $fv = (
                 !defined $v ? '¿✗'
                : ref $v     ? '¿✓'
                  . ref($v) . '#'
                  . Digest::SHA::sha1_hex(
                    Encode::encode_utf8( YAML::Dump($v) )
                  )
                : "¿✓$v"
            );
            $kvmap{$fk}{$fv}{hash} ||= { $k => $v };
            $kvmap{$fk}{$fv}{usemap}[$block] |= $bit;
        }
        unless ( $bit <<= 1 ) {
            ++$block;
            $bit = 1;
        }
    }

    foreach my $k ( keys %keySet ) {
        my $fv    = '¿✗';
        my $fk    = "$k?";
        my $block = 0;
        my $bit   = 1;
        foreach (@rulesets) {
            if ( !exists $_->[1]{$k} ) {
                $kvmap{$fk}{$fv}{hash} ||= {};
                $kvmap{$fk}{$fv}{usemap}[$block] |= $bit;
            }
            unless ( $bit <<= 1 ) {
                ++$block;
                $bit = 1;
            }
        }
    }

  STARTAGAIN: {
        my @fKeys = sort keys %kvmap;
        for ( my $i1 = 0 ; $i1 < @fKeys ; ++$i1 ) {
            for ( my $i2 = $i1 + 1 ; $i2 < @fKeys ; ++$i2 ) {
                my $fk1 = $fKeys[$i1];
                my $fk2 = $fKeys[$i2];
                my %possible;
                my @val1 = sort keys %{ $kvmap{$fk1} };
                my @val2 = sort keys %{ $kvmap{$fk2} };
                next if @val1 == 1 && @val2 > 1 || @val2 == 1 && @val1 > 1;
                foreach my $fv1 (@val1) {
                    foreach my $fv2 (@val2) {
                        my @intersection;
                        my $intersectionFlag;
                        for (
                            my $i = 0 ;
                            $i < @{ $kvmap{$fk1}{$fv1}{usemap} } ;
                            ++$i
                          )
                        {
                            $intersectionFlag = 1
                              if $intersection[$i] =
                              ( $kvmap{$fk1}{$fv1}{usemap}[$i] || 0 ) &
                              ( $kvmap{$fk2}{$fv2}{usemap}[$i] || 0 );
                        }
                        if ($intersectionFlag) {
                            $possible{ $fv1 . $fv2 } = {
                                usemap => \@intersection,
                                hash   => {
                                    %{ $kvmap{$fk1}{$fv1}{hash} },
                                    %{ $kvmap{$fk2}{$fv2}{hash} }
                                }
                            };
                        }
                    }
                }
                my @possibleList = keys %possible;
                if ( @possibleList == 1 || @possibleList < @val1 + @val2 - 1 ) {
                    delete $kvmap{$fk1};
                    delete $kvmap{$fk2};
                    $kvmap{ $fk1 . $fk2 } = \%possible;
                    goto STARTAGAIN;
                }
            }
        }
    }

    \%kvmap;

}

sub traverse_anonymous {
    my ($kvmap) = @_;
    my @rules = sort { rand() <=> 0.5; } traverse_anon_helper(
        map {
            [ map { $_->{hash}; } values %$_ ];
        } values %$kvmap
    );
    for ( my $i = 0 ; $i < @rules ; ++$i ) {
        YAML::DumpFile( $$, $rules[$i] );
        rename $$, '%+' . ( 1_000_000 + $i ) . '.yml';
    }
}

sub traverse_anon_helper {
    my $toExpand = shift or return {};
    map {
        my @a = %$_;
        map {
            { @a, %$_ };
        } @$toExpand;
    } traverse_anon_helper(@_);
}

sub traverse_named {
    my ( $ruleFragments, $fragmentIdsByMenu ) = @_;
    my $helper;
    $helper = sub {
        my $toExpand = pop or return [ '%', {} ];
        map {
            my $n = $_->[0];
            my @a = %{ $_->[1] };
            map {
                [
                    join( '+', $n, @$toExpand > 1 ? $_ : () ),
                    { @a, %{ $ruleFragments->{$_} } }
                ];
            } @$toExpand;
        } $helper->(@_);
    };
    foreach ( $helper->(@$fragmentIdsByMenu) ) {
        YAML::DumpFile( $$, $_->[1] );
        rename $$, $_->[0] . '.yml';
    }
}

sub buildFragmentList {

    my ( $orderedRulesets, $kvmap, $idIssuer ) = @_;

    my %ruleFragments;
    my %fragmentIdsByRuleset;
    my %rulesetsByFragmentId;
    my @fragmentIdsByMenu;
    foreach (
        sort { @$a <=> @$b || YAML::Dump($a) cmp YAML::Dump($b); }
        map {
            [
                sort {
                    join( ' ',
                        map { unpack 'b*', pack 'V', $_ || 0; }
                          @{ $b->{usemap} } ) cmp join( ' ',
                        map { unpack 'b*', pack 'V', $_ || 0; }
                          @{ $a->{usemap} } )
                } values %$_
            ];
        } values %$kvmap
      )
    {
        my @options = $idIssuer->($#$_);
        push @fragmentIdsByMenu, \@options;
        for ( my $i = 0 ; $i < @$_ ; ++$i ) {
            $ruleFragments{ $options[$i] } = $_->[$i]{hash};
            {
                my $mask  = 0;
                my $index = -1;
                foreach my $ruleName (@$orderedRulesets) {
                    unless ( $mask <<= 1 ) {
                        ++$index;
                        $mask = 1;
                    }
                    if ( $mask & ( $_->[$i]{usemap}[$index] || 0 ) ) {
                        push @{ $fragmentIdsByRuleset{$ruleName} },
                          $options[$i];
                        push @{ $rulesetsByFragmentId{ $options[$i] } },
                          $ruleName;
                    }
                }
            }
        }
    }

    \%fragmentIdsByRuleset, \%ruleFragments, \@fragmentIdsByMenu,
      \%rulesetsByFragmentId;

}

sub idIssuerFactory {
    my @suits = split //, '♠♡♢♣';    # '♡♢♧♤'; '♠♣♥♦';
    my @values = ( qw(K Q J T), ( reverse 2 .. 9 ), 'A' );
    my @chessB   = split //, '♚♛♜♝♞♟';
    my @chessW   = split //, '♔♕♖♗♘♙';
    my @dieFaces = split //, '⚀⚁⚂⚃⚄⚅';
    my @coloursNotUsed = (
        '#666666', '#1b9e77', '#d95f02', '#7570b3',
        '#e7298a', '#66a61e', '#e6ab02', '#a6761d',
    );
    my $counter = 0;
    sub {
        my ($lastOption) = @_;
        return '★' unless $lastOption;
        return ( shift(@chessW), shift(@chessB) )
          if $lastOption == 1 && @chessB;
        if ( $lastOption < 4 && @values && @suits == 4 ) {
            my $val = pop @values;
            return map { $val . $suits[$_]; } 0 .. $lastOption;
        }
        if ( $lastOption < 6 && @dieFaces == 6 ) {
            push @dieFaces, 'used';
            return map { $dieFaces[$_]; } 0 .. $lastOption;
        }
        if ( $lastOption < @values && @suits ) {
            my $suit = shift @suits;
            return map { $suit . $values[$_]; } 0 .. $lastOption;
        }
        ++$counter;
        return
          map { qq^<span style="font-size:33%">$counter:$_</span>^; }
          0 .. $lastOption;
    };
}

sub toHtml {
    my ( $orderedRulesetHtml, $fragmentIdsByRuleset, $fragmentByFragmentId ) =
      @_;

    my %ids;
    my $lastUsed = 0;
    my $column;
    my $symbolMaker = sub {
        ( local $_ ) = @_;
        my $id = $ids{$_} ||= 'col' . $column . 'frag' . ++$lastUsed . 'z';
qq^<div style="float:left;text-align:center;font-family:Verdana;font-size:115%;margin:-0.2em 0.5em 0.4em 0;width:2em;height:1.5em;border:0.1em solid" onclick="clicked($column, '$id')" class="$id" onmouseover="x=document.getElementById('$id');r=this.getBoundingClientRect();x.style.top=window.scrollY+r.bottom+'px';x.style.left=window.scrollX+r.left+'px';x.style.display='block'" onmouseout="document.getElementById('$id').style.display='none'">$_</div>^;
    };

    join '', <<EOJ
<!DOCTYPE html>
<html><head><meta charset="UTF-8" /><script>// <![CDATA[
function clicked(col, id) {
    var n = document.getElementById('c' + id).childNodes[0];
    document.getElementById('rfrag' + col).value = n ? n.textContent : '';
    var e = document.getElementsByTagName('div'), i;
    for (i in e) {
        if ( ( ' ' + e[i].className ).indexOf( ' col' + col + 'frag' ) > -1 ) {
            if ( ( ' ' + e[i].className ).indexOf(id) > -1 )
            {
                e[i].style.backgroundColor = '#ffeeee';
                e[i].style.borderColor = 'red';
            }
            else {
                e[i].style.backgroundColor = '#ffffff';
                e[i].style.borderColor = '#cccccc';
            }
        }
    }
}
// ]]>
</script></head><body><div>
EOJ
      , (
        map {
            $column = 0;
            (
                '<p style="clear:both">',
                (
                    map { ++$column; $symbolMaker->($_); }
                      @{ $fragmentIdsByRuleset->{$_} }
                ),
                '<span onclick="',
                do {
                    $column = 0;
                    map { ++$column; "clicked($column,'$ids{$_}');"; }
                      @{ $fragmentIdsByRuleset->{$_} };
                },
                '">',
                $_,
                '</span></p>',
            );
        } @$orderedRulesetHtml
      ),
      (
        '<form>',
        '<div style="height:10em">',
        do {
            $column = 0;
            map {
                ++$column;
qq^<textarea style="width:11em;height:6.8em;display:block;float:left" id="rfrag$column" name="rfrag$column"></textarea>^;
            } @{ $fragmentIdsByRuleset->{ $orderedRulesetHtml->[0] } };
        },
        '<div style="clear:both"><input type="submit" /></div>',
        '</div>',
        '</form>',
      ),
      (
        '<script>window.onload=function(){',
        do {
            $column = 0;
            map { ++$column; "clicked($column,'$ids{$_}');"; }
              @{ $fragmentIdsByRuleset->{ $orderedRulesetHtml->[0] } };
        },
        '}</script>',
      ),
      (
        map {
            my $frag = $fragmentByFragmentId->{$_};
            my $text = '';
            $text .= xmlElement(
                div => { style => 'font-size:80%;font-weight:bold' } =>
                  xmlEscape( delete $frag->{$_} ) )
              foreach sort grep { /^\./ } keys %$frag;
            xmlElement(
                div => {
                    id    => $ids{$_},
                    style => 'background:#ccffff;display:none;'
                      . 'position:absolute;margin:0;padding:0.5em',
                },
                $text,
                xmlElement(
                    pre => {
                        id    => "c$ids{$_}",
                        style => 'margin:0;padding:0',
                    },
                    xmlEscape(
                        keys %$frag
                        ? substr( YAML::Dump($frag), 4 )
                        : ''
                    )
                )
            );
          }
          sort keys %$fragmentByFragmentId,

      ),
      '</div></body></html>';

}

sub xmlEscape {
    local @_ = @_ if defined wantarray;
    for (@_) {
        next unless defined $_;
        s/&/&amp;/g;
        s/</&lt;/g;
        s/>/&gt;/g;
        s/"/&quot;/g;
    }
    wantarray ? @_ : $_[0];
}

sub xmlElement {
    my $e = shift;
    my $a;
    $a = shift if ref $_[0] eq 'HASH';
    my $z = "<$e";
    while ( my ( $k, $v ) = each %$a ) {
        $z .= qq% $k="$v"%;
    }
    @_ ? ( join '', $z, '>', @_, "</$e>" ) : "$z />";
}

1;
