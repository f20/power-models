package Multiyear;

=head Copyright licence and disclaimer

Copyright 2014-2017 Franck Latrémolière, Reckon LLP and others.

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

sub statisticsColumnsets {
    my ( $me, $wbook, $wsheet, $filter ) = @_;
    map {
        my $rows = Labelset( list => $me->{statsRows}[$_] );
        my $statsMaps = $me->{statsMap}[$_];
        $rows->{groups} = 'fake';
        for ( my $r = 0 ; $r < @{ $rows->{list} } ; ++$r ) {
            $rows->{groupid}[$r] = 'fake'
              if grep { $_->[$r] } values %$statsMaps;
        }
        my @columns =
          map {
            if ( my $relevantMap = $statsMaps->{ 0 + $_ } ) {
                SpreadsheetModel::Custom->new(
                    name => $me->modelIdentifier( $_, $wbook, $wsheet ),
                    rows => $rows,
                    custom    => [ map { "=A1$_"; } 0 .. $#$relevantMap ],
                    arguments => {
                        map {
                            my $t;
                            $t = $relevantMap->[$_][0]
                              if $relevantMap->[$_];
                            $t ? ( "A1$_" => $t ) : ();
                        } 0 .. $#$relevantMap
                    },
                    defaultFormat => '0.000copy',
                    rowFormats    => [
                        map {
                            if ( $_ && $_->[0] ) {
                                local $_ = $_->[0]{rowFormats}[ $_->[2] ]
                                  || $_->[0]{defaultFormat};
                                s/(?:soft|hard|con)/copy/ if $_ && !ref $_;
                                $_;
                            }
                            else { 'unavailable'; }
                        } @$relevantMap
                    ],
                    wsPrepare => sub {
                        my ( $self, $wb, $ws, $format, $formula,
                            $pha, $rowh, $colh )
                          = @_;
                        sub {
                            my ( $x, $y ) = @_;
                            my $cellFormat =
                                $self->{rowFormats}[$y]
                              ? $wb->getFormat( $self->{rowFormats}[$y] )
                              : $format;
                            return '', $cellFormat
                              unless $relevantMap->[$y];
                            my ( $table, $offx, $offy ) =
                              @{ $relevantMap->[$y] };
                            my $ph = "A1$y";
                            '', $cellFormat, $formula->[$y],
                              qr/\b$ph\b/ => xl_rowcol_to_cell(
                                $rowh->{$ph} + $offy,
                                $colh->{$ph} + $offx,
                                1, 1,
                              );
                        };
                    },
                );
            }
            else {
                ();
            }
          } @{ $me->{models} };
        $me->{statsColumnsets}[$_] = Columnset(
            name    => $me->{statsSections}[$_],
            columns => \@columns,
        );
      } grep {
        $me->{statsRows}[$_]
          && @{ $me->{statsRows}[$_] }
          && $filter->( $me->{statsSections}[$_] );
      } 0 .. $#{ $me->{statsSections} };
}

sub changeColumnsets {
    my ( $me, $filter ) = @_;
    my %modelMap =
      map { ( 0 + $me->{models}[$_], $_ ) } 0 .. $#{ $me->{models} };
    my @modelNumbers;
    foreach ( 1 .. $#{ $me->{historical} } ) {
        my $old = $me->{historical}[ $_ - 1 ]{dataset}{1000}[2]
          {'Company charging year data version'};
        my $new = $me->{historical}[$_]{dataset}{1000}[2]
          {'Company charging year data version'};
        next unless $old && $new && $old ne $new;
        push @modelNumbers,
          [
            $modelMap{ 0 + $me->{historical}[ $_ - 1 ] },
            $modelMap{ 0 + $me->{historical}[$_] },
          ];
    }
    foreach ( @{ $me->{percentageAssumptionColumns} } ) {
        my $model = $_->{model};
        push @modelNumbers,
          [ $modelMap{ 0 + $model->{sourceModel} }, $modelMap{ 0 + $model }, ];
    }
    map {
        my $cols = $me->{statsColumnsets}[$_]{columns};
        my ( @cola, @colb );
        foreach (@modelNumbers) {
            my ( $before, $after ) = @{$cols}[@$_];
            next unless $before && $after;
            push @cola, Arithmetic(
                name          => $after->{name},
                defaultFormat => '0.000softpm',
                rowFormats    => [
                    map {
                             !defined $before->{rowFormats}[$_]
                          || !defined $after->{rowFormats}[$_] ? undef
                          : $before->{rowFormats}[$_] eq 'unavailable'
                          || $after->{rowFormats}[$_] eq 'unavailable'
                          ? 'unavailable'
                          : eval {
                            local $_ = $after->{rowFormats}[$_];
                            s/copy|soft/softpm/;
                            $_;
                          };
                    } 0 .. $#{ $after->{rows}{list} }
                ],
                arithmetic => '=A1-A2',
                arguments  => { A1 => $after, A2 => $before, },
            );
            push @colb, Arithmetic(
                name          => $after->{name},
                defaultFormat => '%softpm',
                rowFormats    => [
                    map {
                             !defined $before->{rowFormats}[$_]
                          || !defined $after->{rowFormats}[$_] ? undef
                          : $before->{rowFormats}[$_] eq 'unavailable'
                          || $after->{rowFormats}[$_] eq 'unavailable'
                          ? 'unavailable'
                          : undef;
                    } 0 .. $#{ $after->{rows}{list} }
                ],
                arithmetic => '=IF(A2,A1/A3-1,"")',
                arguments  => { A1 => $after, A2 => $before, A3 => $before, },
            );
        }
        (
            @cola ? Columnset(
                name    => "Change: $me->{statsColumnsets}[$_]{name}",
                columns => \@cola,
              )
            : (),
            @colb ? Columnset(
                name    => "Relative change: $me->{statsColumnsets}[$_]{name}",
                columns => \@colb,
              )
            : ()
        );
      } grep {
        $me->{statsColumnsets}[$_] && $filter->( $me->{statsSections}[$_] );
      } 0 .. $#{ $me->{statsSections} };

}

sub addStats {
    my ( $me, $section, $model, @tables ) = @_;
    my ($sectionNumber) = grep { $section eq $me->{statsSections}[$_]; }
      0 .. $#{ $me->{statsSections} };
    unless ( defined $sectionNumber ) {
        push @{ $me->{statsSections} }, $section;
        $sectionNumber = $#{ $me->{statsSections} };
    }
    foreach my $table (@tables) {
        my $lastCol = $table->lastCol;
        if ( my $lastRow = $table->lastRow ) {
            for ( my $row = 0 ; $row <= $lastRow ; ++$row ) {
                my $groupid;
                $groupid = $table->{rows}{groupid}[$row]
                  if $table->{rows}{groupid};
                for ( my $col = 0 ; $col <= $lastCol ; ++$col ) {
                    my $name = "$table->{rows}{list}[$row]";
                    $name .= " — $table->{cols}{list}[$col]"
                      if $lastCol
                      and !$table->{rows}{groups} || defined $groupid;
                    my $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name};
                    unless ( defined $rowNumber ) {
                        if ( defined $groupid ) {
                            my $group = "$table->{rows}{groups}[$groupid]";
                            my $groupRowNumber =
                              $me->{statsRowMap}[$sectionNumber]{$group};
                            if ( defined $groupRowNumber ) {
                                for (
                                    my $i = $groupRowNumber + 1 ;
                                    $i <=
                                    $#{ $me->{statsRows}[$sectionNumber] } ;
                                    ++$i
                                  )
                                {
                                    if ( $me->{statsRows}[$sectionNumber][$i] !~
                                        /^$group \(/ )
                                    {
                                        $rowNumber = $i;
                                        last;
                                    }
                                }
                            }
                            if ( defined $rowNumber ) {
                                splice @{ $me->{statsRows}[$sectionNumber] },
                                  $rowNumber, 0, $name;
                                foreach my $ma (
                                    values %{ $me->{statsMap}[$sectionNumber] }
                                  )
                                {
                                    for ( my $i = @$ma ;
                                        $i > $rowNumber ; --$i )
                                    {
                                        $ma->[$i] = $ma->[ $i - 1 ];
                                    }
                                }
                                map    { ++$_ }
                                  grep { $_ >= $rowNumber }
                                  values %{ $me->{statsRowMap}[$sectionNumber]
                                  };
                                $me->{statsRowMap}[$sectionNumber]{$name} =
                                  $rowNumber;
                            }
                        }
                        unless ( defined $rowNumber ) {
                            push @{ $me->{statsRows}[$sectionNumber] }, $name;
                            $rowNumber =
                              $me->{statsRowMap}[$sectionNumber]{$name} =
                              $#{ $me->{statsRows}[$sectionNumber] };
                        }
                    }
                    $me->{statsMap}[$sectionNumber]{ 0 + $model }[$rowNumber] =
                      [ $table, $col, $row ]
                      unless $table->{rows}{groupid} && !defined $groupid;
                }
            }
        }
        else {
            for ( my $col = 0 ; $col <= $lastCol ; ++$col ) {
                my $name =
                  UNIVERSAL::can( $table->{name}, 'shortName' )
                  ? $table->{name}->shortName
                  : "$table->{name}";
                $name .= " $table->{cols}{list}[$col]" if $lastCol;
                my $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name};
                unless ( defined $rowNumber ) {
                    push @{ $me->{statsRows}[$sectionNumber] }, $name;
                    $rowNumber = $me->{statsRowMap}[$sectionNumber]{$name} =
                      $#{ $me->{statsRows}[$sectionNumber] };
                }
                $me->{statsMap}[$sectionNumber]{ 0 + $model }[$rowNumber] =
                  [ $table, $col, 0 ];
            }
        }
    }
}

1;
