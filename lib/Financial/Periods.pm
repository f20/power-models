package Financial::Periods;

=head Copyright licence and disclaimer

Copyright 2015 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my ( $class, %hash ) = @_;
    my $periods = bless \%hash, $class;
    if ( ref $periods->{periods} ) {
        return $periods->{labelset} =
          Labelset( name => 'Periods', list => $periods->{periods} );
    }
    my ( @firstDay, @lastDay, @name );
    my $startYear  = $periods->{startYear}  || 2014;
    my $startMonth = $periods->{startMonth} || 7;
    my @monthName = qw(Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec);
    if ( my $months = $periods->{numMonths} ) {
        for ( my $i = 0 ; $i < $months ; ++$i ) {
            my $offset = $periods->{reverseTime} ? $months - 1 - $i : $i;
            my $m1     = $startMonth + $offset - 1;
            my $j      = int( $m1 / 12 );
            my $year   = $startYear + $j;
            $m1 -= 12 * $j;
            push @name,     $monthName[$m1] . ' ' . $year;
            push @firstDay, join '', '=DATE(', $year, ',', 1 + $m1, ',1)';
            push @lastDay,  join '', '=DATE(', $year, ',', 2 + $m1, ',1)-1';

            if (   $periods->{insertYears_BadIdea_DoNotUse}
                && $offset % 12 == ( $periods->{reverseTime} ? 0 : 11 ) )
            {
                my $y = $periods->{reverseTime} ? $year : $year - 1;
                push @name, $startMonth == 1 ? $y : join '/', $y, $y + 1;
                push @firstDay, join '', '=DATE(', $y, ',', $startMonth, ',1)';
                push @lastDay, join '', '=DATE(', $y + 1, ',',
                  $startMonth,
                  ',1)-1';
            }
        }
    }
    if ( my $quarters = $periods->{numQuarters} ) {
        for ( my $i = 0 ; $i < $quarters ; ++$i ) {
            my $offset = $periods->{reverseTime} ? $quarters - 1 - $i : $i;
            my $q      = $offset;
            my $j      = int( $offset / 4 );
            my $year   = $startYear + $j;
            $q -= 4 * $j;
            push @name, join ' Q', $year, 1 + $q;
            push @firstDay, join '', '=DATE(', $year, ',',
              $startMonth + 3 * $q, ',1)';
            push @lastDay, join '', '=DATE(', $year, ',',
              $startMonth + 3 * $q + 3, ',1)-1';
        }
    }
    if ( $periods->{numYears} or !@name ) {
        my $years = $periods->{numYears} ||= 7;
        for ( my $i = 0 ; $i < $years ; ++$i ) {
            my $year =
              $startYear + ( $periods->{reverseTime} ? $years - 1 - $i : $i );
            push @name, $startMonth == 1 ? $year : join '/', $year, $year + 1;
            push @firstDay, join '', '=DATE(', $year, ',', $startMonth, ',1)';
            push @lastDay, join '', '=DATE(', $year + 1, ',', $startMonth,
              ',1)-1';
        }
    }
    if ( $periods->{priorPeriod} ) {
        my $month = $startMonth - 1;
        my $year  = $startYear;
        if ( $month < 1 ) { --$year; $month = 12; }
        my $name =
            ( $periods->{priorPeriod} =~ /month/i ? '' : 'End ' )
          . $monthName[ $month - 1 ] . ' '
          . $year;
        my $first = join '', '=DATE(', $year, ',', 1 + $month, ',1)';
        my $last = $first . '-1';
        $first = join '', '=DATE(', $year, ',', $month, ',1)'
          if $periods->{priorPeriod} =~ /month/i;
        if (  !$periods->{reverseTime} & $periods->{priorPeriod} !~ /end/i
            || $periods->{priorPeriod} =~ /start/i )
        {
            unshift @name,     $name;
            unshift @firstDay, $first;
            unshift @lastDay,  $last;
        }
        else {
            push @name,     $name;
            push @firstDay, $first;
            push @lastDay,  $last;
        }
    }
    $periods->{data} = { firstDay => \@firstDay, lastDay => \@lastDay, };
    $periods->{labelset} =
      $periods->{periodsAreFixed}
      ? Labelset(
        name => join( ', ',
            $periods->{numMonths}   ? "$periods->{numMonths} months"     : (),
            $periods->{numQuarters} ? "$periods->{numQuarters} quarters" : (),
            $periods->{numYears}    ? "$periods->{numYears} years"       : (),
        ),
        list => \@name
      )
      : Labelset(
        name     => 'Accounting periods',
        editable => $periods->makeInputDataset(
            1450,
            'puretext',
            name => 'Accounting period labels',
            cols => Labelset(
                list => [ map { "Period $_" } 1 .. @name ]
            ),
            data => \@name,
        ),
      );
    $periods;
}

sub makeInputDataset {
    my ( $periods, $number, $format, @other ) = @_;
    $periods->{model}{periodsAreInputData}
      ? Dataset(
        number        => $number,
        dataset       => $periods->{model}{dataset},
        appendTo      => $periods->{model}{inputTables},
        defaultFormat => $format . 'hard',
        @other,
      )
      : Constant(
        defaultFormat => $format . 'con',
        @other,
      );
}

sub labelset {
    my ($periods) = @_;
    $periods->{labelset};
}

sub firstDay {
    my ($periods) = @_;
    $periods->{inputDataColumns}{firstDay} ||= $periods->makeInputDataset(
        1452, 'date',
        name => $periods->decorate('First day of accounting period'),
        cols => $periods->labelset,
        data => $periods->{data}{firstDay},
    );
}

sub lastDay {
    my ($periods) = @_;
    $periods->{inputDataColumns}{lastDay} ||= $periods->makeInputDataset(
        1451, 'date',
        name => $periods->decorate('Last day of accounting period'),
        cols => $periods->labelset,
        data => $periods->{data}{lastDay},
    );
}

sub indexPrevious {
    my ($periods) = @_;
    $periods->{indexPrevious} ||= Arithmetic(
        name          => $periods->decorate('Index of preceding period'),
        defaultFormat => 'indices',
        arithmetic    => '=MATCH(A1-1,A5_A6,0)',
        arguments     => {
            A1    => $periods->firstDay,
            A5_A6 => $periods->lastDay,
        }
    );
}

sub indexNext {
    my ($periods) = @_;
    $periods->{indexNext} ||= Arithmetic(
        name          => $periods->decorate('Index of following period'),
        defaultFormat => 'indices',
        arithmetic    => '=IF(A2-A3=-1,"Not applicable",MATCH(A1+1,A5_A6,0))',
        arguments     => {
            A1    => $periods->lastDay,
            A2    => $periods->lastDay,
            A3    => $periods->firstDay,
            A5_A6 => $periods->firstDay,
        }
    );
}

sub openingDay {
    my ($periods) = @_;
    $periods->{openingDay} ||= Arithmetic(
        name          => $periods->decorate('Opening day'),
        defaultFormat => 'datesoft',
        arithmetic    => '=IF(A2-A3=-1,"Not applicable",A1)',
        arguments     => {
            A1 => $periods->firstDay,
            A2 => $periods->lastDay,
            A3 => $periods->firstDay,
        }
    );
}

sub decorate {
    ( my $periods, local $_ ) = @_;
    if ( $periods->{suffix} ) {
        $_ .= " ($periods->{suffix})";
        s/\) \(/, /;
    }
    $_;
}

sub subset_temporary_delete_this {
    my ( $periods, $obj ) = @_;
    $obj->{cols} == $periods->labelset
      ? $obj
      : $obj->{ 0 + $periods } ||= Stack(
        name    => $obj->objectShortName . ' (subset)',
        cols    => $periods->labelset,
        sources => [$obj]
      );
}

1;

