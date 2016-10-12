package SpreadsheetModel::Book::FrontSheet;

=head Copyright licence and disclaimer

Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::FormatLegend;

sub new {
    my ( $class, %properties ) = @_;
    bless \%properties, $class;
}

sub closure {
    my ( $self, $wbook ) = @_;
    sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_print_scale(50);
        $wsheet->set_column( 0, 0,   16 );
        $wsheet->set_column( 1, 1,   112 );
        $wsheet->set_column( 2, 250, 32 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $self->title, $self->extraNotes, $self->dataNotes,
          $self->structureNotes, $self->licenceNotes,
          SpreadsheetModel::FormatLegend->new,
          $wbook->{logger}, $self->technicalNotes;
    };
}

sub technicalNotes {
    my ($self) = @_;
    require POSIX;
    Notes(
        name       => '',
        rowFormats => ['caption'],
        lines      => [
            'Technical model rules and version control',
            $self->{model}{yaml},
            '',
            'Generated on '
              . POSIX::strftime(
                '%a %e %b %Y %H:%M:%S',
                $self->{model}{localTime}
                ? @{ $self->{model}{localTime} }
                : localtime
              )
              . ( $ENV{SERVER_NAME} ? " by $ENV{SERVER_NAME}" : '' ),
        ]
    );
}

sub title {
    Notes( name => $_[0]{name} || 'Index' );
}

sub extraNotes {
    my $notice = $_[0]{model}{extraNotice};
    return unless $notice;
    Notes( name => '', lines => $notice );
}

sub dataNotes {
    Notes(
        name  => '',
        lines => '{unlocked} UNLESS STATED OTHERWISE, THIS WORKBOOK '
          . 'IS ONLY A PROTOTYPE FOR TESTING PURPOSES AND '
          . 'ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.',
    );
}

sub structureNotes {

    return Notes( name => '', lines => <<'EOL')
This workbook is structured as a sequential series of named and numbered tables. There is a list of
tables below, with hyperlinks. Above each calculation table, there is a description of the calculations
and hyperlinks to tables from which data are used. Hyperlinks point to the first relevant table column
heading in the relevant table. Scrolling up or down is usually required after clicking a hyperlink in
order to bring the relevant data and/or headings into view. Some versions of Microsoft Excel can
display a "Back" button, which can be useful when using hyperlinks to navigate around the workbook.
EOL
      if !$_[0]{model}{noLinks} && !$_[0]{model}{tolerateMisordering};

    my @notes;

    push @notes, Notes( name => '', lines => <<'EOL')
This workbook has been laid out as a collection of columns in tables. Columns are sequentially numbered
but are not sequentially laid out in the workbook. There is a list of columns below, with hyperlinks.
Scrolling up or down is usually required after clicking a hyperlink in order to bring the relevant data
and/or headings into view.
EOL
      if $_[0]{model}{tolerateMisordering}
      && $_[0]{model}{layout}
      && $_[0]{model}{layout} =~ /matrix/i;

    push @notes, Notes( name => '', lines => <<'EOL')
Above each calculation table or column, there are some hidden rows containing links to tables or
columns from which data are used. Unfortunately, in most circumstances, revealing the hidden rows will
hide the data, and therefore the usability of this feature is very limited in this workbook.
EOL
      if $_[0]{model}{noLinks} && $_[0]{model}{noLinks} =~ /2/;

    @notes;
}

sub licenceNotes {
    my $notice = $_[0]{copyright}
      || 'Copyright 2008-2016 Franck Latrémolière, Reckon LLP and others.';
    Notes( name => '', lines => [ $notice, <<'EOL'] );
The code used to generate this spreadsheet includes open-source software published at https://github.com/f20/power-models.
Use and distribution of the source code is subject to the conditions stated therein.
Any redistribution of this software must retain the following disclaimer:
THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL AUTHORS OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED
AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
EOL
}

1;
