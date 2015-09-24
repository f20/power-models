package Ancillary::CommandRunner;

=head Copyright licence and disclaimer

Copyright 2011-2015 Franck Latrémolière and others. All rights reserved.

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

require Ancillary::CommandRunner::DataTools;
require Ancillary::CommandRunner::MakeModels;
require Ancillary::CommandRunner::ParseSpreadsheet;
require Ancillary::CommandRunner::UseDatabase;
require Ancillary::CommandRunner::YamlTools;

use Encode qw(decode_utf8);
use File::Spec::Functions qw(abs2rel catdir catfile rel2abs);
use File::Basename 'dirname';

use constant {
    C_PERL5DIR => 0,
    C_HOMEDIR  => 1,
    C_FOLDER   => 2,
    C_LOG      => 3,
};

sub factory {
    my ( $class, $perl5dir, $homedir ) = @_;
    bless [ $perl5dir, $homedir ], $class;
}

sub finish {
    my ($self) = @_;
    $self->makeFolder;
}

sub log {
    my ( $self, $verb, @objects ) = @_;
    return if $verb eq 'makeFolder';
    push @{ $self->[C_LOG] },
      join( "\n", $verb, map { "\t$_"; } @objects ) . "\n\n";
}

sub makeFolder {
    my ( $self, $folder ) = @_;
    if ( $self->[C_FOLDER] ) {
        return if $folder && $folder eq $self->[C_FOLDER];
        if ( $self->[C_LOG] ) {
            open my $h, '>', '~$tmptxt' . $$;
            print {$h} @{ $self->[C_LOG] };
            close $h;
            local $_ = "$self->[C_FOLDER].txt";
            s/^_+([^\.])/$1/s;
            rename '~$tmptxt' . $$, $_;
            delete $self->[C_LOG];
        }
        chdir '..';
        my $tmp = '~$tmp-' . $$ . ' ' . $self->[C_FOLDER];
        return if rmdir $tmp;
        rename $self->[C_FOLDER], $tmp . '/~$old-' . $$
          if -e $self->[C_FOLDER];
        rename $tmp, $self->[C_FOLDER];
        delete $self->[C_FOLDER];
    }
    if ($folder) {
        my $tmp = '~$tmp-' . $$ . ' ' . ( $self->[C_FOLDER] = $folder );
        mkdir $tmp;
        chdir $tmp;
    }
}

sub R {
    my ( $self, @commands ) = @_;
    open my $r, '| R --vanilla --slave';
    binmode $r, ':utf8';
    require Compilation::RCode;
    print {$r} Compilation::RCode->rCode(@commands);
}

our $AUTOLOAD;

sub comment { }

sub DESTROY { }

sub AUTOLOAD {
    no strict 'refs';
    warn "$AUTOLOAD not implemented";
    *{$AUTOLOAD} = sub { };
    return;
}

package NOOP_CLASS;
our $AUTOLOAD;

sub AUTOLOAD {
    no strict 'refs';
    *{$AUTOLOAD} = sub { };
    return;
}

1;
