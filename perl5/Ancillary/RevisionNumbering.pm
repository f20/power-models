package Ancillary::RevisionNumbering;

=head Copyright licence and disclaimer

Copyright 2012 Franck Latrémolière and contributors.

Redistribution and use in source and binary forms, with or without modification, are permitted
provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of
conditions and the following disclaimer in the documentation and/or other materials provided
with the distribution.

THIS SOFTWARE IS PROVIDED BY FRANCK LATRÉMOLIÈRE AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
FRANCK LATRÉMOLIÈRE OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF
THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;
use DBI;
use Digest::SHA1;

use constant {
    PATH      => 0,
    HANDLE    => 1,
    STATEMENT => 2,
};

sub connect {
    my ( $class, $path ) = @_;
    $path =~ s/\.sqlite3?$//s;
    $path =~ s/^DBI:SQLite:dbname=//s;
    unless ( -d $path ) {
        unless ( mkdir $path ) {
            warn "mkdir $path: $!";
            return;
        }
    }
    mkdir "$path/scm" unless -e "$path/scm";
    my $dbh = DBI->connect( 'DBI:SQLite:dbname=' . $path . '/~$index.sqlite',
        '', '', { sqlite_unicode => 0, AutoCommit => 1, } );
    $dbh->sqlite_busy_timeout(1_000);
    unless ($dbh) {
        warn "Cannot connect to $path/~\$index.dbh";
        return;
    }
    my $st;
    while ( !( $st = $dbh->prepare('select i from l where h=?') ) ) {
        warn 'Initialising the revisions database';
        $dbh->do(
'create table if not exists l (i integer primary key, h text collate binary)'
        );
        $dbh->do('create index if not exists lh on l (h)');
        my $dh;
        opendir $dh, $path;
        my @importRev = map { /^r([0-9]+)\.yml$/s ? $1 : () } readdir $dh;
        closedir $dh;
        if (@importRev) {
            warn 'Importing data for ' . @importRev . ' revision(s)';
            my $sha1Machine = new Digest::SHA1;
            my $i = $dbh->prepare('insert into l (i, h) values (?, ?)');
            $i->execute( $_,
                $sha1Machine->addfile( new IO::File "< $path/r$_.yml" )
                  ->digest )
              foreach @importRev;
        }
    }
    my $self = bless [ $path, $dbh, $st ], $class;
    eval {
        warn 'Synchronising the revisions database with old monolithic format';
        local undef $/;
        $dbh->do(qq%attach database "$path.sqlite" as old%);
        $dbh->begin_work;
        my $q = $dbh->prepare(
'select i, h from l left join old.revisions on (i=revision and h=sha1) where revision is null'
        );
        $q->execute;
        while ( my ( $i, $h ) = $q->fetchrow_array ) {
            open my $f, '<', "$path/r$i.yml" or warn "Cannot open $path/$i.yml";
            my $yaml = <$f>;
            close $f;
            $dbh->do(
'insert or replace into old.revisions (revision, yaml, sha1) values (?, ?, ?)',
                undef, $i, $yaml, $h
            );
        }
        $q = $dbh->prepare(
'select revision, yaml, sha1 from old.revisions left join l on (i=revision and h=sha1) where sha1 is not null and i is null'
        );
        $q->execute;
        while ( my ( $i, $yaml, $h ) = $q->fetchrow_array ) {
            $dbh->do( 'insert or replace into l values (?, ?)', undef, $i, $h );
            open my $f, '>', "$path/r$i.yml"
              or warn "Cannot open $path/r$i.yml";
            print $f $yaml if $yaml;
            close $f;
            if (
                my $scmData = $dbh->selectrow_array(
                    'select scmdata from old.scmmapping where revision=?',
                    undef, $i
                )
              )
            {
                open $f, '>',
                  "$path/scm/r$i-" . Digest::SHA1::sha1_hex($scmData) . '.yml';
                print $f $scmData;
                close $f;
            }
        }
        $dbh->commit;
        $dbh->do('detach database old');
    } if -e "$path.sqlite";
    $self;
}

sub revisionText {
    my ( $self, $yaml, $scmData ) = @_;
    my $sha1Machine = new Digest::SHA1;
    my $sha1        = $sha1Machine->add($yaml)->digest;
    my $dbh         = $self->[HANDLE];
    my $revision;
    unless ( ($revision) =
        $dbh->selectrow_array( $self->[STATEMENT], undef, $sha1 ) )
    {
        while ( !$dbh->do('begin immediate transaction') ) { }
        $dbh->do( 'insert into l (h) values (?)', undef, $sha1 );
        while ( !$dbh->commit ) { }
        if ( ($revision) =
            $dbh->selectrow_array( $self->[STATEMENT], undef, $sha1 ) )
        {
            open my $f, '>', "$self->[PATH]/r$revision.yml";
            print $f $yaml;
        }
    }
    if ( $revision && $scmData ) {
        my $sha1scm = $sha1Machine->add($scmData)->hexdigest;
        $sha1scm =~ /([0-9a-f]+)/i;
        unless ( -e ( my $fn = "$self->[PATH]/scm/r$revision-$1.yml" ) ) {
            open my $f, '>', $fn;
            print $f $scmData;
        }
    }
    $revision ? "r$revision" : '';
}

1;
