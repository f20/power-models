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
our @ISA = qw(DBI::db);

=head

Revision number databases must be created by hand using something like this:

CREATE TABLE revisions (revision integer primary key, yaml text, sha1 text collate binary);
CREATE TABLE scmmapping (revision integer, scmdata text, primary key (revision, scmdata));
CREATE INDEX revsha1 on revisions (sha1);
INSERT INTO revisions (revision) values (5046);

=cut

sub connect {
    my $class = shift;
    bless( DBI->connect(@_), $class );
}

sub revisionText {
    my ( $db, $yaml, $scmData ) = @_;
    my $revision = '';
    require Digest::SHA1;
    my $sha1Machine = new Digest::SHA1;
    my $sha1        = $sha1Machine->add($yaml)->digest;
    foreach ( 1, 0 ) {
        last
          if ($revision) =
          $db->selectrow_array(
            'select revision from revisions where sha1=? and yaml=?',
            undef, $sha1, $yaml );
        $db->do( 'insert into revisions (sha1, yaml) values (?, ?)',
            undef, $sha1, $yaml )
          if $_;
    }
    $db->do(
        'insert or ignore into scmmapping (revision, scmdata) values (?, ?)',
        undef, $revision, $scmData )
      if $revision && $scmData;
    $revision ? "r$revision" : '';
}

sub getRevision {
    my ( $db, $revisionWanted ) = @_;
    $db->selectrow_array( 'select yaml from revisions where revision=?',
        undef, $revisionWanted );
}

sub scmData {
    my ( $db, $revisionWanted ) = @_;
    map { $_->[0]; }
      $db->selectall_arrayref( 'select scmdata from scmmaping where revision=?',
        undef, $revisionWanted );
}

1;
