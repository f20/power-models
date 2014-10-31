#!/usr/bin/env perl
use warnings;
use strict;
use utf8;
binmode STDIN,  ':utf8';
binmode STDOUT, ':utf8';
binmode STDERR, ':utf8';
use File::Spec::Functions qw(rel2abs catdir);
use File::Basename 'dirname';
my ( $homedir, $perl5dir );

BEGIN {
    $homedir = dirname( rel2abs( -l $0 ? ( readlink $0, dirname $0) : $0 ) );
    while (1) {
        $perl5dir = catdir( $homedir, 'lib' );
        last if -d catdir( $perl5dir, 'SpreadsheetModel' );
        my $parent = dirname $homedir;
        last if $parent eq $homedir;
        $homedir = $parent;
    }
}
use lib catdir( $homedir, 'cpan' ), $perl5dir;
use Ancillary::CommandParser;
my $parser = Ancillary::CommandParser->factory;
@ARGV ? $parser->dispatch(@ARGV) : $parser->interpret( \*STDIN );
use Ancillary::CommandRunner;
my $runner = Ancillary::CommandRunner->factory( $perl5dir, $homedir );
while ( my ( $verb, @objects ) = $parser->readCommand ) {
    warn "$verb: @objects\n";
    $runner->log( $verb, @objects );
    $runner->$verb(@objects);
}
$runner->finish;
