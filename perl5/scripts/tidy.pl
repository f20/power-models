#!/usr/bin/env perl

# Copyright 2012 Franck Latrémolière, Reckon LLP.

use warnings;
use strict;

foreach (@ARGV) {
    next unless -f $_;
    open my $f, '<:utf8', $_ or next;
    local undef $/;
    my $source = <$f>;
    close $f;
    my $destination;
    my $messages;
    my $error;

    if (/\.p[lm]$/) {
        $source =~ s/^\N{U+FEFF}//s;
        require Perl::Tidy;
        local $/ = "\n";    # for perltidy
        $error = Perl::Tidy::perltidy(
            source      => \$source,
            errorfile   => \$messages,
            destination => \$destination,
            argv        => [],
        );
    }
    elsif (/\.ya?ml/) {
        require YAML;
        eval { $destination = YAML::Dump( YAML::Load($source) ); };
        $messages = $@;
    }
    else {
        warn "Ignored: $_";
        next;
    }
    if ( $error || $messages ) {
        warn $messages;
    }
    if ( defined $destination && $source ne $destination ) {
        unlink "$_.tdy";
        link $_, "$_.tdy";
        rename "$_.tdy", "$_.bak";
        open my $g, '>:utf8', "$_.tdy";
        print $g /\.pm$/ ? "\N{U+FEFF}" : (), $destination;
        close $g;
        chmod( 07777 & ( stat $_ )[2], "$_.tdy" );
        rename "$_.tdy", $_;
    }

}
