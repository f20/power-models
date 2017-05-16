package SpreadsheetModel::CLI::CommandRunner;

=head Copyright licence and disclaimer

Copyright 2011-2017 Franck Latrémolière and others. All rights reserved.

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

use SpreadsheetModel::CLI::MakeModels;
use SpreadsheetModel::CLI::UseDatabase;
use SpreadsheetModel::CLI::UseModels;
use SpreadsheetModel::Data::DataTools;
use SpreadsheetModel::Rules::RulesTools;

use constant {
    C_HOMEDIR       => 0,
    C_VALIDATEDLIBS => 1,
    C_DESTINATION   => 2,
    C_LOG           => 3,
};

sub new {
    my $class = shift;
    bless [@_], $class;
}

sub finish {
    my ($self) = @_;
    $self->makeFolder;
}

sub log {
    my ( $self, $verb, @objects ) = @_;
    warn "$verb: @objects\n";
    return if $verb eq 'makeFolder';
    push @{ $self->[C_LOG] },
      join( "\n", $verb, map { "\t$_"; } @objects ) . "\n\n";
}

sub makeFolder {

    my ( $self, $folder ) = @_;

    if ( $self->[C_DESTINATION] ) {    # Close out previous folder
        return if $folder && $folder eq $self->[C_DESTINATION];
        if ( $self->[C_LOG] ) {
            my $tmpFile = '~$tmptxt' . $$;
            open my $h, '>', $tmpFile;
            print {$h} @{ $self->[C_LOG] };
            close $h;
            local $_ = "$self->[C_DESTINATION].txt";
            s/^_+([^\.])/$1/s;
            rename $tmpFile, $_;
            delete $self->[C_LOG];
        }
        chdir '..';
        my $tmp = '~$tmp-' . $$ . ' ' . $self->[C_DESTINATION];
        return if rmdir $tmp;
        rename $self->[C_DESTINATION], $tmp . '/~$old-' . $$
          if -e $self->[C_DESTINATION];
        rename $tmp, $self->[C_DESTINATION];
        system 'open', $self->[C_DESTINATION] if -d '/System/Library';   # macOS
        delete $self->[C_DESTINATION];
    }

    if ($folder) {    # Create temporary folder and go there

        if ( -d '/System/Library' ) {    # macOS: use a memory disk
            my $ramDiskBlocks = 12_000_000;    # About 6G, in 512-byte blocks.
            my $ramDiskName = 'power-models workings';
            my $ramDiskMountPoint = 1 ? "/Volumes/$ramDiskName" : $ramDiskName;
            unless ( -e "$ramDiskMountPoint/.VolumeIcon.icns" ) {

                my $device = `hdiutil attach -nomount ram://$ramDiskBlocks`;
                $device =~ s/\s*$//s;

                if ( $ramDiskMountPoint =~ m#^/Volumes/#s ) {
                    system qw(diskutil erasevolume HFS+), $ramDiskName, $device;
                }
                else {
                    system qw(newfs_hfs -v), $ramDiskName, $device;
                    mkdir $ramDiskMountPoint;
                    system qw(mount -o nobrowse -t hfs), $device,
                      $ramDiskMountPoint;
                }

                my $ramDiskIcns =
                  "$ENV{HOME}/Pictures/Images/power-models RAM disk icon.icns";
                if ( -e $ramDiskIcns ) {
                    system qw(cp), $ramDiskIcns,
                      "$ramDiskMountPoint/.VolumeIcon.icns";
                    system qw(SetFile -a C), $ramDiskMountPoint;
                }

            }
            chdir $ramDiskMountPoint if -d $ramDiskMountPoint && -w _;
        }

        my $tmp = '~$tmp-' . $$ . ' ' . ( $self->[C_DESTINATION] = $folder );
        mkdir $tmp;
        chdir $tmp;

    }

}

sub R {
    my ( $self, @commands ) = @_;
    open my $r, '| R --vanilla --slave';
    binmode $r, ':utf8';
    require SpreadsheetModel::Data::RCode;
    print {$r} SpreadsheetModel::Data::RCode->rCode(@commands);
}

sub Rcode {
    my ( $self, @commands ) = @_;
    open my $r, '>', "$$.R";
    binmode $r, ':utf8';
    print $r "# R code from power-models\n\n";
    require SpreadsheetModel::Data::RCode;
    print {$r} SpreadsheetModel::Data::RCode->rCode(@commands);
    close $r;
    rename "$$.R", 'power-models.R';
    warn <<EOW
To use this R code, say:
    source("power-models.R");
EOW
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

1;
