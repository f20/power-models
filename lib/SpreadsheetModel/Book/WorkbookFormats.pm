﻿package SpreadsheetModel::Book::WorkbookFormats;

# Copyright 2008-2023 Franck Latrémolière and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=head Development note

Keys used from %$options:
* alignment
* colour
* lockedInputs
* validation (affects void cells in input data tables)

?,??0 style formats do not work with OpenOffice.org.  To work around this, set the alignment option to right[0-9]* (where the number is the amount of padding on the right).

=cut

use warnings;
use strict;
use Storable qw(nfreeze);
$Storable::canonical = 1;

sub getFormat {
    my ( $workbook, $key, @decorations ) = @_;
    $workbook->{formats}{ nfreeze( [ $key, @decorations ] ) } ||=
      UNIVERSAL::can( $key, 'copy' )
      ? do {
        my $f = $workbook->add_format;
        $f->copy($key);
        $f->set_format_properties(@$_)
          foreach grep { $_ } @{ $workbook->{decospec} }{@decorations};
        $f;
      }
      : $workbook->add_format(
        ref $key eq 'ARRAY'
        ? (
            $key->[0] eq 'base'
            ? (
                @{ $workbook->{formatspec}{ $key->[1] } },
                @{$key}[ 2 .. $#$key ]
              )
            : @$key
          )
        : @{
            $workbook->{formatspec}{ $key =~ /(.*)nz$/ ? $1 : $key }
              || die "$key is not a valid format"
        },
        map { $_ ? @$_ : (); } @{ $workbook->{decospec} }{@decorations},
      );
}

# These codes (between 8 and 63) are equal to 7 plus the colour number used in VBA (1-56)
use constant {
    EXCELCOL0 => 8,  #000000 Black
    WHITE     => 9,  #FFFFFF White
    RED       => 10, #FF0000 Red
    EXCELCOL3 => 11, #00FF00 Green or Lime or Bright Green
    BLUE      => 12, #0000FF Blue potentially overridden by #0066cc
    BGGOLD    => 13, #FFFF00 Yellow potentially overridden by #ffd700 or #fecb2f
    MAGENTA   => 14, #FF00FF Magenta
    CYAN      => 15, #00FFFF Cyan
    EXCELCOL8 => 16, #800000
    GREEN     => 17, #008000
    EXCELCOL10 => 18,    #000080
    EXCELCOL11 => 19,    #808000
    PURPLE     => 20,    #800080
    EXCELCOL13 => 21,    #008080
    SILVER     => 22,    #C0C0C0 potentially overridden by #e9e9e9
    EXCELCOL15 => 23,    #808080
    EXCELCOL16 => 24,    #9999FF chart fill
    EXCELCOL17 => 25,    #993366 chart fill
    EXCELCOL18 => 26,    #FFFFCC chart fill
    EXCELCOL19 => 27,    #CCFFFF chart fill
    EXCELCOL20 => 28,    #660066 chart fill
    EXCELCOL21 => 29,    #FF8080 chart fill
    EXCELCOL22 => 30,    #0066CC chart fill
    EXCELCOL23 => 31,    #CCCCFF chart fill
    EXCELCOL24 => 32,    #000080 chart line
    EXCELCOL25 => 33,    #FF00FF chart line
    EXCELCOL26 => 34,    #FFFF00 chart line
    EXCELCOL27 => 35,    #00FFFF chart line
    EXCELCOL28 => 36,    #800080 chart line
    EXCELCOL29 => 37,    #800000 chart line
    EXCELCOL30 => 38,    #008080 chart line
    EXCELCOL31 => 39,    #0000FF chart line
    EXCELCOL32 => 40,    #00CCFF
    BGBLUE     => 41,    #CCFFFF
    BGGREEN    => 42,    #CCFFCC
    BGYELLOW   => 43,    #FFFF99 potentially overridden by #ffffcc
    LTPURPLE   => 44,    #99CCFF potentially overridden by #fbf8ff
    BGPINK     => 45,    #FF99CC potentially overridden by #ffccff
    BGPURPLE   => 46,    #CC99FF potentially overridden by #eeddff
    EXCELCOL39 => 47,    #FFCC99
    EXCELCOL40 => 48,    #3366FF
    EXCELCOL41 => 49,    #33CCCC
    EXCELCOL42 => 50,    #99CC00
    EXCELCOL43 => 51,    #FFCC00
    BGORANGE   => 52,    #FF9900 potentially overridden by #ffcc99
    ORANGE     => 53,    #FF6600 potentially overridden by #ff6633
    EXCELCOL46 => 54,    #666699
    GREY       => 55,    #969696 potentially overridden by #999999
    EXCELCOL48 => 56,    #003366
    EXCELCOL49 => 57,    #339966
    EXCELCOL50 => 58,    #003300
    EXCELCOL51 => 59,    #333300
    EXCELCOL52 => 60,    #993300
    EXCELCOL53 => 61,    #993366
    EXCELCOL54 => 62,    #333399
    EXCELCOL55 => 63,    #333333
};

sub setFormats {

    my ( $workbook, $options ) = @_;

    # Ugly hack see https://github.com/jmcnamara/excel-writer-xlsx/issues/59
    $workbook->{_window_width}  = 1280 * 20;
    $workbook->{_window_height} = 800 * 20;
    $workbook->{_tab_ratio}     = 0.88 * 1000;

    my $noColour =
      $options->{colour} && $options->{colour} =~ /nocolour|striped/i;
    my $defaultColours =
      $noColour || $options->{colour} && $options->{colour} =~ /default/i;
    my $darkerBackgrounds = $options->{darkerBackgrounds};
    my $orangeColours = $options->{colour} && $options->{colour} =~ /orange/i;
    my $goldColours   = $options->{colour} && $options->{colour} =~ /gold/i;
    my $borderColour  = $options->{colour} && $options->{colour} =~ /border/i;
    my $textColour    = $options->{colour} && $options->{colour} =~ /text/i;
    my $backgroundColour = !$borderColour  && !$textColour;

    my $boldHeadings =
      $options->{colour} && $options->{colour} =~ /bold|orange/i;

    my $luridFonts = $options->{colour} && $options->{colour} =~ /striped/i;
    $workbook->{captionRowHeight} = $luridFonts ? 35 : 21;

    unless ($defaultColours) {
        if ($goldColours) {
            $workbook->set_custom_color( BGGOLD, '#ffd700' );
            $orangeColours = $goldColours;
        }
        if ($orangeColours) {
            $workbook->set_custom_color( BGORANGE, '#ffcc99' );
        }
        else {
            $workbook->set_custom_color( BGPURPLE,
                $darkerBackgrounds ? '#eeddff' : '#f7eeff' );
            $workbook->set_custom_color( LTPURPLE, '#fbf8ff' );
        }
        $workbook->set_custom_color( BGBLUE, '#eeffff' )
          unless $darkerBackgrounds;
        $workbook->set_custom_color( BGGREEN, '#eeffee' )
          unless $darkerBackgrounds;
        $workbook->set_custom_color( BGYELLOW,
            $darkerBackgrounds ? '#ffffcc' : '#ffffee' );
        $workbook->set_custom_color( SILVER,
            $darkerBackgrounds ? '#e9e9e9' : '#f4f4f4' );
        $workbook->set_custom_color( BGPINK, '#ffccff' );
        $workbook->set_custom_color( ORANGE, '#ff6633' );
        $workbook->set_custom_color( BLUE,   '#0066cc' );
        $workbook->set_custom_color( GREY,   '#999999' );
    }

    my $q3    = $options->{alignment}                        ? '?,' : '??,???,';
    my $q4    = $options->{alignment}                        ? '?,' : '??,';
    my $cyan  = $backgroundColour && !$options->{noCyanText} ? '[Cyan]'  : '';
    my $black = $backgroundColour                            ? '[Black]' : '';
    my $positiveCol = '[Blue]';
    my $negativeCol = '[Magenta]';
    my $rightpad;
    $rightpad = '_)' x ( $1 || 2 )
      if $options->{alignment} && $options->{alignment} =~ /right.*?([0-9]*)/;

    my @alignText = $options->{alignText}
      && $options->{alignText} =~ /general/i ? () : ( align => 'left' );
    my $plus  = $positiveCol . '_-+';
    my $minus = $negativeCol . '_+-';
    my $same  = $rightpad ? "[Green]=_)$rightpad" : '[Green]=';

    my $num_general  = '[Black]General';
    my $num_text     = "${positiveCol}General;${negativeCol}-General;;[Black]@";
    my $num_textonly = '[Black]General;[Black]-General;;[Black]@';
    my $num_textonlycopy = '[Black]General;[Black]-General;;[Black]@';
    my $num_mpan = "${black}00 0000 0000 000;${negativeCol}-General;;$cyan@";

    my @num_percent =
      $rightpad
      ? (
        num_format =>
          "${black}#,##0.00%_)$rightpad;${negativeCol}(#,##0.00%)$rightpad"
          . ";;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format =>
          "${black} _(??,??0.00%_);${negativeCol} (??,??0.00%);;$cyan@",
        align => 'center'
      );
    my @num_percentpm =
      $rightpad
      ? (
        num_format => "$plus#,##0.00%_)$rightpad;$minus#,##0.00%_)$rightpad"
          . ";$same;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format => "$plus?,??0.00%;$minus?,??0.00%;$same;$cyan@",
        align      => 'center'
      );
    my @num_million =
      $rightpad
      ? (
        num_format => qq'${black}_(#,##0.0,,"m"_)$rightpad;'
          . qq'${negativeCol}(#,##0.0,,"m")$rightpad;;$cyan\@$rightpad',
        align => 'right'
      )
      : (
        num_format =>
          qq'${black} _(?,??0.0,,"m"_);${negativeCol} (?,??0.0,,"m");;$cyan@',
        align => 'center'
      );
    my @num_date =
      $rightpad
      ? (
        num_format => qq'${black}d mmm yyyy$rightpad;'
          . qq'${negativeCol}-General$rightpad;;$cyan\@$rightpad',
        align => 'right'
      )
      : (
        num_format => qq'${black}d mmm yyyy;${negativeCol}-General;;$cyan@',
        align      => 'center'
      );
    my @num_datetime =
      $rightpad
      ? (
        num_format =>
qq'${black}ddd d mmm yyyy HH:mm;${negativeCol}-General;;$cyan\@$rightpad',
        align => 'right'
      )
      : (
        num_format =>
          qq'${black}ddd d mmm yyyy  HH:mm  ;${negativeCol}-General  ;;$cyan@',
        align => 'right'
      );
    my @num_time =
      $rightpad
      ? (
        num_format => qq'${black}[hh]:mm$rightpad;'
          . qq'${negativeCol}-General$rightpad;${black}[hh]:mm$rightpad;$cyan\@$rightpad',
        align => 'right'
      )
      : (
        num_format =>
          qq'${black}[hh]:mm;${negativeCol}-General;${black}[hh]:mm;$cyan@',
        align => 'center'
      );
    my @num_monthday =
      $rightpad
      ? (
        num_format => qq'${black}mmmm d$rightpad;'
          . qq'${negativeCol}-General$rightpad;;$cyan\@$rightpad',
        align => 'right'
      )
      : (
        num_format => qq'${black}mmmm d;${negativeCol}-General;;$cyan@',
        align      => 'center'
      );
    my @num_ =
      $rightpad
      ? (
        num_format =>
"${black}#,##0_)$rightpad;${negativeCol}(#,##0)$rightpad;;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format => "${black} _(?$q3??0_);${negativeCol} (?$q3??0);;$cyan@",
        align      => 'center'
      );
    my @num_pm =
      $rightpad
      ? (
        num_format => "$plus#,##0_)$rightpad;$minus#,##0_)$rightpad"
          . ";$same;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format => "$plus$q3??0;$minus$q3??0;$same;$cyan@",
        align      => 'center'
      );
    my @num_0 =
      $rightpad
      ? (
        num_format =>
          "${black}#,##0.0_)$rightpad;${negativeCol}(#,##0.0)$rightpad"
          . ";;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format =>
          "${black} _(?$q4??0.0_);${negativeCol} (?$q4??0.0);;$cyan@",
        align => 'center'
      );
    my @num_0pm =
      $rightpad
      ? (
        num_format =>
"$plus#,##0.0_)$rightpad;$minus#,##0.0_)$rightpad;$same;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format => "$plus$q4??0.0;$minus$q4??0.0;$same;$cyan@",
        align      => 'center'
      );
    my @num_00 =
      $rightpad
      ? (
        num_format =>
          "${black}#,##0.00_)$rightpad;${negativeCol}(#,##0.00)$rightpad"
          . ";;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format =>
          "${black} _(?$q4??0.00_);${negativeCol} (?$q4??0.00);;$cyan@",
        align => 'center'
      );
    my @num_00pm =
      $rightpad
      ? (
        num_format => "$plus#,##0.00_)$rightpad;$minus#,##0.00_)$rightpad"
          . ";$same;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format => "$plus$q4??0.00;$minus$q4??0.00;$same;$cyan@",
        align      => 'center'
      );
    my @num_000 =
      $rightpad
      ? (
        num_format =>
          "${black}#,##0.000_)$rightpad;${negativeCol}(#,##0.000)$rightpad"
          . ";;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format =>
          "${black} _(?$q4??0.000_);${negativeCol} (?$q4??0.000);;$cyan@",
        align => 'center'
      );
    my @num_000pm =
      $rightpad
      ? (
        num_format => "$plus#,##0.000_)$rightpad;$minus#,##0.000_)$rightpad"
          . ";$same;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format => "$plus$q4??0.000;$minus$q4??0.000;$same;$cyan@",
        align      => 'center'
      );
    my @num_00000 =
      $rightpad
      ? (
        num_format =>
          "${black}#,##0.00000_)$rightpad;${negativeCol}(#,##0.00000)$rightpad"
          . ";;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format =>
          "${black} _(?$q4??0.00000_);${negativeCol} (?$q4??0.00000);;$cyan@",
        align => 'center'
      );
    my @num_quad =
      $rightpad
      ? (
        num_format => "${black}0000_)$rightpad;-0000_)$rightpad;"
          . "${black}0000_)$rightpad;$cyan\@$rightpad",
        align => 'right'
      )
      : (
        num_format => "${black}0000;-0000;${black}0000;$cyan@",
        align      => 'center'
      );

    my @cDefault =
      $noColour
      ? ( bottom => 1 )
      : (
        $options->{gridlines}                        ? ( border => 7 ) : (),
        $backgroundColour && !$options->{noCyanText} ? ( color => MAGENTA ) : ()
      );
    my @cCon = (
        locked => 1,
        $noColour ? ( bottom => 1, border_color => GREY )
        : (
            $options->{gridlines} ? ( border   => 7 ) : (),
            $backgroundColour     ? ( bg_color => SILVER, @cDefault, )
            : (
                $borderColour ? ( border => 1, border_color => GREY ) : (),
                $textColour   ? ( color  => GREY )                    : (),
            )
        )
    );
    my @cCopy = (
        locked => 1,
        $noColour ? ( bottom => 3 )
        : (
            $options->{gridlines} ? ( border   => 7 ) : (),
            $backgroundColour     ? ( bg_color => BGGREEN, @cDefault, )
            : (
                $borderColour ? ( border => 1, border_color => GREEN ) : (),
                $textColour   ? ( color  => GREEN )                    : (),
            )
        )
    );
    my @cHard = (
        locked => 0,
        $noColour ? ( bottom => 2 )
        : (
            $options->{gridlines} ? ( border   => 7 ) : (),
            $backgroundColour     ? ( bg_color => BGBLUE, @cDefault, )
            : (
                $borderColour ? ( border => 1, border_color => BLUE ) : (),
                $textColour   ? ( color  => BLUE )                    : (),
            )
        )
    );
    my @cSoft = (
        locked => 1,
        $noColour ? ( bottom => 1 )
        : (
            $options->{gridlines} ? ( border   => 7 ) : (),
            $backgroundColour     ? ( bg_color => BGYELLOW, @cDefault, )
            : (
                $borderColour ? ( border => 1, border_color => EXCELCOL11 )
                : (),
                $textColour ? ( color => EXCELCOL11 ) : (),
            )
        )
    );
    my @cScribbles = (
        locked => 0,
        $noColour ? ( bottom => 4, border_color => SILVER, )
        : (
            color => PURPLE,
            !$backgroundColour ? ()
            : $orangeColours
            ? ( bottom => 3, top => 3, border_color => PURPLE, )
            : 1 ? ()
            :     ( bg_color => LTPURPLE )
        )
    );
    my @cHeader = (
        locked => 1,
        $noColour
        ? ( bottom => 1, border_color => GREY )
        : (
            $goldColours
            ? ( bg_color => SILVER, fg_color => BGGOLD, pattern => 6 )
            : $backgroundColour
            ? ( bg_color => $orangeColours ? BGORANGE : BGPURPLE )
            : ()
        )
    );
    my @cUnavailable = (
        locked => 1,
        $noColour ? ( bottom => 9, border_color => GREY )
        : (
            $options->{gridlines} ? ( border => 7 ) : (),
            $backgroundColour
            ? ( fg_color => SILVER, bg_color => WHITE, pattern => 14, )
            : ( right => 4, border_color => GREY )
        )
    );
    my @cUnused =
      $noColour ? ( bottom => 11, border_color => GREY )
      : (
        $options->{gridlines} ? ( border => 7 ) : (),
        $backgroundColour
        ? ( fg_color => SILVER, bg_color => WHITE, pattern => 15, )
        : ( right => 4, border_color => GREY )
      );
    my @cCaption = (
        locked => 1,
        $noColour || $backgroundColour ? () : ( color => BLUE )
    );
    my @cTitle = (
        locked => 1,
        $noColour || $backgroundColour ? () : ( color => ORANGE )
    );
    my @sExtras =
      $goldColours ? ( top => 1, bottom => 1, border_color => BGGOLD ) : ();
    my @sCaption =
      $luridFonts
      ? ( font => 'Baskerville', size => 25, )
      : ( size => 15, bold => 1, );
    my @sHeading = (
        @sCaption,    # valign => 'vbottom'
    );
    my @sLabel      = ( valign => 1 ? 'vbottom' : 'vcenter', @sExtras );
    my @sLabelGroup = ( valign => 1 ? 'vbottom' : 'vcenter', );
    my @sNumber     = ( valign => 1 ? 'vbottom' : 'vcenter', @sExtras );
    my @sText       = ( valign => 1 ? 'vbottom' : 'vcenter', );
    my %specs       = (
        '%con'         => [ @sNumber, @num_percent,   @cCon, ],
        '%copy'        => [ @sNumber, @num_percent,   @cCopy, ],
        '%copypm'      => [ @sNumber, @num_percentpm, @cCopy, ],
        '%hard'        => [ @sNumber, @num_percent,   @cHard, ],
        '%hardpm'      => [ @sNumber, @num_percentpm, @cHard, ],
        '%soft'        => [ @sNumber, @num_percent,   @cSoft, ],
        '%softpm'      => [ @sNumber, @num_percentpm, @cSoft, ],
        '0.00000con'   => [ @sNumber, @num_00000,     @cCon, ],
        '0.00000copy'  => [ @sNumber, @num_00000,     @cCopy, ],
        '0.00000soft'  => [ @sNumber, @num_00000,     @cSoft, ],
        '0.000con'     => [ @sNumber, @num_000,       @cCon, ],
        '0.000copy'    => [ @sNumber, @num_000,       @cCopy, ],
        '0.000hard'    => [ @sNumber, @num_000,       @cHard, ],
        '0.000hardpm'  => [ @sNumber, @num_000pm,     @cHard, ],
        '0.000soft'    => [ @sNumber, @num_000,       @cSoft, ],
        '0.000softpm'  => [ @sNumber, @num_000pm,     @cSoft, ],
        '0.00con'      => [ @sNumber, @num_00,        @cCon, ],
        '0.00copy'     => [ @sNumber, @num_00,        @cCopy, ],
        '0.00hard'     => [ @sNumber, @num_00,        @cHard, ],
        '0.00hardpm'   => [ @sNumber, @num_00pm,      @cHard, ],
        '0.00soft'     => [ @sNumber, @num_00,        @cSoft, ],
        '0.00softpm'   => [ @sNumber, @num_00pm,      @cSoft, ],
        '0.0con'       => [ @sNumber, @num_0,         @cCon, ],
        '0.0copy'      => [ @sNumber, @num_0,         @cCopy, ],
        '0.0hard'      => [ @sNumber, @num_0,         @cHard, ],
        '0.0soft'      => [ @sNumber, @num_0,         @cSoft, ],
        '0.0softpm'    => [ @sNumber, @num_0pm,       @cSoft, ],
        '0000con'      => [ @sNumber, @num_quad,      @cCon, ],
        '0000copy'     => [ @sNumber, @num_quad,      @cCopy, ],
        '0000hard'     => [ @sNumber, @num_quad,      @cHard, ],
        '0000soft'     => [ @sNumber, @num_quad,      @cSoft, ],
        '0con'         => [ @sNumber, @num_,          @cCon, ],
        '0copy'        => [ @sNumber, @num_,          @cCopy, ],
        '0copypm'      => [ @sNumber, @num_pm,        @cCopy, ],
        '0hard'        => [ @sNumber, @num_,          @cHard, ],
        '0hardpm'      => [ @sNumber, @num_pm,        @cHard, ],
        '0soft'        => [ @sNumber, @num_,          @cSoft, ],
        '0softpm'      => [ @sNumber, @num_pm,        @cSoft, ],
        'datecon'      => [ @sNumber, @num_date,      @cCon, ],
        'datecopy'     => [ @sNumber, @num_date,      @cCopy, ],
        'datehard'     => [ @sNumber, @num_date,      @cHard, ],
        'datesoft'     => [ @sNumber, @num_date,      @cSoft, ],
        'datetimecon'  => [ @sNumber, @num_datetime,  @cCon, ],
        'datetimecopy' => [ @sNumber, @num_datetime,  @cCopy, ],
        'datetimehard' => [ @sNumber, @num_datetime,  @cHard, ],
        'datetimesoft' => [ @sNumber, @num_datetime,  @cSoft, ],
        'millioncon'   => [ @sNumber, @num_million,   @cCon, ],
        'millioncopy'  => [ @sNumber, @num_million,   @cCopy, ],
        'millionhard'  => [ @sNumber, @num_million,   @cHard, ],
        'millionsoft'  => [ @sNumber, @num_million,   @cSoft, ],
        'monthdaycon'  => [ @sNumber, @num_monthday,  @cCon, ],
        'monthdaycopy' => [ @sNumber, @num_monthday,  @cCopy, ],
        'monthdayhard' => [ @sNumber, @num_monthday,  @cHard, ],
        'monthdaysoft' => [ @sNumber, @num_monthday,  @cSoft, ],
        'timecon'      => [ @sNumber, @num_time,      @cCon, ],
        'timecopy'     => [ @sNumber, @num_time,      @cCopy, ],
        'timehard'     => [ @sNumber, @num_time,      @cHard, ],
        'timesoft'     => [ @sNumber, @num_time,      @cSoft, ],
        'boolhard'     => [
            @sNumber,
            align      => 'center',
            num_format => $num_general,
            @cHard,
        ],
        'boolsoft' => [
            @sNumber,
            align      => 'center',
            num_format => $num_general,
            @cSoft,
        ],
        'boolcopy' => [
            @sNumber,
            align      => 'center',
            num_format => $num_general,
            @cCopy,
        ],
        'caption' => [
            @sCaption,
            num_format => '@',
            align      => 'left',
            @cCaption,
        ],
        'captionca' => [
            @sCaption,
            num_format => '@',
            text_wrap  => 1,
            align      => 'center_across',
            @cCaption,
        ],
        'link' => [
            locked => 1,
            @sText,
            num_format => '@',
            underline  => 1,
            color      => BLUE,
        ],
        'loccopy' => [
            @sLabel,
            num_format => $black . '\L\o\c\a\t\i\o\n\ 0',
            align      => 'center',
            @cCopy,
        ],
        'locsoft' => [
            @sLabel,
            num_format => $black . '\L\o\c\a\t\i\o\n\ 0',
            align      => 'center',
            @cSoft,
        ],
        'indexcon' => [
            locked => 1,
            @sLabel,
            num_format => $black . '\I\n\d\e\x\ 0;' . $black . '\I\n\d\e\x\ -0',
            align      => 'center',
            $options->{gridlines} ? ( border => 7 ) : (),
            @cCon,
        ],
        'indexcopy' => [
            locked => 1,
            @sLabel,
            num_format => $black . '\I\n\d\e\x\ 0;' . $black . '\I\n\d\e\x\ -0',
            align      => 'center',
            $options->{gridlines} ? ( border => 7 ) : (),
            @cCopy,
        ],
        'indexsoft' => [
            locked => 1,
            @sLabel,
            num_format => $black . '\I\n\d\e\x\ 0;' . $black . '\I\n\d\e\x\ -0',
            align      => 'center',
            $options->{gridlines} ? ( border => 7 ) : (),
            @cSoft,
        ],
        'stephard' => [
            @sLabel,
            num_format => $black . '\S\t\e\p\ 0;' . $black . '-0',
            align      => 'center',
            @cHard,
        ],
        'notes' => [
            @sHeading,
            num_format => '@',
            align      => 'left',
            @cTitle,
        ],
        'scribbles' => [
            @sText,
            text_wrap => 0,
            @cScribbles,
        ],
        'code' => [
            locked => 1,
            @sText,
            num_format => $num_textonly,
            align      => 'left',
            0 ? ( font => 'Consolas' ) : 0 ? ( font => 'Courier New' ) : (),
            @cDefault,
        ],
        'mpanhard' => [
            @sText,
            num_format => $num_mpan,
            align      => 'center',
            @sExtras,
            text_wrap => 1,
            @cHard,
        ],
        'mpancopy' => [
            @sText,
            num_format => $num_mpan,
            align      => 'center',
            @sExtras,
            text_wrap => 1,
            @cCopy,
        ],
        'puretextcon' => [
            @sText,
            num_format => $num_textonlycopy,
            align      => 'center',
            @sExtras,
            text_wrap => 1,
            @cCon,
        ],
        'puretextcopy' => [
            @sText,
            num_format => $num_textonlycopy,
            align      => 'center',
            @sExtras,
            text_wrap => 1,
            @cCopy,
        ],
        'puretexthard' => [
            @sText,
            num_format => $num_textonly,
            align      => 'center',
            @sExtras,
            text_wrap => 1,
            @cHard,
        ],
        'puretextsoft' => [
            @sText,
            num_format => $num_textonly,
            align      => 'center',
            @sExtras,
            text_wrap => 1,
            @cSoft,
        ],
        'text' => [
            locked => 1,
            @sText,
            num_format => '@',
        ],
        'textcon' => [
            @sText,
            num_format => $num_text,
            @alignText,
            text_wrap => 1,
            @sExtras,
            @cCon,
        ],
        'textcopy' => [
            @sText,
            num_format => $num_text,
            @alignText, @sExtras,
            text_wrap => 1,
            @cCopy,
        ],
        'textcopycentered' => [
            @sText,
            num_format => $num_text,
            align      => 'center',
            @sExtras,
            text_wrap => 1,
            @cCopy,
        ],
        'texthard' => [
            @sText,
            num_format => $num_text,
            @alignText, @sExtras,
            text_wrap => 1,
            @cHard,
        ],
        'texthardnowrap' => [
            @sText,
            num_format => $num_text,
            @alignText, @sExtras,
            @cHard,
        ],
        'textsoft' => [
            @sText,
            num_format => $num_text,
            @alignText, @sExtras,
            text_wrap => 1,
            @cSoft,
        ],
        'textnocolour' => [
            locked => 0,
            @sText,
            num_format => $num_text,
            @alignText, @sExtras,
            text_wrap => 1,
            @cDefault,
        ],
        'th' => [
            @sLabel,
            num_format => $num_textonly,
            align      => 'left',
            $boldHeadings ? ( bold => 1 ) : (),
            text_wrap => 1,
            @cHeader,
            $options->{gridlines} ? ( bottom => 7, right => 1 ) : (),
        ],
        'thc' => [
            @sLabel,
            num_format => $num_textonly,
            $boldHeadings ? ( bold => 1 ) : (),
            text_wrap => 1,
            align     => 'center',
            @cHeader,
            $options->{gridlines} ? ( right => 7, bottom => 1 ) : (),
        ],
        'thca' => [
            @sLabelGroup,
            num_format => $num_textonly,
            italic     => 1,
            text_wrap  => 1,
            align      => 'center_across',
            @cHeader,
            $options->{gridlines}
            ? ( right => 7, bottom => 1, )
            : ( left => 1, right => 1, ),
        ],
        'thcaleft' => [
            @sLabelGroup,
            num_format => $num_textonly,
            align      => 'left',
            italic     => 1,
            @cHeader,
        ],
        'thg' => [
            @sLabelGroup,
            num_format => $num_textonly,
            align      => 'left',
            italic     => 1,
            text_wrap  => 1,
            @cHeader,
        ],
        'colnoteleft' => [
            locked => 1,
            @sText,
            num_format => '@',
            align      => 'left',
        ],
        'colnoteleftwrap' => [
            locked => 1,
            @sText,
            num_format => '@',
            text_wrap  => 1,
            align      => 'left',
        ],
        'colnotecenter' => [
            locked => 1,
            @sText,
            num_format => '@',
            text_wrap  => 1,
            align      => 'center_across',
            left       => 1,
            right      => 1,
        ],
        'thitem' => [
            @sLabel,
            num_format => '\I\t\e\m\ \#0',
            align      => 'left',
            $boldHeadings ? ( bold => 1 ) : (),
            @cHeader,
            $options->{gridlines} ? ( bottom => 7, right => 1 ) : (),
        ],
        'thloc' => [
            @sLabel,
            num_format => '\L\o\c\a\t\i\o\n\ 0',
            align      => 'left',
            $boldHeadings ? ( bold => 1 ) : (),
            @cHeader,
            $options->{gridlines} ? ( bottom => 7, right => 1 ) : (),
        ],
        'thtar' => [
            @sLabel,
            num_format => '\T\a\r\i\f\f\ 0',
            align      => 'left',
            $boldHeadings ? ( bold => 1 ) : (),
            @cHeader,
            $options->{gridlines} ? ( bottom => 7, right => 1 ) : (),
        ],
        'tarhard' => [
            @sNumber,
            align      => 'left',
            num_format =>
              "[Black]\\T\\a\\r\\i\\f\\f\\ 0;${negativeCol}-0;;[Cyan]@",
            @cHard,
        ],
        'unavailable' => [
            @sNumber,
            num_format => $num_general,
            align      => 'center',
            @cUnavailable,
        ],
        'unused' => [
            locked => !$options->{validation}
              || $options->{validation} !~ /lenient/i ? 1 : 0,
            @sNumber,
            num_format => $num_general,
            align      => 'center',
            @cUnused,
        ],
    );

    if ( $options->{lockedInputs} ) {
        $specs{unused} = $specs{unavailable};
        foreach my $key ( grep { /hard/ } keys %specs ) {
            local $_ = $key;
            s/hard/input/;
            $specs{$_} = $specs{$key};
            $_ = $key;
            s/hard/con/;
            $specs{$key} = $specs{$_} if $specs{$_};
        }
    }

    $workbook->{formatspec} = \%specs;

    $workbook->{decospec} = {
        bold   => [ bold => 1, ],
        wrapca => [
            text_wrap     => 1,
            center_across => 1,
            $options->{gridlines}
            ? ( right => 7, bottom => 1, )
            : ( left => 1, right => 1, ),
        ],
        tlttr => [    # thick line to the right
            right       => 5,
            right_color => 8,
        ],
        dottedlineunder => [
            bottom       => 4,
            bottom_color => 8,
        ],
        blue => [
            color    => $orangeColours ? BGORANGE : BGPURPLE,
            size     => 13,
            bg_color => BLUE,
        ],
        red => [
            color    => $orangeColours ? BGORANGE : BGPURPLE,
            size     => 13,
            bg_color => RED,
        ],
        algae => [
            color    => $orangeColours ? BGORANGE : BGPURPLE,
            size     => 13,
            bg_color => EXCELCOL13,
        ],
        purple => [
            color    => $orangeColours ? BGORANGE : BGPURPLE,
            size     => 13,
            bg_color => PURPLE,
        ],
        slime => [
            color    => $orangeColours ? BGORANGE : BGPURPLE,
            size     => 13,
            bg_color => EXCELCOL11,
        ],
    };

    if ( $options->{colour} && $options->{colour} =~ /striped/i ) {
        my @stripes;
        foreach (
            [ 'checksum',              '#ffff99' ],
            [ 'GSP|Transmission exit', '#ccccff' ],
            [ '132.*EHV',              '#ccffe6' ],
            [ '132.*HV',               '#e6e6e6' ],
            [ 'EHV.*HV',               '#e6e6cc' ],
            [ 'HV.*LV',                '#ffe6cc' ],
            [ '132',                   '#ccffff' ],
            [ 'EHV',                   '#ccffcc' ],
            [ 'HV',                    '#ffcccc' ],
            [ 'LV',                    '#ffffcc' ],
            [ 'Red'                   => '#ffcccc' ],
            [ 'Amber'                 => '#ffe6cc' ],
            [ 'Green'                 => '#ccffcc' ],
            [ 'Yellow'                => '#ffffcc' ],
            [ 'Black'                 => '#cccccc' ],
            [ 'adder|scaler|matching' => '#ffccff' ],
            [ 'rate 1'                => '#e6cccc' ],
            [ 'rate 2'                => '#ffeacc' ],
            [ 'rate 3'                => '#ccffcc' ],
            [ 'capacity'              => '#ccffff' ],
            [ 'reactive'              => '#ccccff' ],
            [ '.'                     => '#e6e6e6' ],
          )
        {
            push @stripes, [ $_->[0], "inv$_->[1]", $_->[1] ];
            $workbook->{decospec}{ $_->[1] } ||= [ bg_color => $_->[1] ];
            $workbook->{decospec}{ 'inv' . $_->[1] } ||=
              [ bg_color => 'black', color => $_->[1], num_format => '@', ];
        }
        $workbook->{columnDecorations} = sub {
            my @decos;
            foreach (@_) {
                my $colFormatSpec;
                foreach my $pattern (@stripes) {
                    if (/$pattern->[0]/i) {
                        $colFormatSpec = [ $pattern->[1], $pattern->[2] ];
                        last;
                    }
                }
                push @decos, $colFormatSpec;
            }
            return unless grep { $_; } @decos;
            @decos;
        };
    }

}

1;
