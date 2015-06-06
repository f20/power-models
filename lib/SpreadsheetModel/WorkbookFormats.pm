package SpreadsheetModel::WorkbookFormats;

=head Copyright licence and disclaimer

Copyright 2008-2015 Reckon LLP and others.

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

# SpreadsheetModel::Workbook inherits from this class.

use warnings;
use strict;

sub getFormat {
    my ( $workbook, $key, @decorations ) = @_;
    $workbook->{formats}{ join ' ', $key, @decorations } ||=
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

# These codes (between 8 and 63) are equal to the colour number used in VBA (1-56) plus 7
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

=head setFormats

Keys used in %$options:
* alignment (flag: true means right-aligned, false means centered with padding)
* colour (matched: see below)
* validation (matched: for void cells in input data tables)
* lockedInputs

=cut

    my $defaultColours = $options->{colour} && $options->{colour} =~ /default/i;
    my $orangeColours  = $options->{colour} && $options->{colour} =~ /orange/i;
    my $goldColours    = $options->{colour} && $options->{colour} =~ /gold/i;
    my $borderColour   = $options->{colour} && $options->{colour} =~ /border/i;
    my $textColour     = $options->{colour} && $options->{colour} =~ /text/i;
    my $backgroundColour = !$borderColour && !$textColour;

    unless ($defaultColours) {
        if ($goldColours) {
            $workbook->set_custom_color( BGGOLD, '#ffd700' );
            $orangeColours = $goldColours;
        }
        if ($orangeColours) {
            $workbook->set_custom_color( BGORANGE, '#ffcc99' );
        }
        else {
            $workbook->set_custom_color( BGPURPLE, '#eeddff' );
            $workbook->set_custom_color( LTPURPLE, '#fbf8ff' );
        }
        $workbook->set_custom_color( BGPINK,   '#ffccff' );
        $workbook->set_custom_color( BGYELLOW, '#ffffcc' );
        $workbook->set_custom_color( ORANGE,   '#ff6633' );
        $workbook->set_custom_color( BLUE,     '#0066cc' );
        $workbook->set_custom_color( GREY,     '#999999' );
        $workbook->set_custom_color( SILVER,   '#e9e9e9' );
    }

# ?,??0 style formats do not work with OpenOffice.org.
# Use "alignment: right[0-9]*" where the number is the amount of padding on the right.
    my $q3 = $options->{alignment} ? ',' : '??,???,';
    my $cyan = $backgroundColour && !$options->{noCyanText} ? '[Cyan]' : '';
    my $black = $backgroundColour ? '[Black]' : '';
    my $rightpad;
    $rightpad = '_)' x ( $1 || 2 )
      if $options->{alignment} && $options->{alignment} =~ /right.*?([0-9]*)/;

    my @alignText = ( align => 'left' );
    my $numText = $backgroundColour ? '[Blue]0;[Red]-0;;[Black]@' : '@';
    my $numTextOnly = $backgroundColour ? '[Black]@' : '@';
    if ( $options->{alignText} && $options->{alignText} =~ /general/i ) {
        @alignText = ();
        $numText = $backgroundColour ? '[Black]0;[Red]-0;;[Black]@' : '@';
    }

    my @numPercent =
      $rightpad
      ? (
        num_format => "${black}0.0%_)$rightpad;[Red](??0.0%)$rightpad;;$cyan@",
        align      => 'right'
      )
      : (
        num_format => "${black} _(??0.0%_);[Red] (??0.0%);;$cyan@",
        align      => 'center'
      );
    my @num_million =
      $rightpad
      ? (
        num_format => qq'${black}_(#,##0.0,, "m"_)$rightpad;'
          . qq'[Red](#,##0.0,, "m")$rightpad;;$cyan@',
        align => 'right'
      )
      : (
        num_format =>
          qq'${black} _(?,??0.0,, "m"_);[Red] (?,??0.0,, "m");;$cyan@',
        align => 'center'
      );
    my @num_ =
      $rightpad
      ? (
        num_format => "${black}#,##0_)$rightpad;[Red](#,##0)$rightpad;;$cyan@",
        align      => 'right'
      )
      : (
        num_format => "${black} _(?$q3??0_);[Red] (?$q3??0);;$cyan@",
        align      => 'center'
      );
    my @num_0 =
      $rightpad
      ? (
        num_format =>
          "${black}#,##0.0_)$rightpad;[Red](#,##0.0)$rightpad;;$cyan@",
        align => 'right'
      )
      : (
        num_format => "${black} _(?$q3??0.0_);[Red] (?$q3??0.0);;$cyan@",
        align      => 'center'
      );
    my @num_00 =
      $rightpad
      ? (
        num_format =>
          "${black}#,##0.00_)$rightpad;[Red](#,##0.00)$rightpad;;$cyan@",
        align => 'right'
      )
      : (
        num_format => "${black} _(?$q3??0.00_);[Red] (?$q3??0.00);;$cyan@",
        align      => 'center'
      );
    my @num_000 =
      $rightpad
      ? (
        num_format =>
          "${black}#,##0.000_)$rightpad;[Red](#,##0.000)$rightpad;;$cyan@",
        align => 'right'
      )
      : (
        num_format => "${black} _(?$q3??0.000_);[Red] (?$q3??0.000);;$cyan@",
        align      => 'center'
      );
    my @num_00000 =
      $rightpad
      ? (
        num_format =>
          "${black}#,##0.00000_)$rightpad;[Red](#,##0.00000)$rightpad;;$cyan@",
        align => 'right'
      )
      : (
        num_format =>
          "${black} _(?$q3??0.00000_);[Red] (?$q3??0.00000);;$cyan@",
        align => 'center'
      );

    my @defaultColour =
      $backgroundColour && !$options->{noCyanText} ? ( color => MAGENTA ) : ();
    my @colourCon = (
        $options->{gridlines} ? ( border => 7 ) : (),
        $backgroundColour
        ? ( bg_color => SILVER, @defaultColour, )
        : (
            $borderColour ? ( border => 1, border_color => GREY ) : (),
            $textColour ? ( color => GREY ) : (),
        )
    );
    my @colourCopy = (
        $options->{gridlines} ? ( border => 7 ) : (),
        $backgroundColour
        ? ( bg_color => BGGREEN, @defaultColour, )
        : (
            $borderColour ? ( border => 1, border_color => GREEN ) : (),
            $textColour ? ( color => GREEN ) : (),
        )
    );
    my @colourHard = (
        $options->{gridlines} ? ( border => 7 ) : (),
        $backgroundColour
        ? ( bg_color => BGBLUE, @defaultColour, )
        : (
            $borderColour ? ( border => 1, border_color => BLUE ) : (),
            $textColour ? ( color => BLUE ) : (),
        )
    );
    my @colourSoft = (
        $options->{gridlines} ? ( border => 7 ) : (),
        $backgroundColour
        ? ( bg_color => BGYELLOW, @defaultColour, )
        : (
            $borderColour ? ( border => 1, border_color => EXCELCOL11 ) : (),
            $textColour ? ( color => EXCELCOL11 ) : (),
        )
    );
    my @colourScribbles = (
        color => PURPLE,
        !$backgroundColour ? ()
        : $orangeColours ? ( bottom => 3, top => 3, border_color => PURPLE, )
        :                  ( bg_color => LTPURPLE )
    );
    my @colourHeader = (
        $goldColours
        ? ( bg_color => SILVER, fg_color => BGGOLD, pattern => 6 )
        : $backgroundColour
        ? ( bg_color => $orangeColours ? BGORANGE : BGPURPLE )
        : ()
    );
    my @colourUnavailable = (
        $options->{gridlines} ? ( border => 7 ) : (),
        $backgroundColour
        ? ( fg_color => SILVER, bg_color => WHITE, pattern => 14, )
        : ( right => 4, border_color => GREY )
    );
    my @colourUnused = (
        $options->{gridlines} ? ( border => 7 ) : (),
        $backgroundColour
        ? ( fg_color => SILVER, bg_color => WHITE, pattern => 15, )
        : ( right => 4, border_color => GREY )
    );
    my @colourCaption = $backgroundColour ? () : ( color => BLUE );
    my @colourTitle   = $backgroundColour ? () : ( color => ORANGE );
    my @sizeExtras =
      $goldColours ? ( top => 1, bottom => 1, border_color => BGGOLD ) : ();
    my @sizeCaption    = ( size   => 15, );
    my @sizeHeading    = ( valign => 'vbottom', size => 15, );
    my @sizeLabel      = ( valign => 'vcenter', @sizeExtras );
    my @sizeLabelGroup = ( valign => 'vcenter', );
    my @sizeNumber     = ( valign => 'vcenter', @sizeExtras );
    my @sizeText       = ( valign => 'vcenter', );
    my $plus  = '[Blue]_-+';
    my $minus = '[Red]_+-';
    my %specs = (
        '%con'    => [ locked => 1, @sizeNumber, @numPercent, @colourCon, ],
        '%copy'   => [ locked => 1, @sizeNumber, @numPercent, @colourCopy, ],
        '%hard'   => [ locked => 0, @sizeNumber, @numPercent, @colourHard, ],
        '%soft'   => [ locked => 1, @sizeNumber, @numPercent, @colourSoft, ],
        '%hardpm' => [
            locked => 0,
            @sizeNumber,
            num_format => "$plus????0.0%;$minus????0.0%;[Green]=;$cyan@",
            align      => 'center',
            @colourHard,
        ],
        '%softpm' => [
            locked => 1,
            @sizeNumber,
            num_format => "$plus????0.0%;$minus????0.0%;[Green]=;$cyan@",
            align      => 'center',
            @colourSoft,
        ],
        '%copypm' => [
            locked => 1,
            @sizeNumber,
            num_format => "$plus????0.0%;$minus????0.0%;[Green]=;$cyan@",
            align      => 'center',
            @colourCopy,
        ],
        '0.000con'    => [ locked => 1, @sizeNumber, @num_000,   @colourCon, ],
        '0.00con'     => [ locked => 1, @sizeNumber, @num_00,    @colourCon, ],
        '0.0con'      => [ locked => 1, @sizeNumber, @num_0,     @colourCon, ],
        '0.000copy'   => [ locked => 1, @sizeNumber, @num_000,   @colourCopy, ],
        '0.00copy'    => [ locked => 1, @sizeNumber, @num_00,    @colourCopy, ],
        '0.0copy'     => [ locked => 1, @sizeNumber, @num_0,     @colourCopy, ],
        '0.00000copy' => [ locked => 1, @sizeNumber, @num_00000, @colourCopy, ],
        '0.000hard'   => [ locked => 0, @sizeNumber, @num_000,   @colourHard, ],
        '0.00hard'    => [ locked => 0, @sizeNumber, @num_00,    @colourHard, ],
        '0.0hard'     => [ locked => 0, @sizeNumber, @num_0,     @colourHard, ],
        '0.000soft'   => [ locked => 1, @sizeNumber, @num_000,   @colourSoft, ],
        '0.00soft'    => [ locked => 1, @sizeNumber, @num_00,    @colourSoft, ],
        '0.0soft'     => [ locked => 1, @sizeNumber, @num_0,     @colourSoft, ],
        '0.00000soft' => [ locked => 1, @sizeNumber, @num_00000, @colourSoft, ],
        '0.000softpm' => [
            locked => 1,
            @sizeNumber,
            num_format => "$plus?0.000;$minus?0.000;[Green]=;$cyan@",
            align      => 'center',
            @colourSoft,
        ],
        '0.00softpm' => [
            locked => 1,
            @sizeNumber,
            num_format => "$plus??0.00;$minus??0.00;[Green]=;$cyan@",
            align      => 'center',
            @colourSoft,
        ],
        '0.0softpm' => [
            locked => 1,
            @sizeNumber,
            num_format => "$plus??0.0;$minus??0.0;[Green]=;$cyan@",
            align      => 'center',
            @colourSoft,
        ],
        'millioncopy' =>
          [ locked => 1, @sizeNumber, @num_million, @colourCopy, ],
        'millionhard' =>
          [ locked => 0, @sizeNumber, @num_million, @colourHard, ],
        'millionsoft' =>
          [ locked => 1, @sizeNumber, @num_million, @colourSoft, ],
        '0con'    => [ locked => 1, @sizeNumber, @num_, @colourCon, ],
        '0copy'   => [ locked => 1, @sizeNumber, @num_, @colourCopy, ],
        '0hard'   => [ locked => 0, @sizeNumber, @num_, @colourHard, ],
        '0soft'   => [ locked => 1, @sizeNumber, @num_, @colourSoft, ],
        '0softpm' => [
            locked => 1,
            @sizeNumber,
            $rightpad
            ? (
                num_format =>
                  "$plus#,##0$rightpad;$minus#,##0$rightpad;[Green]=;$cyan@",
                align => 'right'
              )
            : (
                num_format => "$plus?$q3??0;$minus?$q3??0;[Green]=;$cyan@",
                align      => 'center'
            ),
            @colourSoft,
        ],
        '0copypm' => [
            locked => 1,
            @sizeNumber,
            $rightpad
            ? (
                num_format =>
                  "$plus#,##0$rightpad;$minus#,##0$rightpad;[Green]=;$cyan@",
                align => 'right'
              )
            : (
                num_format => "$plus?$q3??0;$minus?$q3??0;[Green]=;$cyan@",
                align      => 'center'
            ),
            @colourCopy,
        ],
        '0000hard' => [
            locked => 0,
            @sizeNumber,
            align      => 'center',
            num_format => "${black}0000;-0000;${black}0000;$cyan@",
            @colourHard,
        ],
        '0000con' => [
            locked => 0,
            @sizeNumber,
            align      => 'center',
            num_format => "${black}0000;-0000;${black}0000;$cyan@",
            @colourCon,
        ],
        '0000copy' => [
            locked => 0,
            @sizeNumber,
            align      => 'center',
            num_format => "${black}0000;-0000;${black}0000;$cyan@",
            @colourCopy,
        ],
        boolhard => [
            locked => 0,
            @sizeNumber,
            align      => 'center',
            num_format => $numText,
            @colourHard,
        ],
        boolsoft => [
            locked => 1,
            @sizeNumber,
            align      => 'center',
            num_format => $numText,
            @colourSoft,
        ],
        boolcopy => [
            locked => 1,
            @sizeNumber,
            align      => 'center',
            num_format => $numText,
            @colourCopy,
        ],
        caption => [
            locked => 1,
            @sizeCaption,
            num_format => '@',
            align      => 'left',
            bold       => 1,
            @colourCaption,
        ],
        captionca => [
            locked => 1,
            @sizeCaption,
            num_format => '@',
            text_wrap  => 1,
            align      => 'center_across',
            bold       => 1,
            @colourCaption,
        ],
        hard => [
            locked => !$options->{validation}
              || $options->{validation} !~ /lenient/i,
            @sizeText,
            text_wrap => 0,
            @sizeExtras,
        ],
        link => [
            locked => 1,
            @sizeText,
            num_format => '@',
            underline  => 1,
            color      => BLUE,
        ],
        locsoft => [
            locked => 1,
            @sizeLabel,
            num_format => $black . '\L\o\c\a\t\i\o\n\ 0',
            align      => 'center',
            @sizeExtras,
            @colourSoft,
        ],
        notes => [
            locked => 1,
            @sizeHeading,
            num_format => '@',
            align      => 'left',
            bold       => 1,
            @colourTitle,
        ],
        scribbles => [
            locked => 0,
            @sizeText,
            text_wrap => 0,
            @colourScribbles,
        ],
        codecopy => [
            locked => 1,
            @sizeText,
            num_format => '@',
            align      => 'left',
            font       => 'Courier New',
            @colourSoft,
        ],
        puretextcentered => [
            locked => 0,
            @sizeText,
            num_format => $numTextOnly,
            align      => 'center',
            @sizeExtras,
            text_wrap => 1,
            @colourHard,
        ],
        text => [
            locked => 1,
            @sizeText,
            num_format => '@',
        ],
        textcon => [
            locked => 1,
            @sizeText,
            num_format => $numText,
            @alignText,
            text_wrap => 1,
            @sizeExtras,
            @colourCon,
        ],
        textcopy => [
            locked => 1,
            @sizeText,
            num_format => $numText,
            @alignText, @sizeExtras,
            text_wrap => 1,
            @colourCopy,
        ],
        textcopycentered => [
            locked => 1,
            @sizeText,
            num_format => $numText,
            align      => 'center',
            @sizeExtras,
            text_wrap => 1,
            @colourCopy,
        ],
        texthard => [
            locked => 0,
            @sizeText,
            num_format => $numText,
            @alignText, @sizeExtras,
            text_wrap => 1,
            @colourHard,
        ],
        textsoft => [
            locked => 1,
            @sizeText,
            num_format => $numText,
            @alignText, @sizeExtras,
            text_wrap => 1,
            @colourSoft,
        ],
        textnocolour => [
            locked => 0,
            @sizeText,
            num_format => $numText,
            @alignText, @sizeExtras,
            text_wrap => 1,
            @defaultColour,
        ],
        th => [
            locked => 1,
            @sizeLabel,
            num_format => '@',
            align      => 'left',
            bold       => 1,
            text_wrap  => 1,
            @colourHeader,
            $options->{gridlines} ? ( bottom => 7, right => 1 ) : (),
        ],
        thc => [
            locked => 1,
            @sizeLabel,
            num_format => '@',
            bold       => 1,
            text_wrap  => 1,
            align      => 'center',
            @colourHeader,
            $options->{gridlines} ? ( right => 7, bottom => 1 ) : (),
        ],
        thca => [
            locked => 1,
            @sizeLabelGroup,
            num_format => '@',
            italic     => 1,
            text_wrap  => 1,
            align      => 'center_across',
            @colourHeader,
            $options->{gridlines} ? ( right => 7, bottom => 1, )
            : ( left => 1, right => 1, ),
        ],
        thg => [
            locked => 1,
            @sizeLabelGroup,
            num_format => '@',
            align      => 'left',
            italic     => 1,
            text_wrap  => 1,
            @colourHeader,
        ],
        thla => [
            locked => 1,
            @sizeLabelGroup,
            num_format => '@',
            align      => 'left',
            italic     => 1,
            @colourHeader,
        ],
        colnoteleft => [
            locked => 1,
            @sizeText,
            num_format => '@',
            align      => 'left',
        ],
        colnoteleftwrap => [
            locked => 1,
            @sizeText,
            num_format => '@',
            text_wrap  => 1,
            align      => 'left',
        ],
        colnotecenter => [
            locked => 1,
            @sizeText,
            num_format => '@',
            text_wrap  => 1,
            align      => 'center_across',
            left       => 1,
            right      => 1,
        ],
        thloc => [
            locked => 1,
            @sizeLabel,
            num_format => '\L\o\c\a\t\i\o\n\ 0',
            align      => 'left',
            bold       => 1,
            @colourHeader,
            $options->{gridlines} ? ( bottom => 7, right => 1 ) : (),
        ],
        thtar => [
            locked => 1,
            @sizeLabel,
            num_format => '\T\a\r\i\f\f\ 0',
            align      => 'left',
            bold       => 1,
            @colourHeader,
            $options->{gridlines} ? ( bottom => 7, right => 1 ) : (),
        ],
        thtarimport => [
            locked => 0,
            @sizeLabel,
            num_format => '\T\a\r\i\f\f\ 0 \i\m\p\o\r\t',
            align      => 'left',
            bold       => 1,
            @colourHeader,
            $options->{gridlines} ? ( bottom => 7, right => 1 ) : (),
        ],
        thtarexport => [
            locked => 0,
            @sizeLabel,
            num_format => '\T\a\r\i\f\f\ 0 \e\x\p\o\r\t',
            align      => 'left',
            bold       => 1,
            @colourHeader,
            $options->{gridlines} ? ( bottom => 7, right => 1 ) : (),
        ],
        unavailable => [
            locked => 1,
            @sizeNumber,
            num_format => $numText,
            @colourUnavailable,
        ],
        unused => [
            locked => !$options->{validation}
              || $options->{validation} !~ /lenient/i,
            @sizeNumber,
            num_format => $numText,
            @colourUnused,
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
        tlttr => [
            right       => 5,
            right_color => 8,
        ],
        blue => [
            color => $orangeColours ? BGORANGE : BGPURPLE,
            size => 13,
            bg_color => BLUE,
        ],
        red => [
            color => $orangeColours ? BGORANGE : BGPURPLE,
            size => 13,
            bg_color => RED,
        ],
        algae => [
            color => $orangeColours ? BGORANGE : BGPURPLE,
            size => 13,
            bg_color => EXCELCOL13,
        ],
        purple => [
            color => $orangeColours ? BGORANGE : BGPURPLE,
            size => 13,
            bg_color => PURPLE,
        ],
        slime => [
            color => $orangeColours ? BGORANGE : BGPURPLE,
            size => 13,
            bg_color => EXCELCOL11,
        ],
    };

}

1;
