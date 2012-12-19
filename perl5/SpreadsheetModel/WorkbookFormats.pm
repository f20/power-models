package SpreadsheetModel::WorkbookFormats;

=head Copyright licence and disclaimer

Copyright 2008-2012 Reckon LLP and others. All rights reserved.

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

=head Development notes

?,??0 style formats do not work with OpenOffice.org.  Should have a right-align option (with a few _m of padding on the right) to deal with that.

SHould probably adding formats like "£"???0.0,,"m" and "£"0.0,"k" (for financial modelling).

=cut

sub getFormat {
    my ( $workbook, $key ) = @_;
    $workbook->{formats}{$key} ||= $workbook->add_format(
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
        }
    );
}

use constant {
    BLACK    => 8,     # black #000000
    WHITE    => 9,     # white #ffffff
    RED      => 10,    # red #ff0000
    LIME     => 11,    # lime
    BLUE     => 12,    # blue #0000ff potentially overridden by #0066cc
    YELLOW   => 13,    # yellow #ffff00
    MAGENTA  => 14,    # magenta #ff00ff
    CYAN     => 15,    # cyan
    BROWN    => 16,    # brown
    GREEN    => 17,    # green #008000
    NAVY     => 18,    # navy
    DKYELLOW => 19,    # #808000
    PURPLE   => 20,    # purple #800080
    SILVER   => 22,    # silver #c0c0c0 potentially overridden by #e9e9e9
    MAGENTA2 => 33,    # defaults to magenta (?)
    BLUE2    => 39,    # defaults to blue or some kind of green
    BGBLUE   => 41,    # #ccffff
    BGGREEN  => 42,    # #ccffcc
    BGYELLOW => 43,    # #ffff99 potentially overridden by #ffffcc
    BGPINK   => 45,    # #ff99cc potentially overridden by #ffccff
    SILVER2  => 47,    # defaults to silver or some kind of salmon
    BGDKYELL => 51,    # #ffcc00
    BGORANGE => 52,    # #ff9900 potentially overridden by #ffcc99
    ORANGE   => 53,    # orange #ff6600 potentially overridden by #ff6633
    GREY     => 55,    # #969696 potentially overridden by #999999
}; # these codes (between 8 and 63) are equal to the colour number used in VBA (1-56) plus 7

sub setFormats {
    my ( $workbook, $options ) = @_;
    unless ( $options->{defaultColours} ) {
        $workbook->set_custom_color( BLUE,     '#0066cc' );
        $workbook->set_custom_color( BGYELLOW, '#ffffcc' );
        $workbook->set_custom_color( BGPINK,   '#ffccff' );
        $workbook->set_custom_color( SILVER,   '#e9e9e9' );
        $workbook->set_custom_color( BGORANGE, '#ffcc99' );
        $workbook->set_custom_color( ORANGE,   '#ff6633' );
        $workbook->set_custom_color( GREY,     '#999999' );
    }
    my $q3 = $options->{alignment} ? ',' : '??,???,';
    my $rightpad;
    $rightpad = '_m' x ( $1 || 2 )
      if $options->{alignment} && $options->{alignment} =~ /right.*([0-9]+)?/;
    my @numPercent =
      $rightpad
      ? (
        num_format => "0.0%_)$rightpad;[Red](??0.0%)$rightpad;;@",
        align      => 'right'
      )
      : ( num_format => ' _(??0.0%_);[Red] (??0.0%);;@', align => 'center' );
    my @num_ =
      $rightpad
      ? (
        num_format => "#,##0_)$rightpad;[Red](#,##0)$rightpad;;@",
        align      => 'right'
      )
      : (
        num_format => ' _(?' . $q3 . '??0_);[Red] (?' . $q3 . '??0);;@',
        align      => 'center'
      );
    my @num_0 =
      $rightpad
      ? (
        num_format => "#,##0.0_)$rightpad;[Red](#,##0.0)$rightpad;;@",
        align      => 'right'
      )
      : (
        num_format => ' _(?' . $q3 . '??0.0_);[Red] (?' . $q3 . '??0.0);;@',
        align      => 'center'
      );
    my @num_00 =
      $rightpad
      ? (
        num_format => "#,##0.00_)$rightpad;[Red](#,##0.00)$rightpad;;@",
        align      => 'right'
      )
      : (
        num_format => ' _(?' . $q3 . '??0.00_);[Red] (?' . $q3 . '??0.00);;@',
        align      => 'center'
      );
    my @num_000 =
      $rightpad
      ? (
        num_format => "#,##0.000_)$rightpad;[Red](#,##0.000)$rightpad;;@",
        align      => 'right'
      )
      : (
        num_format => ' _(?' . $q3 . '??0.000_);[Red] (?' . $q3 . '??0.000);;@',
        align      => 'center'
      );
    my @num_00000 =
      $rightpad
      ? (
        num_format => "#,##0.00000_)$rightpad;[Red](#,##0.00000)$rightpad;;@",
        align      => 'right'
      )
      : (
        num_format => ' _(?'
          . $q3
          . '??0.00000_);[Red] (?'
          . $q3
          . '??0.00000);;@',
        align => 'center'
      );
    my @colourCon =
       !$options->{colour} ? ( bg_color => SILVER )
      : $options->{colour} !~ /t/i ? ( border => 1, border_color => GREY )
      :                              ( color => GREY );
    my @colourCopy =
       !$options->{colour} ? ( bg_color => BGGREEN )
      : $options->{colour} !~ /t/i ? ( border => 1, border_color => GREEN )
      :                              ( color => GREEN );
    my @colourHard =
       !$options->{colour} ? ( bg_color => BGBLUE )
      : $options->{colour} !~ /t/i ? ( border => 1, border_color => BLUE )
      :                              ( color => BLUE );
    my @colourSoft =
       !$options->{colour} ? ( bg_color => BGYELLOW )
      : $options->{colour} !~ /t/i ? ( border => 1, border_color => DKYELLOW )
      :                              ( color => DKYELLOW );
    my @colourScribbles =
      !$options->{colour}
      ? ( color => PURPLE, bottom => 3, top => 3, border_color => PURPLE, )
      : ( color => PURPLE );
    my @colourHeader = !$options->{colour} ? ( bg_color => BGORANGE ) : ();
    my @colourUnavailable =
      !$options->{colour}
      ? ( fg_color => SILVER, bg_color => WHITE, pattern => 14, )
      : ( right => 4, border_color => GREY );
    my @colourUnused =
      !$options->{colour}
      ? ( fg_color => SILVER, bg_color => WHITE, pattern => 15, )
      : ( right => 4, border_color => GREY );
    my @colourCaption = !$options->{colour} ? () : ( color => BLUE );
    my @colourTitle   = !$options->{colour} ? () : ( color => ORANGE );
    my @sizeCaption    = ( size   => 15 );
    my @sizeHeading    = ( valign => 'vbottom', size => 15 );
    my @sizeLabel      = ( valign => 'vcenter', );
    my @sizeLabelGroup = ( valign => 'vcenter', );
    my @sizeNumber     = ( valign => 'vcenter', );
    my @sizeText       = ( valign => 'vcenter', );
    $workbook->{formatspec} = {
        '%con'    => [ locked => 1, @sizeNumber, @numPercent, @colourCon, ],
        '%copy'   => [ locked => 1, @sizeNumber, @numPercent, @colourCopy, ],
        '%hard'   => [ locked => 0, @sizeNumber, @numPercent, @colourHard, ],
        '%soft'   => [ locked => 1, @sizeNumber, @numPercent, @colourSoft, ],
        '%softpm' => [
            locked => 1,
            @sizeNumber,
            num_format => '[Blue]+??0.0%;[Red]-??0.0%;[Green]=',
            align      => 'center',
            @colourSoft,
        ],
        '%copypm' => [
            locked => 1,
            @sizeNumber,
            num_format => '[Blue]+??0.0%;[Red]-??0.0%;[Green]=',
            align      => 'center',
            @colourCopy,
        ],
        '0.000con'    => [ locked => 1, @sizeNumber, @num_000,   @colourCon, ],
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
        '0.00softpm'  => [
            locked => 1,
            @sizeNumber,
            num_format => '[Blue]+??0.00;[Red]-??0.00;;@',
            align      => 'center',
            @colourSoft,
        ],
        '0.000softpm' => [
            locked => 1,
            @sizeNumber,
            num_format => '[Blue]+?0.000;[Red]-?0.000;;@',
            align      => 'center',
            @colourSoft,
        ],
        '0con'  => [ locked => 1, @sizeNumber, @num_, @colourCon, ],
        '0copy' => [ locked => 1, @sizeNumber, @num_, @colourCopy, ],
        '0hard' => [ locked => 0, @sizeNumber, @num_, @colourHard, ],
        '0000hard' =>
          [ locked => 0, @sizeNumber, num_format => '0000', @colourHard, ],
        '0000copy' =>
          [ locked => 0, @sizeNumber, num_format => '0000', @colourCopy, ],
        '0soft'   => [ locked => 1, @sizeNumber, @num_, @colourSoft, ],
        '0softpm' => [
            locked => 1,
            @sizeNumber,
            $rightpad
            ? (
                num_format => "[Blue]+#,##0$rightpad;[Red]-#,##0$rightpad;;@",
                align      => 'right'
              )
            : (
                num_format => '[Blue]+?' . $q3 . '??0;[Red]-?' . $q3 . '??0;;@',
                align      => 'center'
            ),
            @colourSoft,
        ],
        '0copypm' => [
            locked => 1,
            @sizeNumber,
            $rightpad
            ? (
                num_format => "[Blue]+#,##0$rightpad;[Red]-#,##0$rightpad;;@",
                align      => 'right'
              )
            : (
                num_format => '[Blue]+?' . $q3 . '??0;[Red]-?' . $q3 . '??0;;@',
                align      => 'center'
            ),
            @colourCopy,
        ],
        boolhard => [
            locked => 0,
            @sizeNumber,
            align => 'center',
            @colourHard,
        ],
        boolsoft => [
            locked => 1,
            @sizeNumber,
            align => 'center',
            @colourSoft,
        ],
        caption => [
            locked => 1,
            @sizeCaption,
            num_format => '@',
            align      => 'left',
            bold       => 1,
            @colourCaption,
        ],
        hard => [
            locked => !$options->{validation}
              || $options->{validation} !~ /lenient/i,
            @sizeText, text_wrap => 0,
        ],
        link => [
            locked => 1,
            @sizeText,
            num_format => '@',
            align      => 'left',
            underline  => 1,
            color      => BLUE,
            1 ? () : ( bg_color => WHITE ),
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
        text => [
            locked => 1,
            @sizeText,
            num_format => '@',
            align      => 'left',
        ],
        textcell => [
            locked => 1,
            @sizeText,
            num_format => '@',
            align      => 'left',
            text_wrap  => 1,
        ],
        textcon => [
            locked => 1,
            @sizeText,
            num_format => '@',
            align      => 'left',
            text_wrap  => 1,
            @colourCon,
        ],
        textcopy => [
            locked => 1,
            @sizeText,
            num_format => '@',
            align      => 'left',
            text_wrap  => 1,
            @colourCopy,
        ],
        texthard => [
            locked => 0,
            @sizeText,
            num_format => '@',
            align      => 'left',
            text_wrap  => 1,
            @colourHard,
        ],
        textlrap => [
            locked => 1,
            @sizeText,
            num_format => '@',
            align      => 'left',
            1 ? () : ( bg_color => WHITE ),
        ],
        textwrap => [
            locked => 1,
            @sizeText,
            num_format => '@',
            text_wrap  => 1,
            align      => 'center_across',
            1 ? () : ( bg_color => WHITE ),
            left  => 1,
            right => 1,
        ],
        th => [
            locked => 1,
            @sizeLabel,
            num_format => '@',
            align      => 'left',
            bold       => 1,
            text_wrap  => 1,
            @colourHeader,
        ],
        thc => [
            locked => 1,
            @sizeLabel,
            num_format => '@',
            bold       => 1,
            text_wrap  => 1,
            align      => 'center',
            @colourHeader,
        ],
        thca => [
            locked => 1,
            @sizeLabelGroup,
            num_format => '@',
            italic     => 1,
            text_wrap  => 1,
            align      => 'center_across',
            @colourHeader,
            left  => 1,
            right => 1,
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
        thloc => [
            locked => 1,
            @sizeLabel,
            num_format => q"\L\o\c\a\t\i\o\n\ 0",
            align      => 'left',
            bold       => 1,
            @colourHeader,
        ],
        thtar => [
            locked => 1,
            @sizeLabel,
            num_format => q"\T\a\r\i\f\f\ 0",
            align      => 'left',
            bold       => 1,
            @colourHeader,
        ],
        unavailable => [
            locked => 1,
            @sizeNumber,
            num_format => '0.000;-0.000;;@',
            align      => 'center',
            @colourUnavailable,
        ],
        unused => [
            locked => !$options->{validation}
              || $options->{validation} !~ /lenient/i,
            @sizeNumber,
            num_format => '0.000;-0.000;;@',
            align      => 'center',
            @colourUnused,
        ],
    };
}

1;
