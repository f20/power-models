package EDCM2;

=head Copyright licence and disclaimer

Copyright 2013 Franck Latrémolière, Reckon LLP and others.

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
require Spreadsheet::WriteExcel::Utility;

sub templates {
    my (
        $model,                            $tariffs,
        $importCapacity,                   $exportCapacityExempt,
        $exportCapacityChargeablePre2005,  $exportCapacityChargeable20052010,
        $exportCapacityChargeablePost2010, $tariffSoleUseMeav,
        $tariffLoc,                        $tariffCategory,
        $useProportions,                   $activeCoincidence,
        $reactiveCoincidence,              $indirectExposure,
        $nonChargeableCapacity,            $activeUnits,
        $creditableCapacity,               $tariffNetworkSupportFactor,
        $tariffDaysInYearNot,              $tariffHoursInRedNot,
        $previousChargeImport,             $previousChargeExport,
        $llfcImport,                       $llfcExport,
        $thisIsTheTariffTable,             $daysInYear,
        $hoursInRed,
    ) = @_;

    push @{ $model->{tablesTemplateImport} },
      $model->templateImport(
        $tariffs,        $llfcImport,          $thisIsTheTariffTable,
        $importCapacity, $activeCoincidence,   $daysInYear,
        $hoursInRed,     $tariffDaysInYearNot, $tariffHoursInRedNot,
      );

    push @{ $model->{tablesTemplateExport} },
      $model->templateExport( $tariffs, $llfcImport, $thisIsTheTariffTable );
}

sub templateImport {

    my (
        $model,                $tariffs,        $llfcImport,
        $thisIsTheTariffTable, $importCapacity, $activeCoincidence,
        $daysInYear,           $hoursInRed,     $tariffDaysInYearNot,
        $tariffHoursInRedNot,
    ) = @_;

    $model->{importTariffIndex} = my $index = Dataset(
        name          => 'Number',
        data          => [ [1] ],
        defaultFormat => 'thtarimport',
    );

    my @tariffComponents = map {
        Arithmetic(
            name          => $_->{name}->shortName,
            arguments     => { IV1 => $index, IV2_IV3 => $_ },
            arithmetic    => '=INDEX(IV2_IV3,IV1)',
            defaultFormat => $_->{name} =~ /k(?:VAr|W)h/
            ? '0.000copy'
            : '0.00copy',
        );
    } @{ $thisIsTheTariffTable->{columns} }[ 1 .. 4 ];

    my $agreedCapacity = Arithmetic(
        name          => 'Maximum import capacity (kVA)',
        arguments     => { IV1 => $index, IV2_IV3 => $importCapacity },
        arithmetic    => '=INDEX(IV2_IV3,IV1)',
        defaultFormat => '0hard',
    );

    my $exceededCapacity = Dataset(
        name          => 'Exceeded import capacity (kVA)',
        data          => [ [0] ],
        defaultFormat => '0hard',
    );

    foreach ( $activeCoincidence, $tariffDaysInYearNot, $tariffHoursInRedNot, )
    {
        my $df = $_->{defaultFormat} || '0.000soft';
        $df =~ s/copy|soft/hard/;
        $_ = Arithmetic(
            name          => $_->{name}->shortName,
            arguments     => { IV1 => $index, IV2_IV3 => $_ },
            arithmetic    => '=INDEX(IV2_IV3,IV1)',
            defaultFormat => $df,
        );
    }

    $_ = Stack( sources => [$_] ) foreach $daysInYear, $hoursInRed;

    my $units = Arithmetic(
        name          => 'Units consumed in super-red time band (kWh)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*IV2*(IV3-IV4)',
        arguments     => {
            IV1 => $agreedCapacity,
            IV2 => $activeCoincidence,
            IV3 => $hoursInRed,
            IV4 => $tariffHoursInRedNot,
        },
    );

    my $redPounds = Arithmetic(
        name          => 'Annual super-red charge (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*IV2/100',
        arguments     => { IV1 => $units, IV2 => $tariffComponents[0], },
    );

    my $fixedPounds = Arithmetic(
        name          => 'Annual fixed charge (£)',
        defaultFormat => '0soft',
        arithmetic    => '=(IV1-IV3)*IV2/100',
        arguments     => {
            IV1 => $daysInYear,
            IV2 => $tariffComponents[1],
            IV3 => $tariffDaysInYearNot,
        },
    );

    my $capacityPounds = Arithmetic(
        name          => 'Annual capacity charge (£)',
        defaultFormat => '0soft',
        arithmetic    => '=(IV1-IV3)*(IV5*IV6+IV7*IV8)/100',
        arguments     => {
            IV1 => $daysInYear,
            IV3 => $tariffDaysInYearNot,
            IV5 => $tariffComponents[2],
            IV7 => $tariffComponents[3],
            IV6 => $agreedCapacity,
            IV8 => $exceededCapacity,
        },
    );

    my $totalPounds = Arithmetic(
        name          => 'Total annual DUoS charge (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1+IV2+IV3',
        arguments =>
          { IV1 => $redPounds, IV2 => $fixedPounds, IV3 => $capacityPounds, },
    );

    my @psv;
    my $col = 1;
    foreach my $d ( @{ $model->{table935}{columns} } ) {
        if ( my $nc = $d->lastCol ) {
            push @psv, map {
                Arithmetic(
                    defaultFormat => 'codecopy',
                    name          => '',
                    arithmetic    => '="935|'
                      . ( $col + $_ )
                      . '|"&IV1&"|"&INDEX(IV3_IV4,IV2,'
                      . ( 1 + $_ ) . ')',
                    arguments => {
                        IV1     => $index,
                        IV2     => $index,
                        IV3_IV4 => $d
                    },
                  )
            } 0 .. $nc;
            $col += $nc + 1;
        }
        else {
            push @psv,
              Arithmetic(
                defaultFormat => 'codecopy',
                name          => '',
                arithmetic => '="935|' . $col . '|"&IV1&"|"&INDEX(IV3_IV4,IV2)',
                arguments  => {
                    IV1     => $index,
                    IV2     => $index,
                    IV3_IV4 => $d
                },
              ) unless $d->{name} =~ /export/i;
            ++$col;
        }
    }

    $_ = Arithmetic(
        name          => $_->{name}->shortName,
        arguments     => { IV1 => $index, IV2_IV3 => $_ },
        arithmetic    => '=INDEX(IV2_IV3,IV1)',
        defaultFormat => 'textcopy',
    ) foreach $tariffs, $llfcImport;

    Notes(
        name  => 'Import template',
        lines => <<'EOX'

ELECTRICITY DISTRIBUTION CHARGES INFORMATION FOR IMPORT

This template is intended to illustrate the use of system charges that a distributor might levy on a
supplier under an EHV Distribution Charging Methodology (EDCM).

Charges between supplier and end customer are a bilateral contractual matter.  A supplier may apply
its own charges in addition to, or instead of, the charges that this template illustrates.

This template is for illustration only.  In case of conflict, the published statement of Distribution Use of System charges takes precedence. 

EOX
      ),

      map {
        my $t = ref $_ eq 'ARRAY' ? $_->[1] : $_;
        ref $t
          ? Columnset(
            noHeaders     => 1,
            noSpacing     => ref $_ ne 'ARRAY',
            name          => ref $_ eq 'ARRAY' ? $_->[0] : '',
            singleRowName => ref $t->{name}
            ? $t->{name}->shortName
            : $t->{name},
            columns => [$t],
            ref $_ eq 'ARRAY' && $_->[2] ? ( lines => $_->[2] ) : (),
          )
          : Notes( name => '', lines => $t );
      }

      [ 'Tariff identification', $index, ], $tariffs, $llfcImport,

      [
        'Distribution Use of System (DUoS) tariff (excluding VAT)',
        $tariffComponents[0]
      ],
      @tariffComponents[ 1 .. $#tariffComponents ],

      [ 'Calendar and time band information', $daysInYear, ],
      $tariffDaysInYearNot,
      $hoursInRed, $tariffHoursInRedNot,

      [ 'Capacity and consumption', $agreedCapacity ],
      $exceededCapacity, $activeCoincidence, $units,

      [
        'Distribution Use of System (DUoS) charges (excluding VAT)', $redPounds
      ],
      $fixedPounds, $capacityPounds, $totalPounds,

      <<EOL,
From the National Terms of Connection

12.6 Except where a variation requires a Modification, either party may propose a variation to the Maximum Import Capacity and/or Maximum Export
Capacity by notice in writing to the other Party. The Company and the Customer shall negotiate in good faith such a variation, but where it is
not agreed section 23 of the Act may entitle the Customer to refer the matter to the Authority.

12.7 Any reduction in the Maximum Import Capacity or the Maximum Export Capacity pursuant to Clause 12.6 shall, where the Parties have within
the preceding 12 months agreed the Maximum Import Capacity or the Maximum Export Capacity (as applicable), only take effect following the expiry
of 12 months from the date of such previous agreement (unless the Company expressly agrees otherwise).
EOL

      <<EOL,
From the CDCM

149. The level of MIC will be agreed at the time of connection and when an increase has been approved. Following such an agreement (be it at the
time of connection or an increase) no reduction in MIC will be allowed for a period of one year.

150. Reductions to the MIC may only be permitted once in a 12 month period and no retrospective changes will be allowed. Where MIC is reduced
the new lower level will be agreed with reference to the level of the customers’ maximum demand. It should be noted that where a new lower level
is agreed the original capacity may not be available in the future without the need for network reinforcement and associated cost.
EOL

      [
        'Disclosure of detailed data for advanced modelling',
        $psv[0],
        [
            'The following advanced technical information '
              . 'is not necessary to understand '
              . 'your charges but might be useful '
              . 'if you wish to conduct additional analysis.',
            'For further information about '
              . 'advanced modelling options, see:',
            'http://dcmf.co.uk/models/edcm.html'
        ]
      ],
      @psv[ 1 .. $#psv ];

}

sub templateExport {

    my ( $model, $tariffs, $llfcImport, $thisIsTheTariffTable, ) = @_;

    $model->{exportTariffIndex} = my $index = Dataset(
        name          => 'Number',
        data          => [ [1] ],
        defaultFormat => 'thtarexport',
    );

    my @tariffComponents = map {
        Arithmetic(
            name          => $_->{name}->shortName,
            arguments     => { IV1 => $index, IV2_IV3 => $_ },
            arithmetic    => '=INDEX(IV2_IV3,IV1)',
            defaultFormat => $_->{name} =~ /k(?:VAr|W)h/
            ? '0.000copy'
            : '0.00copy',
        );
    } @{ $thisIsTheTariffTable->{columns} }[ 5 .. 8 ];

    my @psv;
    my $col = 1;
    foreach my $d ( @{ $model->{table935}{columns} } ) {
        if ( my $nc = $d->lastCol ) {
            $col += $nc + 1;
        }
        else {
            push @psv,
              Arithmetic(
                defaultFormat => 'codecopy',
                name          => '',
                arithmetic => '="935|' . $col . '|"&IV1&"|"&INDEX(IV3_IV4,IV2)',
                arguments  => {
                    IV1     => $index,
                    IV2     => $index,
                    IV3_IV4 => $d
                },
              ) if $d->{name} =~ /export/i;
            ++$col;
        }
    }

    $_ = Arithmetic(
        name          => $_->{name}->shortName,
        arguments     => { IV1 => $index, IV2_IV3 => $_ },
        arithmetic    => '=INDEX(IV2_IV3,IV1)',
        defaultFormat => 'textcopy',
    ) foreach $tariffs, $llfcImport;

    Notes(
        name  => 'Export template',
        lines => <<'EOX'),

ELECTRICITY DISTRIBUTION CHARGES INFORMATION FOR EXPORT

This template is intended to illustrate the use of system charges that a distributor might levy on a
generator or supplier under an EHV Distribution Charging Methodology (EDCM).

Any charges between generator, supplier, customer and site owner are contractual matters.  They may
or may not reflect the charges that this template illustrates.

This template is for illustration only.  In case of conflict, the published statement of Distribution Use of System charges takes precedence.

This template is not complete.

EOX

      map {
        my $t = ref $_ eq 'ARRAY' ? $_->[1] : $_;
        ref $t
          ? Columnset(
            noHeaders     => 1,
            noSpacing     => ref $_ ne 'ARRAY',
            name          => ref $_ eq 'ARRAY' ? $_->[0] : '',
            singleRowName => ref $t->{name}
            ? $t->{name}->shortName
            : $t->{name},
            columns => [$t],
            ref $_ eq 'ARRAY' && $_->[2] ? ( lines => $_->[2] ) : (),
          )
          : Notes( name => '', lines => $t );
      }

      [ 'Tariff identification', $index, ], $tariffs, $llfcImport,

      [
        'Distribution Use of System (DUoS) tariff (excluding VAT)',
        $tariffComponents[0]
      ],
      @tariffComponents[ 1 .. $#tariffComponents ],

      <<EOL,
From the National Terms of Connection

12.6 Except where a variation requires a Modification, either party may propose a variation to the Maximum Import Capacity and/or Maximum Export
Capacity by notice in writing to the other Party. The Company and the Customer shall negotiate in good faith such a variation, but where it is
not agreed section 23 of the Act may entitle the Customer to refer the matter to the Authority.

12.7 Any reduction in the Maximum Import Capacity or the Maximum Export Capacity pursuant to Clause 12.6 shall, where the Parties have within
the preceding 12 months agreed the Maximum Import Capacity or the Maximum Export Capacity (as applicable), only take effect following the expiry
of 12 months from the date of such previous agreement (unless the Company expressly agrees otherwise).
EOL

      [
        'Disclosure of detailed data for advanced modelling',
        $psv[0],
        [
            'The following advanced technical information '
              . 'is not necessary to understand '
              . 'your charges but might be useful '
              . 'if you wish to conduct additional analysis.',
            'For further information about '
              . 'advanced modelling options, see:',
            'http://dcmf.co.uk/models/edcm.html'
        ]
      ],
      @psv[ 1 .. $#psv ];

}

sub vbaWrite {

    my ( $model, $wb, $ws ) = @_;

    my $sheetList = join ',', map { qq%"$_"% } 11,
      $model->{method} =~ /FCP/i ? 911 : $model->{method} =~ /LRIC/i ? 913 : (),
      935, qw(Results OneLiners);

    ( undef, my $importTariffRow, undef ) =
      $model->{importTariffIndex}->wsWrite( $wb, $ws );
    ( undef, my $exportTariffRow, undef ) =
      $model->{exportTariffIndex}->wsWrite( $wb, $ws );
    ++$_ foreach $importTariffRow, $exportTariffRow;

    Notes( name => '', lines => <<EOX )->wsWrite( $wb, $ws );

Sub Autorun()
    Worksheets("Index").Activate
    l = 0.5*(ActiveSheet.Columns(1).Left+ActiveSheet.Columns(2).Left)
    w = ActiveSheet.Columns(3).Left - ActiveSheet.Columns(2).Left
    Dim myButton As Shape
    Set myButton = ActiveSheet.Shapes.AddFormControl(xlButtonControl, _
    l, ActiveSheet.Rows(4).Top, w, ActiveSheet.Rows(7).Top - ActiveSheet.Rows(4).Top)
    myButton.TextFrame.Characters.Text = "Import data from master EDCM model"
    With myButton.TextFrame.Characters.Font ' Not reliable
        .FontStyle = "Bold"
        .Size = 15
    End With
    myButton.OnAction = "ImportData"
    Set myButton = ActiveSheet.Shapes.AddFormControl(xlButtonControl, _
    l, ActiveSheet.Rows(9).Top, w, ActiveSheet.Rows(12).Top - ActiveSheet.Rows(9).Top)
    With myButton
        .TextFrame.Characters.Text = "Export information from this workbook"
        .OnAction = "ExportAll"
        With .TextFrame.Characters.Font ' Not reliable
            .FontStyle = "Bold"
            .Size = 15
        End With
    End With
    ActiveSheet.Shapes.SelectAll
    ' Failed to automate increasing the font size of these buttons because of apparent Excel bugs
End Sub

Sub ExportAll()
    Dim model, core As String
    model = ActiveWorkbook.FullName
    core = Left(model, Len(model) - 5)
    Sheets(Array($sheetList)).Copy
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Select
        ws.Unprotect
        ws.UsedRange.Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
          False, Transpose:=False
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        Range("A1").Select
    Next ws
    Sheets("935").Select
    ActiveSheet.Unprotect
    ntariffsplus6 = ActiveSheet.UsedRange.Rows.Count
    Range("C7:C" & ntariffsplus6).Select
    Selection.FormulaR1C1 = "REDACTED"
    Selection.Copy
    Range("D7").Select
    ActiveSheet.Paste
    Range("E7").Select
    ActiveSheet.Paste
    Range("F7").Select
    ActiveSheet.Paste
    Range("G7").Select
    ActiveSheet.Paste
    Range("P7").Select
    ActiveSheet.Paste
    Range("Q7").Select
    ActiveSheet.Paste
    Range("S7").Select
    ActiveSheet.Paste
    Range("T7").Select
    ActiveSheet.Paste
    Range("U7").Select
    ActiveSheet.Paste
    Range("Y7").Select
    ActiveSheet.Paste
    Range("Z7").Select
    ActiveSheet.Paste
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("A1").Select
    Sheets("11").Select
    ActiveWorkbook.SaveAs Filename:=core & " - redacted for publication.xlsx", FileFormat:= _
      xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close

    For t = 1 To ntariffsplus6 - 6
        If Sheets("935").Cells(6 + t, 3) = "VOID" OR Sheets("935").Cells(6 + t, 3) = "" Then GoTo GEN_TARIFF
        Sheets("ImpT").Select
        Cells($importTariffRow, 2).Formula = t
        Sheets("ImpT").Copy
        Sheets("ImpT").Select
        ActiveSheet.Unprotect
        ActiveWorkbook.BreakLink Name:=model, Type:=xlExcelLinks
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        ActiveWorkbook.SaveAs Filename:=core & " - tariff " & t & " import.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWorkbook.Close
GEN_TARIFF:
        If Sheets("935").Cells(6 + t, 4) = "VOID" AND Sheets("935").Cells(6 + t, 5) = "VOID" AND Sheets("935").Cells(6 + t, 6) = "VOID" AND Sheets("935").Cells(6 + t, 7) = "VOID" Then GoTo NEXT_TARIFF
        Sheets("ExpT").Select
        Cells($exportTariffRow, 2).Formula = t
        Sheets("ExpT").Copy
        Sheets("ExpT").Select
        ActiveSheet.Unprotect
        ActiveWorkbook.BreakLink Name:=model, Type:=xlExcelLinks
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        ActiveWorkbook.SaveAs Filename:=core & " - tariff " & t & " export.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWorkbook.Close
NEXT_TARIFF:
    Next t

    Sheets("Index").Activate
    MsgBox "All done: files have been saved in the same folder as this model."

End Sub

Sub ImportData
    On Error GoTo RestoreSettings
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo 0

    Dim wsDatabase, iix As Worksheet
    On Error Resume Next
    Set iix = ActiveWorkbook.sheets("Index")
    Set wsDatabase = ActiveWorkbook.sheets("InputDB")
    On Error GoTo 0

    If wsDatabase Is Nothing Then GoTo RestoreSettings
    Dim x As String
    x = Application.GetOpenFilename()
    If x = "" Then GoTo RestoreSettings            
    Workbooks.Open x

    Dim dbLine As Long
    dbLine = 1
    wsDatabase.Rows("1:100").Insert Shift:=xlDown
    wsDatabase.Rows("1:100").Interior.ColorIndex = 1
    wsDatabase.Rows("1:100").Font.ColorIndex = 8

    wsDatabase.Cells(dbLine, 1).Value = "Table"
    wsDatabase.Cells(dbLine, 2).Value = "Column"
    wsDatabase.Cells(dbLine, 3).Value = "Row"
    wsDatabase.Cells(dbLine, 4).Value = "Key"
    wsDatabase.Cells(dbLine, 5).Value = "Value from " & ActiveWorkbook.Name

    Dim wsInput As Worksheet
    Dim numRows, numCols As Integer
    Dim cl As Range
    Dim table As Integer
    Dim twidth As Integer
    For Each wsInput In Worksheets
        If wsInput.Name <> wsDatabase.Name And Not (wsInput.Name Like "Calc*") And wsInput.UsedRange.row = 1 And wsInput.UsedRange.Column = 1 Then
            If modifyThisBook Then wsInput.Unprotect
            numRows = wsInput.UsedRange.Rows.Count
            numCols = wsInput.UsedRange.Columns.Count
            table = 0
            For r = 1 To wsInput.UsedRange.Rows.Count
                Dim row As String
                On Error Resume Next
                row = wsInput.Cells(r, 1).Value
                On Error GoTo 0
                If row Like "###.*" Then
                    table = Left(row, 3)
                    twidth = 0
                End If
                If row Like "####.*" Then
                    table = Left(row, 4)
                    twidth = 0
                End If
                If table = 0 Then GoTo DontDoThisRow
                If row Like "> *" Then
                    row = Right(row, Len(row) - 2)
                End If
                If wsInput.Cells(r - 1, 1).Value = "" And wsInput.Cells(r + 1, 1).Value = "" Then row = ""
                For c = 1 To wsInput.UsedRange.Columns.Count
                    Set cl = wsInput.Cells(r, c)
                    If c > twidth Then
                        If cl.Value <> "" Then twidth = c Else GoTo DontDoThisCell
                    End If
                    If cl.Locked Then GoTo DontDoThisCell
                    
                    ' It does not seem to be possible to test for the existence of the Validation property
                    ' Credits to Jamie Ham for workaround idea
                    Dim validationType As Long
                    validationType = xlValidateCustom ' something which is not xlValidateTextLength
                    On Error Resume Next
                    validationType = cl.Validation.Type
                    If validationType = xlValidateTextLength Then GoTo DontDoThisCell
                    ' if there is a validation by text length then the cell is not to be used
                    
                    On Error GoTo 0
                    dbLine = dbLine + 1
                    If dbLine Mod 100 = 0 Then wsDatabase.Rows(dbLine & ":" & (dbLine + 99)).Insert Shift:=xlDown
                    wsDatabase.Cells(dbLine, 1).Value = table
                    wsDatabase.Cells(dbLine, 2).Value = c - 1
                    wsDatabase.Cells(dbLine, 3).Value = row
                    wsDatabase.Cells(dbLine, 4).Formula = "=CONCATENATE(A" & dbLine & ",""|"",B" & dbLine & ",""|"",C" & dbLine & ")"
                    If internalFormulas Then
                        wsDatabase.Cells(dbLine, 5).FormulaR1C1 = "=INDEX('" & wsInput.Name & "'!R" & r & "C1:R" & r & "C" & twidth & "," & c & ")"
                    Else
                        With wsDatabase.Cells(dbLine, 5)
                            .Value = cl.Value
                            .WrapText = False
                        End With
                    End If
DontDoThisCell:
                On Error GoTo 0
                Next c
DontDoThisRow:
            Next r
            If modifyThisBook Then
                wsInput.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
                wsInput.EnableSelection = xlNoRestrictions
            End If
        End If
    Next wsInput

    dbLine = dbLine + 1
    wsDatabase.Cells(dbLine, 1).Value = 1190
    wsDatabase.Cells(dbLine, 2).Value = 1
    wsDatabase.Cells(dbLine, 3).Value = ""
    wsDatabase.Cells(dbLine, 4).Value = "1190|1|"
    wsDatabase.Cells(dbLine, 5).Value = TRUE

    If Not iix Is Nothing Then iix.Activate

RestoreSettings:
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
End Sub

EOX

}

1;
