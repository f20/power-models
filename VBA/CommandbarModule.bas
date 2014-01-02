'
' Copyright 2013 Franck Latrémolière, Reckon LLP and others.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'
' 1. Redistributions of source code must retain the above copyright notice,
' this list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright notice,
' this list of conditions and the following disclaimer in the documentation
' and/or other materials provided with the distribution.
'
' THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
' EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
' THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'

Sub MakeCommandBar()
    On Error Resume Next
    Application.CommandBars("Franck Spreadsheet Tools").Delete
    On Error GoTo 0
    Dim bar As CommandBar
    Set bar = Application.CommandBars.Add("Franck Spreadsheet Tools")
    
    Dim btn As CommandBarButton
    
    Set btn = bar.Controls.Add(Type:=msoControlButton)
    With btn
     .Style = msoButtonCaption
     .Caption = "InputDB..."
     .BeginGroup = True
     .OnAction = "ShowInputDBUI"
    End With

    Set btn = bar.Controls.Add(Type:=msoControlButton)
    With btn
     .Style = msoButtonCaption
     .Caption = "Show Formulas..."
     .BeginGroup = True
     .OnAction = "ShowFormulas"
    End With
        
    Set btn = bar.Controls.Add(Type:=msoControlButton)
    With btn
     .Style = msoButtonCaption
     .Caption = "Sheet list"
     .OnAction = "ListWorksheets"
    End With
    
    Set btn = bar.Controls.Add(Type:=msoControlButton)
    With btn
     .Style = msoButtonCaption
     .Caption = "VBACode"
     .BeginGroup = True
     .OnAction = "MakeVBAModules"
    End With
        
    Set btn = bar.Controls.Add(Type:=msoControlButton)
    With btn
     .Style = msoButtonCaption
     .BeginGroup = True
     .Caption = "Exp"
     .OnAction = "ExportVBA"
    End With

    Set btn = bar.Controls.Add(Type:=msoControlButton)
    With btn
     .Style = msoButtonCaption
     .Caption = "Imp"
     .OnAction = "ImportVBA"
    End With
    
    bar.Visible = True

End Sub
