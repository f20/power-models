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
