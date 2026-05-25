Sub AddCheckboxes()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    Application.ScreenUpdating = False

    Dim NAVY As Long: NAVY = RGB(31, 56, 100)
    Dim BLUE As Long: BLUE = RGB(46, 114, 182)

    Dim r As Long
    Dim bg As Long
    Dim chk As CheckBox
    Dim cellLeft As Double
    Dim cellTop As Double
    Dim cellWidth As Double
    Dim cellHeight As Double

    For r = 18 To 300
        ' Check background colour of col A — skip headers
        bg = wsMM.Cells(r, 1).Interior.Color
        If bg = NAVY Or bg = BLUE Then GoTo NextRow

        ' Skip if row is empty — check col B
        If Trim(wsMM.Cells(r, 2).Value) = "" And _
           Trim(wsMM.Cells(r, 1).Value) = "" Then GoTo NextRow

        ' Add checkbox to cols C, D, E
        Dim col As Integer
        For col = 3 To 5
            cellLeft = wsMM.Cells(r, col).Left
            cellTop = wsMM.Cells(r, col).Top
            cellWidth = wsMM.Cells(r, col).Width
            cellHeight = wsMM.Cells(r, col).Height

            ' Add checkbox — centred in cell
            Set chk = wsMM.CheckBoxes.Add( _
                cellLeft + (cellWidth - 12) / 2, _
                cellTop + (cellHeight - 12) / 2, _
                12, 12)

            With chk
                .Caption = ""
                .LinkedCell = wsMM.Cells(r, col).Address
                .Display3DShading = False
            End With
        Next col

NextRow:
    Next r

    Application.ScreenUpdating = True
    MsgBox "Checkboxes added to cols C, D, E rows 18-300." & Chr(10) & _
           "Skipped section and subsection header rows.", vbInformation
End Sub
Sub HideTrueFalseText()
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ws.Range("C9:C14").NumberFormat = ";;;"
    ws.Range("C18:C" & lastRow).NumberFormat = ";;;"
    ws.Range("D18:D" & lastRow).NumberFormat = ";;;"
    ws.Range("E18:E" & lastRow).NumberFormat = ";;;"
    
    MsgBox "Done!", vbInformation
End Sub
Sub AddRow()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    
    Dim activeRow As Long
    activeRow = ActiveCell.Row
    
    If activeRow < 18 Then
        MsgBox "Please click a row in the data area first.", vbExclamation
        Exit Sub
    End If
    
    ' Insert row below current row
    wsMM.Rows(activeRow + 1).Insert Shift:=xlDown
    
    ' Clear everything on new row
    wsMM.Rows(activeRow + 1).ClearContents
    wsMM.Rows(activeRow + 1).ClearFormats
    wsMM.Rows(activeRow + 1).rowHeight = 15
    
    ' Work out correct alternating shade
    Dim dataRow As Long
    Dim toggle As Boolean
    toggle = True
    For dataRow = 18 To activeRow + 1
        Dim bg As Long
        bg = wsMM.Cells(dataRow, 1).Interior.Color
        If bg = RGB(31, 56, 100) Or bg = RGB(46, 114, 182) Then
            toggle = True
        Else
            If dataRow = activeRow + 1 Then
                If toggle Then
                    wsMM.Range("A" & (activeRow + 1) & ":K" & (activeRow + 1)).Interior.Color = RGB(250, 250, 250)
                Else
                    wsMM.Range("A" & (activeRow + 1) & ":K" & (activeRow + 1)).Interior.Color = RGB(242, 242, 242)
                End If
            Else
                toggle = Not toggle
            End If
        End If
    Next dataRow
    
    ' Set font
    With wsMM.Range("A" & (activeRow + 1) & ":K" & (activeRow + 1))
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Color = RGB(26, 26, 46)
        .Font.bold = False
        .VerticalAlignment = xlCenter
    End With
    wsMM.Range("C" & (activeRow + 1) & ":J" & (activeRow + 1)).HorizontalAlignment = xlCenter
    wsMM.Range("A" & (activeRow + 1)).HorizontalAlignment = xlLeft
    wsMM.Range("B" & (activeRow + 1)).HorizontalAlignment = xlLeft
    wsMM.Range("K" & (activeRow + 1)).HorizontalAlignment = xlLeft
    
    ' Days Overdue formula col J
    Dim nr As Long
    nr = activeRow + 1
    wsMM.Cells(nr, 10).Formula = _
        "=IF(OR(H" & nr & "=""Done"",H" & nr & "=""On Hold"",H" & nr & "=""Waiting""),"""",IF(NOT(C" & nr & "),"""",IF(G" & nr & "="""","""",IF(TODAY()>G" & nr & ",TODAY()-G" & nr & ",""""))))"
    wsMM.Cells(nr, 10).NumberFormat = "0"
    wsMM.Cells(nr, 10).HorizontalAlignment = xlCenter
    
    ' Copy hidden project col L from current row
    wsMM.Cells(nr, 12).Value = wsMM.Cells(activeRow, 12).Value
    
    ' Add checkbox col C — Action?
    Dim chkC As CheckBox
    Dim cellC As Range
    Set cellC = wsMM.Cells(nr, 3)
    Set chkC = wsMM.CheckBoxes.Add(cellC.Left, cellC.Top, cellC.Width, cellC.Height)
    With chkC
        .Caption = ""
        .LinkedCell = cellC.Address
        .Name = "chkAction_" & nr
    End With
    ' Hide the TRUE/FALSE text
    cellC.NumberFormat = ";;;"
    
    ' Add checkbox col D — Risk?
    Dim chkD As CheckBox
    Dim cellD As Range
    Set cellD = wsMM.Cells(nr, 4)
    Set chkD = wsMM.CheckBoxes.Add(cellD.Left, cellD.Top, cellD.Width, cellD.Height)
    With chkD
        .Caption = ""
        .LinkedCell = cellD.Address
        .Name = "chkRisk_" & nr
    End With
    ' Hide the TRUE/FALSE text
    cellD.NumberFormat = ";;;"
    
    ' Copy data validations from row above — Owner, Status, Priority dropdowns
    wsMM.Cells(activeRow, 6).Copy
    wsMM.Cells(nr, 6).PasteSpecial Paste:=xlPasteValidation
    wsMM.Cells(activeRow, 8).Copy
    wsMM.Cells(nr, 8).PasteSpecial Paste:=xlPasteValidation
    wsMM.Cells(activeRow, 9).Copy
    wsMM.Cells(nr, 9).PasteSpecial Paste:=xlPasteValidation
    Application.CutCopyMode = False
    
    ' Activate new row col B
    wsMM.Cells(nr, 2).Activate
End Sub
Sub DiagnoseRows()
    Dim ws As Worksheet
    Dim r As Long
    Dim msg As String
    Set ws = Sheets("Meeting Minutes")
    
    For r = ActiveCell.Row + 1 To ActiveCell.Row + 10
        msg = msg & "Row " & r & ": MergeCells=" & ws.Cells(r, 1).MergeCells & Chr(10)
    Next r
    
    MsgBox msg
End Sub

Sub AddSubsection()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    
    Dim activeRow As Long
    activeRow = ActiveCell.Row
    
    If activeRow < 18 Then
        MsgBox "Please click a row in the data area first.", vbExclamation
        Exit Sub
    End If
    
    ' Ask for subsection title
    Dim title As String
    title = InputBox("Enter subsection title:" & Chr(10) & Chr(10) & "e.g. The Avenue  |  Budget: $1.95M  |  PM: TM1 / TM4", "Add Subsection")
    If title = "" Then Exit Sub
    
    ' Insert 2 rows — subsection header + first blank data row
    wsMM.Rows(activeRow + 1).Insert Shift:=xlDown
    wsMM.Rows(activeRow + 1).Insert Shift:=xlDown
    
    ' Clear both rows
    wsMM.Rows(activeRow + 1).ClearContents
    wsMM.Rows(activeRow + 1).ClearFormats
    wsMM.Rows(activeRow + 2).ClearContents
    wsMM.Rows(activeRow + 2).ClearFormats
    
    ' Subsection header — blue, merged A:K
    wsMM.Range("A" & (activeRow + 1) & ":K" & (activeRow + 1)).Merge
    With wsMM.Range("A" & (activeRow + 1) & ":K" & (activeRow + 1))
        .Interior.Color = RGB(46, 114, 182)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .Font.Name = "Arial"
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Value = title
    End With
    wsMM.Rows(activeRow + 1).rowHeight = 15
    
    ' First data row below — white, standard formatting
    With wsMM.Range("A" & (activeRow + 2) & ":K" & (activeRow + 2))
        .Interior.Color = RGB(250, 250, 250)
        .Font.Name = "Arial": .Font.Size = 9
        .Font.Color = RGB(26, 26, 46): .Font.bold = False
        .VerticalAlignment = xlCenter
    End With
    wsMM.Range("C" & (activeRow + 2) & ":J" & (activeRow + 2)).HorizontalAlignment = xlCenter
    wsMM.Range("A" & (activeRow + 2)).HorizontalAlignment = xlLeft
    wsMM.Range("B" & (activeRow + 2)).HorizontalAlignment = xlLeft
    wsMM.Range("K" & (activeRow + 2)).HorizontalAlignment = xlLeft
    wsMM.Rows(activeRow + 2).rowHeight = 15
    
    ' Days Overdue formula in data row
    Dim nr As Long: nr = activeRow + 2
    wsMM.Cells(nr, 10).Formula = _
        "=IF(OR(H" & nr & "=""Done"",H" & nr & "=""On Hold"",H" & nr & "=""Waiting""),"""",IF(NOT(C" & nr & "),"""",IF(G" & nr & "="""","""",IF(TODAY()>G" & nr & ",TODAY()-G" & nr & ",""""))))"
    wsMM.Cells(nr, 10).NumberFormat = "0"
    wsMM.Cells(nr, 10).HorizontalAlignment = xlCenter
    
    ' Copy project name to hidden col L
    wsMM.Cells(nr, 12).Value = wsMM.Cells(activeRow, 12).Value
    
    ' Activate first data row col B
    wsMM.Cells(nr, 2).Activate
End Sub
' ---- 4. PopulateActionRegister_AddOnly ----------------------
' MM ? AR sync. ADD ONLY — never updates existing AR rows.
' Once an action exists in AR (matched on Description+Owner),
' AR owns it. MM only contributes brand-new actions.
Sub PopulateActionRegister()
    Dim wsMin As Worksheet
    Dim wsAR As Worksheet
    Dim r As Long
    Dim arRow As Long
    Dim arKeys As Object
    Dim mmKey As String
    Dim added As Long

    Set wsMin = ThisWorkbook.Sheets("Meeting Minutes")
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    Set arKeys = CreateObject("Scripting.Dictionary")
    added = 0

    ' Build set of existing AR keys (Description|Owner)
    For arRow = 4 To 103
        If Trim(CStr(wsAR.Cells(arRow, 2).Value)) <> "" Then
            arKeys(Trim(CStr(wsAR.Cells(arRow, 2).Value)) & "|" & _
                   Trim(CStr(wsAR.Cells(arRow, 3).Value))) = True
        End If
    Next arRow

    ' Find first empty AR row
    arRow = 4
    Do While Trim(CStr(wsAR.Cells(arRow, 2).Value)) <> "" And arRow <= 103
        arRow = arRow + 1
    Loop

    ' Scan MM for ticked actions not already in AR
    For r = 15 To 300
        If wsMin.Cells(r, 1).MergeCells Then GoTo NextRow
        If wsMin.Cells(r, 3).Value <> True Then GoTo NextRow
        If Trim(CStr(wsMin.Cells(r, 2).Value)) = "" Then GoTo NextRow

        mmKey = Trim(CStr(wsMin.Cells(r, 2).Value)) & "|" & _
                Trim(CStr(wsMin.Cells(r, 6).Value))

        ' Skip if already in AR
        If arKeys.Exists(mmKey) Then GoTo NextRow

        If arRow > 103 Then
            MsgBox "Action Register full at row 103. Some MM actions not added.", vbExclamation
            Exit For
        End If

        ' Add new row
        wsAR.Cells(arRow, 1).Value = wsMin.Cells(r, 12).Value   ' Project
        wsAR.Cells(arRow, 2).Value = wsMin.Cells(r, 2).Value    ' Description
        wsAR.Cells(arRow, 3).Value = wsMin.Cells(r, 6).Value    ' Owner
        wsAR.Cells(arRow, 4).Value = wsMin.Cells(r, 7).Value    ' Due Date
        wsAR.Cells(arRow, 4).NumberFormat = "DD/MM/YYYY"
        wsAR.Cells(arRow, 5).Value = wsMin.Cells(r, 9).Value    ' Priority
        wsAR.Cells(arRow, 6).Value = wsMin.Cells(r, 8).Value    ' Status
        wsAR.Cells(arRow, 7).Value = Date                       ' Date Added
        wsAR.Cells(arRow, 7).NumberFormat = "DD/MM/YYYY"
        wsAR.Cells(arRow, 8).Value = wsMin.Cells(r, 11).Value   ' Notes
        wsAR.Cells(arRow, 9).Formula = _
            "=IF(OR(F" & arRow & "=""Done"",F" & arRow & "=""On Hold"",F" & arRow & "=""Waiting""),"""",IF(D" & arRow & "="""","""",IF(TODAY()>D" & arRow & ",TODAY()-D" & arRow & ","""")))"
        wsAR.Cells(arRow, 9).NumberFormat = "0"

        ' Mark as now-existing so duplicate MM rows don't add twice in same run
        arKeys(mmKey) = True

        arRow = arRow + 1
        added = added + 1

NextRow:
    Next r

    ' (Optional) — could MsgBox added count, but SyncAll will summarise
End Sub

Sub SyncStatusBack()
    Dim wsMin As Worksheet
    Dim wsAR As Worksheet
    Dim arRow As Long
    Dim minRow As Long
    Dim arDesc As String
    Dim arOwner As String
    Dim arStatus As String
    Dim arPriority As String
    Dim arNotes As String
    Dim arDueDate As Variant
    Dim updated As Long
    Dim notFound As Long
    Dim found As Boolean
    updated = 0
    notFound = 0
    Set wsMin = ThisWorkbook.Sheets("Meeting Minutes")
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    For arRow = 4 To 103
        arDesc = Trim(wsAR.Cells(arRow, 2).Value)    ' AR col B — Description
        arOwner = Trim(wsAR.Cells(arRow, 3).Value)   ' AR col C — Owner
        arDueDate = wsAR.Cells(arRow, 4).Value        ' AR col D — Due Date
        arPriority = Trim(wsAR.Cells(arRow, 5).Value) ' AR col E — Priority
        arStatus = Trim(wsAR.Cells(arRow, 6).Value)   ' AR col F — Status
        arNotes = Trim(wsAR.Cells(arRow, 8).Value)    ' AR col H — Notes
        If arDesc = "" Then GoTo NextAR
        found = False
        For minRow = 18 To 300
            If wsMin.Cells(minRow, 1).MergeCells Then GoTo NextMin
            If Trim(wsMin.Cells(minRow, 2).Value) = arDesc And _
               Trim(wsMin.Cells(minRow, 6).Value) = arOwner Then  ' MM col F — Owner
                wsMin.Cells(minRow, 8).Value = arStatus            ' MM col H — Status
                wsMin.Cells(minRow, 9).Value = arPriority          ' MM col I — Priority
                wsMin.Cells(minRow, 11).Value = arNotes            ' MM col K — Notes
                If arDueDate <> "" Then
                    wsMin.Cells(minRow, 7).Value = arDueDate       ' MM col G — Due Date
                    wsMin.Cells(minRow, 7).NumberFormat = "DD/MM/YYYY"
                End If
                updated = updated + 1
                found = True
                GoTo NextAR
            End If
NextMin:
        Next minRow
        If Not found Then notFound = notFound + 1
NextAR:
    Next arRow
    MsgBox "Sync complete." & Chr(10) & Chr(10) & _
           updated & " actions updated in Minutes." & Chr(10) & _
           notFound & " actions in Register not matched in Minutes." & Chr(10) & Chr(10) & _
           "Check unmatched rows — description or owner may have changed.", _
           vbInformation, "SyncStatusBack"
End Sub
Sub NewWeek()
    Dim wb As Workbook
    Dim currentPath As String
    Dim newFileName As String
    Dim nextTuesday As Date
    Dim today As Date
    Dim daysUntilTuesday As Integer

    Set wb = ThisWorkbook
    today = Date

    daysUntilTuesday = (3 - Weekday(today, vbMonday) + 7) Mod 7
    If daysUntilTuesday = 0 Then daysUntilTuesday = 7
    nextTuesday = today + daysUntilTuesday

    newFileName = "Weekly_Meeting_" & Format(nextTuesday, "DDMMMYYYY") & ".xlsm"
    currentPath = Left(wb.FullName, InStrRev(wb.FullName, "\"))

    wb.SaveAs Filename:=currentPath & newFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled

    MsgBox "Saved as: " & newFileName & Chr(10) & Chr(10) & _
           "Location: " & currentPath & Chr(10) & Chr(10) & _
           "SharePoint will sync automatically.", _
           vbInformation, "New Week File Created"
End Sub

Sub ToggleDoneRegister()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim hidingDone As Boolean
    
    Set ws = ThisWorkbook.Sheets("Action Register")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Check current state — find any visible Done row
    hidingDone = False
    For i = 4 To lastRow
        If ws.Cells(i, 6).Value = "Done" And Not ws.Rows(i).Hidden Then
            hidingDone = True
            Exit For
        End If
    Next i
    
    ' Hide or show all Done rows
    Application.ScreenUpdating = False
    For i = 4 To lastRow
        If ws.Cells(i, 6).Value = "Done" Then
            ws.Rows(i).Hidden = hidingDone
        End If
    Next i
    Application.ScreenUpdating = True
    
    ' Update button text
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.OnAction = "ToggleDoneRegister" Then
            If hidingDone Then
                shp.TextFrame.Characters.Text = "Show Done"
            Else
                shp.TextFrame.Characters.Text = "Hide / Show Done"
            End If
        End If
    Next shp

End Sub
Sub AddSection()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    
    Dim activeRow As Long
    activeRow = ActiveCell.Row
    
    If activeRow < 18 Then
        MsgBox "Please click a row in the data area first.", vbExclamation
        Exit Sub
    End If
    
    ' Insert 2 rows — section header + first blank data row
    wsMM.Rows(activeRow + 1).Insert Shift:=xlDown
    wsMM.Rows(activeRow + 1).Insert Shift:=xlDown
    
    ' Clear both rows first
    wsMM.Rows(activeRow + 1).ClearContents
    wsMM.Rows(activeRow + 1).ClearFormats
    wsMM.Rows(activeRow + 2).ClearContents
    wsMM.Rows(activeRow + 2).ClearFormats
    
    ' Section header — navy, merged A:K
    wsMM.Range("A" & (activeRow + 1) & ":K" & (activeRow + 1)).Merge
    With wsMM.Range("A" & (activeRow + 1) & ":K" & (activeRow + 1))
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .Font.Name = "Arial"
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .rowHeight = 15
    End With
    wsMM.Rows(activeRow + 1).rowHeight = 15
    
    ' First data row below — white, standard formatting
    With wsMM.Range("A" & (activeRow + 2) & ":K" & (activeRow + 2))
        .Interior.Color = RGB(250, 250, 250)
        .Font.Name = "Arial": .Font.Size = 9
        .Font.Color = RGB(26, 26, 46): .Font.bold = False
        .VerticalAlignment = xlCenter
    End With
    wsMM.Range("C" & (activeRow + 2) & ":J" & (activeRow + 2)).HorizontalAlignment = xlCenter
    wsMM.Range("A" & (activeRow + 2)).HorizontalAlignment = xlLeft
    wsMM.Range("B" & (activeRow + 2)).HorizontalAlignment = xlLeft
    wsMM.Range("K" & (activeRow + 2)).HorizontalAlignment = xlLeft
    wsMM.Rows(activeRow + 2).rowHeight = 15
    
    ' Days Overdue formula in data row
    Dim nr As Long: nr = activeRow + 2
    wsMM.Cells(nr, 10).Formula = _
        "=IF(OR(H" & nr & "=""Done"",H" & nr & "=""On Hold"",H" & nr & "=""Waiting""),"""",IF(NOT(C" & nr & "),"""",IF(G" & nr & "="""","""",IF(TODAY()>G" & nr & ",TODAY()-G" & nr & ",""""))))"
    wsMM.Cells(nr, 10).NumberFormat = "0"
    wsMM.Cells(nr, 10).HorizontalAlignment = xlCenter
    
    ' Activate section header for typing
    wsMM.Cells(activeRow + 1, 1).Activate
End Sub
Sub AddPriorityDropdowns()
    Dim wsAR As Worksheet
    Dim wsMM As Worksheet
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")

    ' Action Register — Priority col E
    wsAR.Range("E4:E103").Validation.Delete
    wsAR.Range("E4:E103").Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="Critical,High,Medium,Low"

    ' Meeting Minutes — Priority col G
    wsMM.Range("G22:G300").Validation.Delete
    wsMM.Range("G22:G300").Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="Critical,High,Medium,Low"

    MsgBox "Priority dropdowns added to AR col E and MM col G.", vbInformation
End Sub

Sub AddOverdueFormulas()
    Dim wsAR As Worksheet
    Dim wsMM As Worksheet
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")

    ' Action Register — Days Overdue col I
    ' Blank if Done/On Hold/Waiting, blank if no due date, blank if not yet due
    ' F=Status, D=Due Date
    wsAR.Range("I4:I103").Formula = _
        "=IF(OR(F4=""Done"",F4=""On Hold"",F4=""Waiting""),"""",IF(D4="""","""",IF(TODAY()>D4,TODAY()-D4,"""")))"
    wsAR.Range("I4:I103").NumberFormat = "0"
    wsAR.Range("I4:I103").HorizontalAlignment = xlCenter

    ' Meeting Minutes — Days Overdue col J
    ' Blank if Done/On Hold/Waiting, blank if not an action, blank if no due date, blank if not yet due
    ' F=Status, C=Action?, E=Due Date
    wsMM.Range("J22:J300").Formula = _
        "=IF(OR(F22=""Done"",F22=""On Hold"",F22=""Waiting""),"""",IF(NOT(C22),"""",IF(E22="""","""",IF(TODAY()>E22,TODAY()-E22,""""))))"
    wsMM.Range("J22:J300").NumberFormat = "0"
    wsMM.Range("J22:J300").HorizontalAlignment = xlCenter

    MsgBox "Days Overdue formulas added." & Chr(10) & Chr(10) & _
        "AR col I — checks Status(F) and Due Date(D)" & Chr(10) & _
        "MM col J — checks Action?(C) Status(F) and Due Date(E)", vbInformation
End Sub

Sub FixMMColumns()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")

    ' Remove wrong Status dropdown from col H (Days Overdue column)
    wsMM.Range("H22:H301").Validation.Delete

    ' Add correct Days Overdue formula to col H
    wsMM.Range("H22:H300").Formula = _
        "=IF(OR(F22=""Done"",F22=""On Hold"",F22=""Waiting""),"""",IF(NOT(C22),"""",IF(E22="""","""",IF(TODAY()>E22,TODAY()-E22,""""))))"
    wsMM.Range("H22:H300").NumberFormat = "0"
    wsMM.Range("H22:H300").HorizontalAlignment = xlCenter

    MsgBox "Done — Status dropdown removed from col H, Days Overdue formula added.", vbInformation
End Sub

Sub AddCF()
    Dim wsAR As Worksheet
    Dim wsMM As Worksheet
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")

    Dim WHITE     As Long: WHITE = RGB(255, 255, 255)
    Dim P_CRIT    As Long: P_CRIT = RGB(192, 0, 0)
    Dim P_HIGH    As Long: P_HIGH = RGB(237, 125, 49)
    Dim P_MED     As Long: P_MED = RGB(255, 192, 0)
    Dim P_LOW     As Long: P_LOW = RGB(169, 209, 142)
    Dim P_MED_FG  As Long: P_MED_FG = RGB(125, 82, 0)
    Dim P_LOW_FG  As Long: P_LOW_FG = RGB(55, 86, 35)
    Dim S_OPEN_FG As Long: S_OPEN_FG = RGB(89, 89, 89)
    Dim S_IP_FG   As Long: S_IP_FG = RGB(125, 82, 0)
    Dim S_DONE_FG As Long: S_DONE_FG = RGB(55, 86, 35)
    Dim S_HOLD_FG As Long: S_HOLD_FG = RGB(31, 56, 100)
    Dim S_WAIT_FG As Long: S_WAIT_FG = RGB(132, 60, 12)
    Dim OV_RED    As Long: OV_RED = RGB(192, 0, 0)
    Dim OV_AMB    As Long: OV_AMB = RGB(237, 125, 49)
    Dim OV_ROW    As Long: OV_ROW = RGB(255, 199, 206)
    Dim OV_FG     As Long: OV_FG = RGB(156, 0, 6)
    Dim ACT_BG    As Long: ACT_BG = RGB(255, 243, 205)
    Dim ACT_FG    As Long: ACT_FG = RGB(125, 82, 0)

    ' -- ACTION REGISTER CF ---------------------------------------

    ' Priority col E
    wsAR.Range("E4:E103").FormatConditions.Delete
    With wsAR.Range("E4:E103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Critical""")
        .Interior.Pattern = xlSolid: .Interior.Color = P_CRIT
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("E4:E103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""High""")
        .Interior.Pattern = xlSolid: .Interior.Color = P_HIGH
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("E4:E103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Medium""")
        .Interior.Pattern = xlSolid: .Interior.Color = P_MED
        .Font.Color = P_MED_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("E4:E103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Low""")
        .Interior.Pattern = xlSolid: .Interior.Color = P_LOW
        .Font.Color = P_LOW_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Status col F — font only
    wsAR.Range("F4:F103").FormatConditions.Delete
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Open""")
        .Font.Color = S_OPEN_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""In Progress""")
        .Font.Color = S_IP_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Done""")
        .Font.Color = S_DONE_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""On Hold""")
        .Font.Color = S_HOLD_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Waiting""")
        .Font.Color = S_WAIT_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Overdue row A:I — whole row pink when col I > 0
    wsAR.Range("A4:I103").FormatConditions.Delete
    With wsAR.Range("A4:I103").FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND($I4<>"""",$I4>0)")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_ROW
        .Font.Color = OV_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Days Overdue cell col I — amber/red on top of row
    wsAR.Range("I4:I103").FormatConditions.Delete
    With wsAR.Range("I4:I103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="8")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_RED
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("I4:I103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="1", Formula2:="7")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_AMB
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With

    ' -- MEETING MINUTES CF ---------------------------------------

    ' Action rows amber — C=Action? col, F=Status, E=Due Date
    wsMM.Range("A22:G300").FormatConditions.Delete
    With wsMM.Range("A22:G300").FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=$C22=TRUE")
        .Interior.Pattern = xlSolid: .Interior.Color = ACT_BG
        .Font.Color = ACT_FG: .StopIfTrue = False
    End With

    ' Overdue action rows red — overrides amber
    With wsMM.Range("A22:G300").FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND($C22=TRUE,NOT(OR($F22=""Done"",$F22=""On Hold"",$F22=""Waiting"")),$E22<>"""",$E22<TODAY())")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_ROW
        .Font.Color = OV_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Priority col G
    wsMM.Range("G22:G300").FormatConditions.Delete
    With wsMM.Range("G22:G300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Critical""")
        .Interior.Pattern = xlSolid: .Interior.Color = P_CRIT
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("G22:G300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""High""")
        .Interior.Pattern = xlSolid: .Interior.Color = P_HIGH
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("G22:G300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Medium""")
        .Interior.Pattern = xlSolid: .Interior.Color = P_MED
        .Font.Color = P_MED_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("G22:G300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Low""")
        .Interior.Pattern = xlSolid: .Interior.Color = P_LOW
        .Font.Color = P_LOW_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Status col F — font only
    wsMM.Range("F22:F300").FormatConditions.Delete
    With wsMM.Range("F22:F300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Open""")
        .Font.Color = S_OPEN_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("F22:F300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""In Progress""")
        .Font.Color = S_IP_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("F22:F300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Done""")
        .Font.Color = S_DONE_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("F22:F300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""On Hold""")
        .Font.Color = S_HOLD_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("F22:F300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="""Waiting""")
        .Font.Color = S_WAIT_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Days Overdue cell col H — amber/red
    wsMM.Range("H22:H300").FormatConditions.Delete
    With wsMM.Range("H22:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="8")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_RED
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H22:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="1", Formula2:="7")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_AMB
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With

    MsgBox "CF added to both sheets." & Chr(10) & Chr(10) & _
        "Reminder: set fill colours manually in Manage Rules for any rules showing transparent.", _
        vbInformation, "AddCF Done"
End Sub
Sub FixOverdueCF()
    Dim wsAR As Worksheet
    Dim wsMM As Worksheet
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")

    Dim WHITE  As Long: WHITE = RGB(255, 255, 255)
    Dim OV_RED As Long: OV_RED = RGB(192, 0, 0)
    Dim OV_AMB As Long: OV_AMB = RGB(237, 125, 49)

    ' Fix AR col I — only fire when cell has a number
    wsAR.Range("I4:I103").FormatConditions.Delete
    With wsAR.Range("I4:I103").FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($I4<>"""",$I4>=8)")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_RED
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("I4:I103").FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($I4<>"""",$I4>=1,$I4<=7)")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_AMB
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With

    ' Fix MM col H — same
    wsMM.Range("H22:H300").FormatConditions.Delete
    With wsMM.Range("H22:H300").FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($H22<>"""",$H22>=8)")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_RED
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H22:H300").FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($H22<>"""",$H22>=1,$H22<=7)")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_AMB
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With

    MsgBox "Overdue CF fixed — blank cells no longer trigger red.", vbInformation
End Sub
Sub FixSummaryOverdue()
    Dim wsSum As Worksheet
    Set wsSum = ThisWorkbook.Sheets("Summary")

    Dim r As Long
    Dim totRow As Long
    Dim lastRow As Long
    Dim ownerRef As String

    ' Find TOTAL row
    lastRow = wsSum.Cells(wsSum.Rows.Count, 1).End(xlUp).Row
    For r = lastRow To 1 Step -1
        If wsSum.Cells(r, 1).Value = "TOTAL" Then
            totRow = r
            Exit For
        End If
    Next r
    If totRow = 0 Then totRow = 12

    ' Overdue formula — AR Owner=col C, Status=col F, DueDate=col D
    ' Only count Open + In Progress past due date
    For r = 4 To totRow - 1
        ownerRef = "A" & r
        wsSum.Cells(r, 8).Formula = _
            "=COUNTIFS('Action Register'!C4:C103," & ownerRef & "," & _
            "'Action Register'!F4:F103,""Open""," & _
            "'Action Register'!D4:D103,""<""&TODAY())" & _
            "+COUNTIFS('Action Register'!C4:C103," & ownerRef & "," & _
            "'Action Register'!F4:F103,""In Progress""," & _
            "'Action Register'!D4:D103,""<""&TODAY())"
        wsSum.Cells(r, 8).NumberFormat = "0"
        wsSum.Cells(r, 8).HorizontalAlignment = xlCenter
        wsSum.Cells(r, 8).Font.Name = "Arial"
        wsSum.Cells(r, 8).Font.Size = 9
    Next r

    ' CF — red when > 0
    wsSum.Range("H4:H" & totRow).FormatConditions.Delete
    With wsSum.Range("H4:H" & totRow).FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .StopIfTrue = False
    End With

    MsgBox "Summary overdue column updated.", vbInformation
End Sub

Sub FixMMOwnerDropdown()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")

    wsMM.Range("D22:D301").Validation.Delete
    wsMM.Range("D22:D301").Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="TM1,TM2,TM3,TM4,TM5,TM6,ALL"

    MsgBox "MM owner dropdown updated.", vbInformation
End Sub
Sub FormatMM()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    Application.ScreenUpdating = False

    Dim NAVY     As Long: NAVY = RGB(31, 56, 100)
    Dim BLUE     As Long: BLUE = RGB(46, 114, 182)
    Dim WHITE    As Long: WHITE = RGB(255, 255, 255)
    Dim P_CRIT   As Long: P_CRIT = RGB(192, 0, 0)
    Dim P_HIGH   As Long: P_HIGH = RGB(237, 125, 49)
    Dim P_MED    As Long: P_MED = RGB(255, 192, 0)
    Dim P_LOW    As Long: P_LOW = RGB(169, 209, 142)
    Dim P_MED_FG As Long: P_MED_FG = RGB(125, 82, 0)
    Dim P_LOW_FG As Long: P_LOW_FG = RGB(55, 86, 35)
    Dim S_OPEN_FG  As Long: S_OPEN_FG = RGB(89, 89, 89)
    Dim S_IP_FG    As Long: S_IP_FG = RGB(125, 82, 0)
    Dim S_DONE_FG  As Long: S_DONE_FG = RGB(55, 86, 35)
    Dim S_HOLD_FG  As Long: S_HOLD_FG = RGB(31, 56, 100)
    Dim S_WAIT_FG  As Long: S_WAIT_FG = RGB(132, 60, 12)
    Dim ACT_BG   As Long: ACT_BG = RGB(255, 243, 205)
    Dim ACT_FG   As Long: ACT_FG = RGB(125, 82, 0)
    Dim OV_BG    As Long: OV_BG = RGB(255, 199, 206)
    Dim OV_FG    As Long: OV_FG = RGB(156, 0, 6)
    Dim OV_RED   As Long: OV_RED = RGB(192, 0, 0)
    Dim OV_AMB   As Long: OV_AMB = RGB(237, 125, 49)
    Dim Q As String: Q = Chr(34)

    ' -- COLUMN WIDTHS ---------------------------------------------
    wsMM.Columns(1).ColumnWidth = 16
    wsMM.Columns(2).ColumnWidth = 58
    wsMM.Columns(3).ColumnWidth = 7
    wsMM.Columns(4).ColumnWidth = 7
    wsMM.Columns(5).ColumnWidth = 8
    wsMM.Columns(6).ColumnWidth = 8
    wsMM.Columns(7).ColumnWidth = 11
    wsMM.Columns(8).ColumnWidth = 13
    wsMM.Columns(9).ColumnWidth = 11
    wsMM.Columns(10).ColumnWidth = 11
    wsMM.Columns(11).ColumnWidth = 26
    wsMM.Columns(12).ColumnWidth = 0.5

    ' -- ROW HEIGHTS -----------------------------------------------
    wsMM.Rows(21).rowHeight = 28
    wsMM.Rows("23:300").rowHeight = 15

    ' -- HEADER ROW 21 ---------------------------------------------
    With wsMM.Range("A21:K21")
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
        .Font.Color = WHITE: .Interior.Color = NAVY
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    wsMM.Range("B21").HorizontalAlignment = xlLeft

    ' -- DATA RANGE DEFAULT FONT -----------------------------------
    With wsMM.Range("A23:K300")
        .Font.Name = "Arial": .Font.Size = 9
        .VerticalAlignment = xlCenter: .WrapText = False
        .Font.Color = RGB(26, 26, 46)
        .Interior.ColorIndex = xlNone
    End With
    wsMM.Range("C23:J300").HorizontalAlignment = xlCenter
    wsMM.Range("A23:B300").HorizontalAlignment = xlLeft
    wsMM.Range("K23:K300").HorizontalAlignment = xlLeft
    wsMM.Range("G23:G300").NumberFormat = "DD/MM/YYYY"

    ' -- RESTORE SECTION HEADERS -----------------------------------
    ' Scan for merged rows and restore navy/blue formatting
    Dim r As Long
    For r = 22 To 300
        If wsMM.Cells(r, 1).MergeCells Then
            Dim cellVal As String
            cellVal = Trim(wsMM.Cells(r, 1).Value)
            If cellVal = "" Then GoTo NextRow
            ' Check if it's a sub-header (contains |) or main section header
            If InStr(cellVal, "|") > 0 Then
                ' Sub-header — blue
                With wsMM.Cells(r, 1).MergeArea
                    .Interior.Color = BLUE
                    .Font.Color = WHITE
                    .Font.bold = True
                    .Font.Name = "Arial"
                    .Font.Size = 9
                End With
            ElseIf Left(cellVal, 1) >= "1" And Left(cellVal, 1) <= "9" Then
                ' Main section header — navy
                With wsMM.Cells(r, 1).MergeArea
                    .Interior.Color = NAVY
                    .Font.Color = WHITE
                    .Font.bold = True
                    .Font.Name = "Arial"
                    .Font.Size = 9
                End With
            End If
        End If
NextRow:
    Next r

    ' -- RESTORE CHECKBOX DISPLAY ----------------------------------
    ' Set col C D E to show TRUE/FALSE as checkboxes via centre alignment
    ' The actual checkbox objects are form controls — just fix the font
    With wsMM.Range("C23:E300")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
    End With

    ' -- CLEAR ALL CF ----------------------------------------------
    wsMM.Cells.FormatConditions.Delete

    ' -- CF RULES — Priority FIRST so badges show over amber rows --

    ' Priority col I — written FIRST so they take precedence
    With wsMM.Range("I23:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=Q & "Critical" & Q)
        .Interior.Pattern = xlSolid: .Interior.Color = P_CRIT
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("I23:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=Q & "High" & Q)
        .Interior.Pattern = xlSolid: .Interior.Color = P_HIGH
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("I23:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=Q & "Medium" & Q)
        .Interior.Pattern = xlSolid: .Interior.Color = P_MED
        .Font.Color = P_MED_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("I23:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=Q & "Low" & Q)
        .Interior.Pattern = xlSolid: .Interior.Color = P_LOW
        .Font.Color = P_LOW_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Status col H — font colours only
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=Q & "Open" & Q)
        .Font.Color = S_OPEN_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=Q & "In Progress" & Q)
        .Font.Color = S_IP_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=Q & "Done" & Q)
        .Font.Color = S_DONE_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=Q & "On Hold" & Q)
        .Font.Color = S_HOLD_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=Q & "Waiting" & Q)
        .Font.Color = S_WAIT_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Days Overdue col J
    With wsMM.Range("J23:J300").FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND($J23<>"""",$J23>=8)")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_RED
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("J23:J300").FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND($J23<>"""",$J23>=1,$J23<=7)")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_AMB
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With

    ' Action rows amber — AFTER priority so priority badges show on top
    With wsMM.Range("A23:K300").FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=$C23=TRUE")
        .Interior.Pattern = xlSolid: .Interior.Color = ACT_BG
        .Font.Color = ACT_FG: .StopIfTrue = False
    End With

    ' Overdue action rows red — LAST so it overrides amber
    With wsMM.Range("A23:K300").FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND($C23=TRUE,NOT(OR($H23=" & Q & "Done" & Q & ",$H23=" & Q & "On Hold" & Q & ",$H23=" & Q & "Waiting" & Q & ")),$G23<>"""",$G23<TODAY())")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_BG
        .Font.Color = OV_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' -- GRIDLINES -------------------------------------------------
    wsMM.Activate
    ActiveWindow.DisplayGridlines = False

    Application.ScreenUpdating = True
    MsgBox "MM formatting and CF done!", vbInformation, "FormatMM Done"
End Sub
Sub AlternateRowShading()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    Application.ScreenUpdating = False

    Dim WHITE As Long: WHITE = RGB(255, 255, 255)
    Dim GREY  As Long: GREY = RGB(242, 242, 242)
    Dim NAVY  As Long: NAVY = RGB(31, 56, 100)
    Dim BLUE  As Long: BLUE = RGB(46, 114, 182)

    Dim r         As Long
    Dim toggle    As Boolean
    Dim cellVal   As String
    Dim isHeader  As Boolean

    toggle = True  ' Start with white

    For r = 23 To 300
        Dim rng As Range
        Set rng = wsMM.Range("A" & r & ":K" & r)

        ' Check if this row is a section or subsection header
        cellVal = Trim(CStr(wsMM.Cells(r, 1).Value))
        isHeader = False

        ' Main section header — starts with a number like "1.0", "2.0" etc
        If Len(cellVal) >= 3 Then
            If IsNumeric(Left(cellVal, 1)) And Mid(cellVal, 2, 1) = "." Then
                isHeader = True
            End If
        End If

        ' Sub-header — contains | character
        If InStr(cellVal, "|") > 0 Then isHeader = True

        ' Also check cols B-K for | in case col A is blank but row is a sub-header
        If Not isHeader Then
            Dim bVal As String
            bVal = Trim(CStr(wsMM.Cells(r, 2).Value))
            If InStr(bVal, "|") > 0 Then isHeader = True
        End If

        If isHeader Then
            ' Apply navy or blue depending on type
            If InStr(cellVal, "|") > 0 Then
                rng.Interior.Color = BLUE
                rng.Font.Color = WHITE
                rng.Font.bold = True
            Else
                rng.Interior.Color = NAVY
                rng.Font.Color = WHITE
                rng.Font.bold = True
            End If
            ' Reset toggle after each header so data rows restart fresh
            toggle = True
        Else
            ' Data row — apply alternating shade
            If toggle Then
                rng.Interior.Color = WHITE
            Else
                rng.Interior.Color = GREY
            End If
            rng.Font.Color = RGB(26, 26, 46)
            toggle = Not toggle
        End If
    Next r

    Application.ScreenUpdating = True
    MsgBox "Alternating row shading applied." & Chr(10) & Chr(10) & _
        "CF rules will override amber/pink on action and overdue rows.", _
        vbInformation, "AlternateRowShading Done"
End Sub

Sub FixPriorityStatusCF()
    Dim wsMM As Worksheet
    Dim wsAR As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    Set wsAR = ThisWorkbook.Sheets("Action Register")

    Dim WHITE    As Long: WHITE = RGB(255, 255, 255)
    Dim P_CRIT   As Long: P_CRIT = RGB(192, 0, 0)
    Dim P_HIGH   As Long: P_HIGH = RGB(237, 125, 49)
    Dim P_MED    As Long: P_MED = RGB(255, 192, 0)
    Dim P_LOW    As Long: P_LOW = RGB(169, 209, 142)
    Dim P_MED_FG As Long: P_MED_FG = RGB(125, 82, 0)
    Dim P_LOW_FG As Long: P_LOW_FG = RGB(55, 86, 35)
    Dim S_OPEN   As Long: S_OPEN = RGB(89, 89, 89)
    Dim S_IP     As Long: S_IP = RGB(125, 82, 0)
    Dim S_DONE   As Long: S_DONE = RGB(55, 86, 35)
    Dim S_HOLD   As Long: S_HOLD = RGB(31, 56, 100)
    Dim S_WAIT   As Long: S_WAIT = RGB(132, 60, 12)

    ' -- MEETING MINUTES ------------------------------------------
    ' Clear only Priority (col I) and Status (col H) CF rules
    wsMM.Range("I23:I300").FormatConditions.Delete
    wsMM.Range("H23:H300").FormatConditions.Delete

    ' Priority col I — no quotes around value, Excel adds its own
    With wsMM.Range("I23:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Critical")
        .Interior.Pattern = xlSolid: .Interior.Color = P_CRIT
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("I23:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="High")
        .Interior.Pattern = xlSolid: .Interior.Color = P_HIGH
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("I23:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Medium")
        .Interior.Pattern = xlSolid: .Interior.Color = P_MED
        .Font.Color = P_MED_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("I23:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Low")
        .Interior.Pattern = xlSolid: .Interior.Color = P_LOW
        .Font.Color = P_LOW_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Status col H — font only, no quotes
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Open")
        .Font.Color = S_OPEN: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="In Progress")
        .Font.Color = S_IP: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Done")
        .Font.Color = S_DONE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="On Hold")
        .Font.Color = S_HOLD: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H23:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Waiting")
        .Font.Color = S_WAIT: .Font.bold = True: .StopIfTrue = False
    End With

    ' -- ACTION REGISTER ------------------------------------------
    wsAR.Range("E4:E103").FormatConditions.Delete
    wsAR.Range("F4:F103").FormatConditions.Delete

    ' Priority col E
    With wsAR.Range("E4:E103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Critical")
        .Interior.Pattern = xlSolid: .Interior.Color = P_CRIT
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("E4:E103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="High")
        .Interior.Pattern = xlSolid: .Interior.Color = P_HIGH
        .Font.Color = WHITE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("E4:E103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Medium")
        .Interior.Pattern = xlSolid: .Interior.Color = P_MED
        .Font.Color = P_MED_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("E4:E103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Low")
        .Interior.Pattern = xlSolid: .Interior.Color = P_LOW
        .Font.Color = P_LOW_FG: .Font.bold = True: .StopIfTrue = False
    End With

    ' Status col F — font only
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Open")
        .Font.Color = S_OPEN: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="In Progress")
        .Font.Color = S_IP: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Done")
        .Font.Color = S_DONE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="On Hold")
        .Font.Color = S_HOLD: .Font.bold = True: .StopIfTrue = False
    End With
    With wsAR.Range("F4:F103").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Waiting")
        .Font.Color = S_WAIT: .Font.bold = True: .StopIfTrue = False
    End With

    MsgBox "Priority and Status CF fixed on both MM and AR.", vbInformation
End Sub
Sub RebuildMMHeader()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim NAVY    As Long: NAVY = RGB(31, 56, 100)
    Dim BLUE    As Long: BLUE = RGB(46, 114, 182)
    Dim GREEN   As Long: GREEN = RGB(112, 173, 71)
    Dim WHITE   As Long: WHITE = RGB(255, 255, 255)
    Dim LBLUE   As Long: LBLUE = RGB(122, 156, 200)
    Dim SHBG    As Long: SHBG = RGB(238, 244, 251)
    Dim SHBD    As Long: SHBD = RGB(46, 114, 182)
    Dim FAFAFA  As Long: FAFAFA = RGB(250, 250, 250)
    Dim F2F2F2  As Long: F2F2F2 = RGB(242, 242, 242)
    Dim DARK    As Long: DARK = RGB(26, 26, 46)
    Dim NAVY_FG As Long: NAVY_FG = RGB(31, 56, 100)
    Dim GREEN2  As Long: GREEN2 = RGB(112, 173, 71)

    ' -- STEP 1: Clear rows 1-20 -----------------------------------
    wsMM.Rows("1:20").UnMerge
    wsMM.Rows("1:20").ClearContents
    wsMM.Rows("1:20").ClearFormats
    wsMM.Rows("1:20").rowHeight = 15

    ' -- STEP 2: ROW 1 — Banner -----------------------------------
    wsMM.Range("A1:J1").Merge
    With wsMM.Range("A1:J1")
        .Interior.Color = NAVY
        .Font.Color = WHITE
        .Font.Name = "Arial": .Font.Size = 13: .Font.bold = True
        .Value = "     Engineering Team  —  Internal Meeting Minutes"
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    wsMM.Rows(1).rowHeight = 36

    With wsMM.Range("A1:J1").Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Color = GREEN: .Weight = xlThick
    End With
    With wsMM.Range("A1:J1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = GREEN: .Weight = xlMedium
    End With

    ' Shortcuts panel K1:K2
    wsMM.Range("K1:K2").Merge
    With wsMM.Range("K1:K2")
        .Interior.Color = SHBG
        .Font.Name = "Arial": .Font.Size = 8: .Font.Color = NAVY_FG
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlTop: .WrapText = True
        .Value = "SHORTCUTS" & Chr(10) & _
                 "Ctrl+R  Add row" & Chr(10) & _
                 "Ctrl+Shift+R  Add section" & Chr(10) & _
                 "Refresh  Update register" & Chr(10) & _
                 "Hide/Show  Toggle done" & Chr(10) & _
                 "Full guide: How To Use tab"
    End With
    With wsMM.Range("K1:K2").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Color = SHBD: .Weight = xlMedium
    End With

    ' -- STEP 3: ROW 2 — Info row ---------------------------------
    wsMM.Rows(2).rowHeight = 16

    With wsMM.Cells(2, 1)
        .Interior.Color = FAFAFA
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
        .Font.Color = DARK: .Value = "Date"
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    With wsMM.Cells(2, 2)
        .Interior.Color = FAFAFA
        .Font.Name = "Arial": .Font.Size = 9: .Font.Color = DARK
        .Value = "": .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .NumberFormat = "DD/MM/YYYY"
    End With

    wsMM.Range("C2:D2").Merge
    With wsMM.Range("C2:D2")
        .Interior.Color = FAFAFA
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
        .Font.Color = DARK: .Value = "Location"
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With

    wsMM.Range("E2:G2").Merge
    With wsMM.Range("E2:G2")
        .Interior.Color = FAFAFA
        .Font.Name = "Arial": .Font.Size = 9: .Font.Color = DARK
        .Value = "": .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With

    With wsMM.Cells(2, 8)
        .Interior.Color = FAFAFA
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
        .Font.Color = DARK: .Value = "Meeting No."
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    With wsMM.Cells(2, 9)
        .Interior.Color = FAFAFA
        .Font.Name = "Arial": .Font.Size = 9: .Font.Color = DARK
        .Value = "": .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    With wsMM.Cells(2, 10)
        .Interior.Color = FAFAFA
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
        .Font.Color = DARK: .Value = "Minutes by"
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With

    ' -- STEP 4: ROW 3 — Attendees header -------------------------
    wsMM.Range("A3:K3").Merge
    With wsMM.Range("A3:K3")
        .Interior.Color = NAVY: .Font.Color = WHITE
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
        .Value = "ATTENDEES"
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    wsMM.Rows(3).rowHeight = 16

    ' -- STEP 5: ROW 4 — Attendees col headers --------------------
    wsMM.Range("A4:J4").Merge
    With wsMM.Range("A4:J4")
        .Interior.Color = BLUE: .Font.Color = WHITE
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
        .Value = "Name — Role"
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    With wsMM.Cells(4, 11)
        .Interior.Color = BLUE: .Font.Color = WHITE
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
        .Value = "Present?": .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    wsMM.Rows(4).rowHeight = 15

    ' -- STEP 6: ROWS 5-11 — Attendee rows ------------------------
    Dim attendees(1 To 7, 1 To 2) As String
    attendees(1, 1) = "Lead Johnson": attendees(1, 2) = "Principal, Project Delivery"
    attendees(2, 1) = "Rich Mielke": attendees(2, 2) = "Senior Civil Designer"
    attendees(3, 1) = "Leah Riung-Kalliosaari": attendees(3, 2) = "Project Manager"
    attendees(4, 1) = "Adina Lieblich": attendees(4, 2) = "Project Manager"
    attendees(5, 1) = "Chris Martin": attendees(5, 2) = "Senior Civil Designer"
    attendees(6, 1) = "Dylan De Bruyn": attendees(6, 2) = "Assistant Project Engineer"
    attendees(7, 1) = "Trisha Sharma": attendees(7, 2) = "Intern"

    Dim i As Long
    For i = 1 To 7
        Dim r As Long: r = 4 + i
        Dim rowBG As Long
        If i Mod 2 = 1 Then rowBG = FAFAFA Else rowBG = F2F2F2
        wsMM.Range("A" & r & ":J" & r).Merge
        With wsMM.Range("A" & r & ":J" & r)
            .Interior.Color = rowBG: .Font.Name = "Arial": .Font.Size = 9
            .Font.Color = DARK
            .Value = attendees(i, 1) & "  —  " & attendees(i, 2)
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
        With wsMM.Cells(r, 11)
            .Interior.Color = rowBG: .Font.Name = "Arial": .Font.Size = 9
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
        wsMM.Rows(r).rowHeight = 15
    Next i

    ' -- STEP 7: ROWS 12-13 — Spacers -----------------------------
    wsMM.Rows(12).rowHeight = 6
    wsMM.Range("A12:K12").Interior.Color = FAFAFA
    wsMM.Rows(13).rowHeight = 6
    wsMM.Range("A13:K13").Interior.ColorIndex = xlNone

    ' -- STEP 8: ROW 14 — Column headers --------------------------
    Dim headers As Variant
    headers = Array("Section / Project", "Description / Notes", "ACTION?", "Risk?", "Decision?", "Owner", "Due Date", "Status", "Priority", "Days Overdue", "Notes / Update")
    Dim c As Long
    For c = 0 To 10
        With wsMM.Cells(14, c + 1)
            .Interior.Color = NAVY: .Font.Color = WHITE
            .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
            .Value = headers(c)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        End With
    Next c
    wsMM.Cells(14, 2).HorizontalAlignment = xlLeft
    wsMM.Rows(14).rowHeight = 28

    ' -- STEP 9: Days Overdue formula from row 15 -----------------
 Dim jCell As Range
For Each jCell In wsMM.Range("J15:J300")
    If Not jCell.MergeCells Then jCell.ClearContents
Next jCell
    wsMM.Range("J15:J300").Formula = _
        "=IF(OR(H15=""Done"",H15=""On Hold"",H15=""Waiting""),"""",IF(NOT(C15),"""",IF(G15="""","""",IF(TODAY()>G15,TODAY()-G15,""""))))"
    wsMM.Range("J15:J300").NumberFormat = "0"
    wsMM.Range("J15:J300").HorizontalAlignment = xlCenter

    ' -- STEP 10: Rebuild all CF from row 15 ----------------------
    wsMM.Cells.FormatConditions.Delete

    Dim WHITE2   As Long: WHITE2 = RGB(255, 255, 255)
    Dim P_CRIT   As Long: P_CRIT = RGB(192, 0, 0)
    Dim P_HIGH   As Long: P_HIGH = RGB(237, 125, 49)
    Dim P_MED    As Long: P_MED = RGB(255, 192, 0)
    Dim P_LOW    As Long: P_LOW = RGB(169, 209, 142)
    Dim P_MED_FG As Long: P_MED_FG = RGB(125, 82, 0)
    Dim P_LOW_FG As Long: P_LOW_FG = RGB(55, 86, 35)
    Dim S_OPEN   As Long: S_OPEN = RGB(89, 89, 89)
    Dim S_IP     As Long: S_IP = RGB(125, 82, 0)
    Dim S_DONE   As Long: S_DONE = RGB(55, 86, 35)
    Dim S_HOLD   As Long: S_HOLD = RGB(31, 56, 100)
    Dim S_WAIT   As Long: S_WAIT = RGB(132, 60, 12)
    Dim ACT_BG   As Long: ACT_BG = RGB(255, 243, 205)
    Dim ACT_FG   As Long: ACT_FG = RGB(125, 82, 0)
    Dim OV_BG    As Long: OV_BG = RGB(255, 199, 206)
    Dim OV_FG    As Long: OV_FG = RGB(156, 0, 6)
    Dim OV_RED   As Long: OV_RED = RGB(192, 0, 0)
    Dim OV_AMB   As Long: OV_AMB = RGB(237, 125, 49)

    With wsMM.Range("I15:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Critical")
        .Interior.Pattern = xlSolid: .Interior.Color = P_CRIT
        .Font.Color = WHITE2: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("I15:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="High")
        .Interior.Pattern = xlSolid: .Interior.Color = P_HIGH
        .Font.Color = WHITE2: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("I15:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Medium")
        .Interior.Pattern = xlSolid: .Interior.Color = P_MED
        .Font.Color = P_MED_FG: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("I15:I300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Low")
        .Interior.Pattern = xlSolid: .Interior.Color = P_LOW
        .Font.Color = P_LOW_FG: .Font.bold = True: .StopIfTrue = False
    End With

    With wsMM.Range("H15:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Open")
        .Font.Color = S_OPEN: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H15:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="In Progress")
        .Font.Color = S_IP: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H15:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Done")
        .Font.Color = S_DONE: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H15:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="On Hold")
        .Font.Color = S_HOLD: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("H15:H300").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Waiting")
        .Font.Color = S_WAIT: .Font.bold = True: .StopIfTrue = False
    End With

    With wsMM.Range("J15:J300").FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($J15<>"""",$J15>=8)")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_RED
        .Font.Color = WHITE2: .Font.bold = True: .StopIfTrue = False
    End With
    With wsMM.Range("J15:J300").FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($J15<>"""",$J15>=1,$J15<=7)")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_AMB
        .Font.Color = WHITE2: .Font.bold = True: .StopIfTrue = False
    End With

    With wsMM.Range("A15:K300").FormatConditions.Add(Type:=xlExpression, Formula1:="=$C15=TRUE")
        .Interior.Pattern = xlSolid: .Interior.Color = ACT_BG
        .Font.Color = ACT_FG: .StopIfTrue = False
    End With

    With wsMM.Range("A15:K300").FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND($C15=TRUE,NOT(OR($H15=""Done"",$H15=""On Hold"",$H15=""Waiting"")),$G15<>"""",$G15<TODAY())")
        .Interior.Pattern = xlSolid: .Interior.Color = OV_BG
        .Font.Color = OV_FG: .Font.bold = True: .StopIfTrue = False
    End With

    With wsMM.Range("A5:K11").FormatConditions.Add(Type:=xlExpression, Formula1:="=$K5=TRUE")
        .Interior.Pattern = xlSolid: .Interior.Color = RGB(226, 239, 218)
        .Font.Color = RGB(55, 86, 35): .Font.bold = True: .StopIfTrue = False
    End With

    ' -- STEP 11: Gridlines off ------------------------------------
    wsMM.Activate
    ActiveWindow.DisplayGridlines = False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Header rebuilt!" & Chr(10) & Chr(10) & _
        "Row 1 = Banner | Row 2 = Info + Shortcuts" & Chr(10) & _
        "Rows 3-11 = Attendees | Row 14 = Column headers" & Chr(10) & _
        "Row 15+ = Data" & Chr(10) & Chr(10) & _
        "Next steps:" & Chr(10) & _
        "1. Insert your logo into row 1 left side" & Chr(10) & _
        "2. Add checkboxes to col K rows 5-11 for attendance" & Chr(10) & _
        "3. Update dropdown validations to start at row 15" & Chr(10) & _
        "4. Run AlternateRowShading for data rows", _
        vbInformation, "RebuildMMHeader Done"
End Sub

Sub DeleteMMCheckboxes()
    Dim wsMM As Worksheet
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    
    Dim shp As Shape
    For Each shp In wsMM.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlCheckBox Then
                shp.Delete
            End If
        End If
    Next shp
    
    MsgBox "All checkboxes deleted from Meeting Minutes.", vbInformation
End Sub
Sub AddBannerText()
    Dim ws As Worksheet
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    Set cell = ws.Range("A1")
    
    ' Write both lines with a line break between them
    cell.Value = "CITY OF NEDLANDS" & Chr(10) & "Engineering Team — Internal Meeting Minutes"
    
    ' Wrap text must be on for the line break to show
    cell.WrapText = True
    
    ' Format "CITY OF NEDLANDS" (characters 1–17)
    With cell.Characters(1, 17).Font
        .Name = "Arial"
        .Size = 7
        .bold = True
        .Color = RGB(122, 156, 200)   ' muted blue-grey
    End With
    
    ' Format "Engineering Team — Internal Meeting Minutes" (characters 19 onwards)
    ' 19 = 17 chars + Chr(10) + first char of title
    With cell.Characters(19, 100).Font
        .Name = "Arial"
        .Size = 13
        .bold = True
        .Color = RGB(255, 255, 255)   ' white
    End With
    
    ' Vertical alignment — bottom so both lines sit naturally
    cell.VerticalAlignment = xlVAlignCenter
    
End Sub
Sub AddTopStripe()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    With ws.Range("A1:K1").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 173, 71)
    End With
End Sub

Sub AddBottomStripe()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    With ws.Range("A1:K1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 173, 71)
    End With
End Sub
Sub AddShortcutPanelBorders()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    
    ' Left edge of entire panel I1:I5 — blue medium
    With ws.Range("I1:I5").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(46, 114, 182)
    End With
    
    ' Bottom of row 1 (green stripe to match banner)
    With ws.Range("I1:K1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 173, 71)
    End With
    
    ' Bottom of row 5 (close off the panel)
    With ws.Range("I5:K5").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(46, 114, 182)
    End With
    
    ' Internal horizontal borders between rows 1-5
    With ws.Range("I1:K5").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 218, 240)
    End With

End Sub

Sub ToggleDoneRows()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim hidingDone As Boolean
    
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Check current state by looking for any visible Done row
    ' If we find one, we need to hide. If all Done rows already hidden, we show.
    hidingDone = False
    For i = 16 To lastRow
        If ws.Cells(i, 8).Value = "Done" And Not ws.Rows(i).Hidden Then
            hidingDone = True
            Exit For
        End If
    Next i
    
    ' Now act on every row from row 16 down
    Application.ScreenUpdating = False
    For i = 16 To lastRow
        If ws.Cells(i, 8).Value = "Done" Then
            ws.Rows(i).Hidden = hidingDone
        End If
    Next i
    Application.ScreenUpdating = True
    
    ' Update button text to reflect new state
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.OnAction = "ToggleDoneRows" Then
            If hidingDone Then
                shp.TextFrame.Characters.Text = "Show Done"
            Else
                shp.TextFrame.Characters.Text = "Hide Done"
            End If
        End If
    Next shp

End Sub
Sub AddAttendeeCheckboxes()
    Dim ws As Worksheet
    Dim chk As CheckBox
    Dim cell As Range
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    
    ' Remove any existing checkboxes in that area first to avoid duplicates
    For Each chk In ws.CheckBoxes
        If chk.TopLeftCell.Row >= 8 And chk.TopLeftCell.Row <= 14 And _
           chk.TopLeftCell.Column = 3 Then
            chk.Delete
        End If
    Next chk
    
    ' Add a checkbox to each attendee row
    For i = 8 To 14
        Set cell = ws.Cells(i, 3)
        
        Set chk = ws.CheckBoxes.Add( _
            cell.Left, _
            cell.Top, _
            cell.Width, _
            cell.Height)
        
        With chk
            .Caption = ""
            .LinkedCell = cell.Address
            .Name = "chkAttendee_" & i
        End With
    Next i
    
End Sub
Sub HideCheckboxText()
    Dim ws As Worksheet
    Dim i As Long
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    
    For i = 8 To 14
        Set cell = ws.Cells(i, 3)
        ' Set font colour to match the cell fill colour
        cell.Font.Color = cell.Interior.Color
    Next i
    
End Sub
Sub BuildRiskRegisterTab()
    Dim ws As Worksheet
    Dim shp As Shape
    
    ' Delete existing tab if it exists to start fresh
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Risk Register").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new sheet and position it after Action Register
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Action Register"))
    ws.Name = "Risk Register"
    
    ' -- Hide gridlines --
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    
    ' -- Column widths --
    ws.Columns("A").ColumnWidth = 8      ' Risk ID
    ws.Columns("B").ColumnWidth = 18     ' Project
    ws.Columns("C").ColumnWidth = 12     ' Date Identified
    ws.Columns("D").ColumnWidth = 40     ' Description
    ws.Columns("E").ColumnWidth = 13     ' Likelihood
    ws.Columns("F").ColumnWidth = 13     ' Risk Rating
    ws.Columns("G").ColumnWidth = 10     ' Owner
    ws.Columns("H").ColumnWidth = 38     ' Mitigation / Action
    ws.Columns("I").ColumnWidth = 13     ' Status
    ws.Columns("J").ColumnWidth = 25     ' Notes
    
    ' -- Row heights --
    ws.Rows(1).rowHeight = 28    ' Banner
    ws.Rows(2).rowHeight = 15    ' Subtitle
    ws.Rows(3).rowHeight = 28    ' Column headers
    
    ' -- Row 1: Banner --
    With ws.Range("A1:J1")
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .Font.Size = 13
        .Font.Name = "Arial"
        .Value = "Engineering Team — Risk Register"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
        .IndentLevel = 1
    End With
    
    ' Green bottom border on banner
    With ws.Range("A1:J1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 173, 71)
    End With
    
    ' -- Row 2: Subtitle --
    With ws.Range("A2:J2")
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(238, 244, 251)
        .Font.Color = RGB(31, 56, 100)
        .Font.Size = 8
        .Font.Italic = True
        .Font.Name = "Arial"
        .Value = "Risks flagged from Meeting Minutes where RISK?=TRUE. Complete Likelihood, Rating, Owner and Mitigation within 24hrs. Reviewed by Lead at next meeting."
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
        .IndentLevel = 1
    End With
    
    ' -- Row 3: Column headers --
    Dim headers As Variant
    headers = Array("Risk ID", "Project", "Date Identified", "Description", _
                    "Likelihood", "Risk Rating", "Owner", "Mitigation / Action", _
                    "Status", "Notes")
    
    Dim i As Integer
    For i = 0 To 9
        With ws.Cells(3, i + 1)
            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(31, 56, 100)
            .Font.Color = RGB(255, 255, 255)
            .Font.bold = True
            .Font.Size = 9
            .Font.Name = "Arial"
            .Value = headers(i)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
        End With
    Next i
    ' Description and Mitigation left-aligned
    ws.Cells(3, 4).HorizontalAlignment = xlLeft
    ws.Cells(3, 8).HorizontalAlignment = xlLeft
    
    ' -- Data rows 4–103: alternating fills --
    Dim r As Long
    For r = 4 To 103
        With ws.Rows(r)
            .rowHeight = 15
        End With
        If r Mod 2 = 0 Then
            With ws.Range("A" & r & ":J" & r)
                .Interior.Pattern = xlSolid
                .Interior.Color = RGB(252, 228, 214)
            End With
        Else
            With ws.Range("A" & r & ":J" & r)
                .Interior.Pattern = xlSolid
                .Interior.Color = RGB(245, 208, 188)
            End With
        End If
    Next r
    
    ' -- Date format on col C --
    ws.Columns("C").NumberFormat = "DD/MM/YYYY"
    
    ' -- Dropdown: Likelihood col E --
    With ws.Range("E4:E103").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="High,Medium,Low"
        .ShowError = False
    End With
    
    ' -- Dropdown: Risk Rating col F --
    With ws.Range("F4:F103").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="High,Medium,Low"
        .ShowError = False
    End With
    
    ' -- Dropdown: Status col I --
    With ws.Range("I4:I103").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Open,Monitored,Closed"
        .ShowError = False
    End With
    
    ' -- Alignment for data rows --
    ws.Range("A4:A103").HorizontalAlignment = xlCenter  ' Risk ID
    ws.Range("C4:C103").HorizontalAlignment = xlCenter  ' Date
    ws.Range("E4:E103").HorizontalAlignment = xlCenter  ' Likelihood
    ws.Range("F4:F103").HorizontalAlignment = xlCenter  ' Risk Rating
    ws.Range("G4:G103").HorizontalAlignment = xlCenter  ' Owner
    ws.Range("I4:I103").HorizontalAlignment = xlCenter  ' Status
    
    ' -- "Populate from Minutes" button shape --
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("B1").Left, _
        ws.Range("B1").Top + 4, _
        180, 20)
    
    With shp
        .Name = "btnPopulateRisk"
        .Fill.ForeColor.RGB = RGB(46, 114, 182)
        .Line.Visible = msoFalse
        With .TextFrame
            .Characters.Text = "Populate from Minutes"
            .Characters.Font.Name = "Arial"
            .Characters.Font.Size = 9
            .Characters.Font.bold = True
            .Characters.Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    End With
    
    MsgBox "Risk Register tab built successfully.", vbInformation

End Sub

Sub PopulateRiskRegister()
    Dim wsMin As Worksheet
    Dim wsRR As Worksheet
    Dim r As Long
    Dim rrRow As Long
    Dim lookupKey As String
    Dim existingKeys As Object
    Dim nextID As Long
    
    Set wsMin = ThisWorkbook.Sheets("Meeting Minutes")
    Set wsRR = ThisWorkbook.Sheets("Risk Register")
    Set existingKeys = CreateObject("Scripting.Dictionary")
    
    ' -- Step 1: Read existing risks into dictionary (Description | Project) --
    ' Also find the highest Risk ID number already assigned
    nextID = 0
    Dim existRow As Long
    For existRow = 4 To 103
        If wsRR.Cells(existRow, 2).Value <> "" Then
            lookupKey = Trim(wsRR.Cells(existRow, 4).Value) & "|" & Trim(wsRR.Cells(existRow, 2).Value)
            existingKeys(lookupKey) = existRow
            ' Track highest ID number
            Dim idVal As String
            idVal = Trim(wsRR.Cells(existRow, 1).Value)
            If Left(idVal, 2) = "R-" Then
                Dim idNum As Long
                idNum = CLng(Mid(idVal, 3))
                If idNum > nextID Then nextID = idNum
            End If
        End If
    Next existRow
    nextID = nextID + 1
    
    ' -- Step 2: Find first empty row in RR --
    rrRow = 4
    Do While wsRR.Cells(rrRow, 2).Value <> "" And rrRow <= 103
        rrRow = rrRow + 1
    Loop
    
    ' -- Step 3: Scan MM for flagged risks --
    Dim newCount As Long
    newCount = 0
    
    For r = 18 To 300
        ' Skip merged rows (section/subsection headers)
        If wsMin.Cells(r, 1).MergeCells Then GoTo NextRow
        ' Skip if RISK? not ticked (col D)
        If wsMin.Cells(r, 4).Value <> True Then GoTo NextRow
        ' Skip if no description
        If wsMin.Cells(r, 2).Value = "" Then GoTo NextRow
        
        ' Build lookup key — Description | Project
        lookupKey = Trim(wsMin.Cells(r, 2).Value) & "|" & Trim(wsMin.Cells(r, 12).Value)
        
        ' Skip if already exists in register
        If existingKeys.Exists(lookupKey) Then GoTo NextRow
        
        ' -- Write new risk row --
        ' Col A: Risk ID
        wsRR.Cells(rrRow, 1).Value = "R-" & Format(nextID, "000")
        
        ' Col B: Project (from MM col L)
        wsRR.Cells(rrRow, 2).Value = wsMin.Cells(r, 12).Value
        
        ' Col C: Date Identified — today
        wsRR.Cells(rrRow, 3).Value = Date
        wsRR.Cells(rrRow, 3).NumberFormat = "DD/MM/YYYY"
        
        ' Col D: Description (from MM col B)
        wsRR.Cells(rrRow, 4).Value = wsMin.Cells(r, 2).Value
        
        ' Col E: Likelihood — blank, PM fills
        ' Col F: Risk Rating — blank, PM fills
        
        ' Col G: Owner (from MM col F)
        wsRR.Cells(rrRow, 7).Value = wsMin.Cells(r, 6).Value
        
        ' Col H: Mitigation — blank, PM fills
        
        ' Col I: Status — default to Open
        wsRR.Cells(rrRow, 9).Value = "Open"
        
        ' Col J: Notes (from MM col K)
        wsRR.Cells(rrRow, 10).Value = wsMin.Cells(r, 11).Value
        
        ' Row height
        wsRR.Rows(rrRow).rowHeight = 15
        
        nextID = nextID + 1
        rrRow = rrRow + 1
        newCount = newCount + 1
        
        If rrRow > 103 Then Exit For
        
NextRow:
    Next r
    
    MsgBox newCount & " new risk(s) added to the Risk Register.", vbInformation

End Sub
Sub ToggleClosedRisks()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim hidingClosed As Boolean
    Set ws = ThisWorkbook.Sheets("Risk Register")
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    hidingClosed = False
    For i = 4 To lastRow
        If ws.Cells(i, 9).Value = "Closed" And Not ws.Rows(i).Hidden Then
            hidingClosed = True
            Exit For
        End If
    Next i
    Application.ScreenUpdating = False
    For i = 4 To lastRow
        If ws.Cells(i, 9).Value = "Closed" Then
            ws.Rows(i).Hidden = hidingClosed
        End If
    Next i
    Application.ScreenUpdating = True
    Dim shp As Shape
    Dim onAct As String
    For Each shp In ws.Shapes
        On Error Resume Next
        onAct = shp.OnAction
        On Error GoTo 0
        If onAct = "ToggleClosedRisks" Then
            If hidingClosed Then
                shp.TextFrame.Characters.Text = "Show Closed"
            Else
                shp.TextFrame.Characters.Text = "Hide / Show Closed"
            End If
        End If
    Next shp
End Sub
Sub AddRiskRegisterHideButton()
    Dim ws As Worksheet
    Dim shp As Shape
    Set ws = ThisWorkbook.Sheets("Risk Register")
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("D1").Left, ws.Range("D1").Top + 4, 180, 20)
    With shp
        .Name = "btnHideClosedRisks"
        .Fill.ForeColor.RGB = RGB(46, 114, 182)
        .Line.Visible = msoFalse
        With .TextFrame
            .Characters.Text = "Hide / Show Closed"
            .Characters.Font.Name = "Arial"
            .Characters.Font.Size = 9
            .Characters.Font.bold = True
            .Characters.Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    End With
    MsgBox "Hide / Show Closed button added.", vbInformation
End Sub

Sub BuildHowToUseTab()
    Dim ws As Worksheet
    Dim r As Long

    ' -- Delete and recreate the sheet --
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("How To Use").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "How To Use"
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' -- Column widths --
    ws.Columns("A").ColumnWidth = 4       ' Left margin
    ws.Columns("B").ColumnWidth = 22      ' Label
    ws.Columns("C").ColumnWidth = 75      ' Content
    ws.Columns("D").ColumnWidth = 4       ' Right margin

    ' --------------------------------------------------
    ' HELPER: reusable formatting subs defined inline
    ' --------------------------------------------------

    ' -- Row 1: Banner --
    ws.Rows(1).rowHeight = 36
    With ws.Range("A1:D1")
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.bold = True
        .Value = "Engineering Team Meeting Minutes — How To Use"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
        .IndentLevel = 1
    End With
    With ws.Range("A1:D1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 173, 71)
    End With

    ' -- Row 2: Subtitle --
    ws.Rows(2).rowHeight = 18
    With ws.Range("A2:D2")
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(238, 244, 251)
        .Font.Color = RGB(31, 56, 100)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Italic = True
        .Value = "Engineering Team — Engineering Team   |   April 2026   |   Read once, refer back as needed."
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
        .IndentLevel = 1
    End With

    r = 3

    ' --------------------------------------------------
    ' SECTION HEADER helper — writes a navy section header row
    ' --------------------------------------------------
    ' We'll use a local approach — write each section inline

    ' -- SPACER --
    ws.Rows(r).rowHeight = 8
    ws.Range("A" & r & ":D" & r).Interior.Color = RGB(240, 240, 240)
    r = r + 1

    ' --------------
    ' 01  OVERVIEW
    ' --------------
    ws.Rows(r).rowHeight = 22
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.bold = True
        .Value = "  01   OVERVIEW"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim overviewData As Variant
    overviewData = Array( _
        Array("What is this?", "A meeting minutes and action tracking system built entirely in Microsoft Excel. Designed for weekly project team meetings — it captures notes, flags actions and risks, and populates dedicated registers with one click. No plugins. No licences. Works inside SharePoint."), _
        Array("Who is it for?", "Project managers, engineers, and team leads who run regular status meetings and need a simple, reliable system to track what was discussed, what needs to be done, and who is doing it."), _
        Array("How does it work?", "You take minutes during the meeting. Tick ACTION? for anything requiring follow-up. Tick RISK? for anything that needs monitoring. At end of meeting — one click populates the Action Register and Risk Register. The team updates statuses during the week. The Summary tab always shows who has what outstanding."), _
        Array("What does it need?", "Microsoft Excel only. No plugins, no admin rights required. Works within SharePoint and standard corporate IT environments. Always open from the local synced SharePoint folder in File Explorer — not from the browser."))

    Dim i As Integer
    For i = 0 To UBound(overviewData)
        ws.Rows(r).rowHeight = 30
        With ws.Range("B" & r)
            .Value = overviewData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
            .IndentLevel = 1
        End With
        With ws.Range("C" & r)
            .Value = overviewData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        r = r + 1
    Next i

    ' -- SPACER --
    ws.Rows(r).rowHeight = 8
    ws.Range("A" & r & ":D" & r).Interior.Color = RGB(240, 240, 240)
    r = r + 1

    ' ------------------
    ' 02  THE FIVE TABS
    ' ------------------
    ws.Rows(r).rowHeight = 22
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.bold = True
        .Value = "  02   THE FIVE TABS"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim tabsData As Variant
    tabsData = Array( _
        Array("Meeting Minutes", "Where you take notes during the meeting. One row per agenda item. Use section headers (dark navy) and project sub-headers (mid blue) to structure the meeting. Tick ACTION? or RISK? checkboxes as needed."), _
        Array("Action Register", "Auto-populated from Meeting Minutes. Shows all ticked ACTION? rows with owner, due date, priority and status. Team updates statuses here during the week."), _
        Array("Risk Register", "Auto-populated from Meeting Minutes. Shows all ticked RISK? rows. PM completes Likelihood, Risk Rating and Mitigation within 24hrs of the meeting. Reviewed by Lead at the start of each meeting."), _
        Array("Summary", "Live count of actions by team member and status — Total, Open, In Progress, Done, On Hold, Waiting, and Overdue. Updates when Action Register is refreshed. Review at the start of each meeting."), _
        Array("How To Use", "This tab. Read once, refer back as needed."))

    For i = 0 To UBound(tabsData)
        ws.Rows(r).rowHeight = 28
        With ws.Range("B" & r)
            .Value = tabsData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
            .IndentLevel = 1
        End With
        With ws.Range("C" & r)
            .Value = tabsData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        r = r + 1
    Next i

    ' -- SPACER --
    ws.Rows(r).rowHeight = 8
    ws.Range("A" & r & ":D" & r).Interior.Color = RGB(240, 240, 240)
    r = r + 1

    ' ------------------------------------
    ' 03  MEETING MINUTES — COLUMN GUIDE
    ' ------------------------------------
    ws.Rows(r).rowHeight = 22
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.bold = True
        .Value = "  03   MEETING MINUTES — COLUMN GUIDE"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim colData As Variant
    colData = Array( _
        Array("A  —  Section / Project", "The project or agenda section this row relates to. Only fill on the FIRST row of each project block. Leave blank on continuation rows — keeps the minutes clean and readable."), _
        Array("B  —  Description / Notes", "What was discussed, decided, or what needs to be done. One clear point per row. For action rows write as an instruction — e.g. 'Email Western Power re works date'."), _
        Array("C  —  Action?", "Tick if the row requires someone to do something. The row turns amber immediately. Unticked rows are general notes — no follow-up required. Fill in Owner, Due Date and Status for all ticked rows."), _
        Array("D  —  Risk?", "Tick if the item represents a risk that needs to be documented and monitored. At end of meeting, click Populate Risk Register — the row will appear in the Risk Register tab for the PM to complete."), _
        Array("F  —  Owner", "Who is responsible for this action. Select from the dropdown: TM1, TM2, TM3, TM4, TM5, TM6, ALL. Required for all action rows."), _
        Array("G  —  Due Date", "When the action must be completed. Always DD/MM/YYYY — e.g. 07/04/2026. Required for all action rows."), _
        Array("H  —  Status", "Open / In Progress / Done / On Hold / Waiting. Update this as the action progresses. Done rows can be hidden using the Hide Done button."), _
        Array("I  —  Priority", "Critical / High / Medium / Low. Optional but recommended for actions with a due date. Drives the colour badge in the Action Register."), _
        Array("J  —  OD Days", "Auto-calculated. Shows how many days overdue the action is. Blank if not overdue, or if Status is Done / On Hold / Waiting."), _
        Array("K  —  Notes / Update", "Additional context, weekly progress notes, or relevant information. Free text — use to capture updates without changing the status."), _
        Array("L  —  Project (hidden)", "A hidden column storing the project name for each row. Used by the populate macros to group items by project in the registers. Do not unhide or edit."))

    For i = 0 To UBound(colData)
        ws.Rows(r).rowHeight = 28
        With ws.Range("B" & r)
            .Value = colData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
            .IndentLevel = 1
        End With
        With ws.Range("C" & r)
            .Value = colData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        r = r + 1
    Next i

    ' -- SPACER --
    ws.Rows(r).rowHeight = 8
    ws.Range("A" & r & ":D" & r).Interior.Color = RGB(240, 240, 240)
    r = r + 1

    ' ----------------------
    ' 04  COLOUR CODE GUIDE
    ' ----------------------
    ws.Rows(r).rowHeight = 22
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.bold = True
        .Value = "  04   COLOUR CODE GUIDE"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    ' Sub-header: Row colours
    ws.Rows(r).rowHeight = 18
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(46, 114, 182)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.bold = True
        .Value = "  Row colours — Meeting Minutes"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim rowColData As Variant
    rowColData = Array( _
        Array("Dark navy row", "Main section header — e.g. 1.0 TEAM UPDATE / REVIEW. Do not edit these rows directly."), _
        Array("Mid blue row", "Project sub-header — e.g. The Avenue | Budget: $1.95M | PM: TM1. Do not edit these rows directly."), _
        Array("Pink / red row", "ACTION? ticked AND past due date AND status is not Done/Hold/Waiting. Overdue — needs attention immediately."), _
        Array("Amber row", "ACTION? ticked — action required. Fill in Owner, Due Date and Status."), _
        Array("White / grey row", "General note or update — ACTION? not ticked. No follow-up required."))

    For i = 0 To UBound(rowColData)
        ws.Rows(r).rowHeight = 22
        With ws.Range("B" & r)
            .Value = rowColData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignCenter
            .IndentLevel = 1
        End With
        With ws.Range("C" & r)
            .Value = rowColData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignCenter
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        r = r + 1
    Next i

    ' Sub-header: Status colours
    ws.Rows(r).rowHeight = 18
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(46, 114, 182)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.bold = True
        .Value = "  Status values — col H"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim statusData As Variant
    statusData = Array( _
        Array("Open", "Acknowledged — not yet started."), _
        Array("In Progress", "Actively being worked on."), _
        Array("Done", "Completed. Will be hidden when Hide Done is toggled."), _
        Array("On Hold", "Paused — blocked or deferred. Not counted as overdue."), _
        Array("Waiting", "Waiting on another person or team to act. Not counted as overdue."))

    For i = 0 To UBound(statusData)
        ws.Rows(r).rowHeight = 18
        With ws.Range("B" & r)
            .Value = statusData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignCenter
            .IndentLevel = 1
        End With
        With ws.Range("C" & r)
            .Value = statusData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignCenter
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        r = r + 1
    Next i

    ' -- SPACER --
    ws.Rows(r).rowHeight = 8
    ws.Range("A" & r & ":D" & r).Interior.Color = RGB(240, 240, 240)
    r = r + 1

    ' ----------------------
    ' 05  WEEKLY WORKFLOW
    ' ----------------------
    ws.Rows(r).rowHeight = 22
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.bold = True
        .Value = "  05   WEEKLY WORKFLOW"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    ' Sub-header: During the meeting
    ws.Rows(r).rowHeight = 18
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(46, 114, 182)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.bold = True
        .Value = "  Tuesday — During the meeting"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim duringData As Variant
    duringData = Array( _
        Array("Step 1", "Open the file from the synced SharePoint folder in File Explorer — not from the browser."), _
        Array("Step 2", "Update the Date field in the header to today's date. Update the attendees table — tick Present? for each person in the room."), _
        Array("Step 3", "Take minutes row by row. Use Ctrl+Shift+A to add rows, Ctrl+Shift+R for new sections, Ctrl+Shift+S for new project subsections."), _
        Array("Step 4", "For any item requiring action — tick the ACTION? checkbox (col C). Fill in Owner, Due Date, and Status."), _
        Array("Step 5", "For any item that sounds like a risk — tick the RISK? checkbox (col D). Add a brief note in col K. The PM will complete the full risk record after the meeting."), _
        Array("Step 6", "At end of meeting — click Populate from Minutes on the Action Register tab. Then click Populate from Minutes on the Risk Register tab."), _
        Array("Step 7", "Save the file. SharePoint will sync automatically."))

    For i = 0 To UBound(duringData)
        ws.Rows(r).rowHeight = 28
        With ws.Range("B" & r)
            .Value = duringData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
            .IndentLevel = 1
        End With
        With ws.Range("C" & r)
            .Value = duringData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        r = r + 1
    Next i

    ' Sub-header: During the week
    ws.Rows(r).rowHeight = 18
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(46, 114, 182)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.bold = True
        .Value = "  Tuesday to Monday — During the week"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim weekData As Variant
    weekData = Array( _
        Array("Step 1", "Open the file from the synced SharePoint folder in File Explorer."), _
        Array("Step 2", "Go to the Action Register tab. Filter by your name using the Owner filter arrow."), _
        Array("Step 3", "Update the Status of your actions as you complete them. Add notes in the Notes column if useful."), _
        Array("Step 4", "If you identify a risk during the week — add it directly to the Risk Register tab and complete all fields."), _
        Array("Step 5", "Save the file. Changes sync to SharePoint automatically."))

    For i = 0 To UBound(weekData)
        ws.Rows(r).rowHeight = 22
        With ws.Range("B" & r)
            .Value = weekData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
            .IndentLevel = 1
        End With
        With ws.Range("C" & r)
            .Value = weekData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        r = r + 1
    Next i

    ' Sub-header: Before next meeting
    ws.Rows(r).rowHeight = 18
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(46, 114, 182)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.bold = True
        .Value = "  Next Tuesday — Before the meeting"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim beforeData As Variant
    beforeData = Array( _
        Array("Step 1", "Open the file from SharePoint."), _
        Array("Step 2", "Use Hide / Show Done on the Meeting Minutes tab to declutter — hides completed actions."), _
        Array("Step 3", "Review the Summary tab — check who has outstanding or overdue actions before the meeting starts."), _
        Array("Step 4", "Review the Risk Register — check for any risks that need to be discussed at the meeting."), _
        Array("Step 5", "Take this week's minutes continuing from where the last meeting ended."))

    For i = 0 To UBound(beforeData)
        ws.Rows(r).rowHeight = 22
        With ws.Range("B" & r)
            .Value = beforeData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
            .IndentLevel = 1
        End With
        With ws.Range("C" & r)
            .Value = beforeData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        r = r + 1
    Next i

    ' -- SPACER --
    ws.Rows(r).rowHeight = 8
    ws.Range("A" & r & ":D" & r).Interior.Color = RGB(240, 240, 240)
    r = r + 1

    ' ------------------------------------
    ' 06  KEYBOARD SHORTCUTS & BUTTONS
    ' ------------------------------------
    ws.Rows(r).rowHeight = 22
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.bold = True
        .Value = "  06   KEYBOARD SHORTCUTS & BUTTONS"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    ' Sub-header: Keyboard shortcuts
    ws.Rows(r).rowHeight = 18
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(46, 114, 182)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.bold = True
        .Value = "  Keyboard shortcuts — Meeting Minutes tab"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim shortcutData As Variant
    shortcutData = Array( _
        Array("Ctrl + Shift + A", "Add Row — inserts a blank formatted row at the end of the current project subsection. Click any row in the project first."), _
        Array("Ctrl + Shift + R", "Add Section — adds a full new section header, sub-header and blank row. Prompts for section name and project name."), _
        Array("Ctrl + Shift + S", "Add Subsection — adds a new project sub-header and blank row within the current section. Prompts for project name."))

    For i = 0 To UBound(shortcutData)
        ws.Rows(r).rowHeight = 28
        With ws.Range("B" & r)
            .Value = shortcutData(i)(0)
            .Font.Name = "Courier New"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        With ws.Range("C" & r)
            .Value = shortcutData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignCenter
        End With
        ws.Range("A" & r).Interior.Color = RGB(238, 244, 251)
        ws.Range("D" & r).Interior.Color = RGB(238, 244, 251)
        r = r + 1
    Next i

    ' Sub-header: Buttons
    ws.Rows(r).rowHeight = 18
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(46, 114, 182)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.bold = True
        .Value = "  Buttons"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim buttonData As Variant
    buttonData = Array( _
        Array("Populate from Minutes" & Chr(10) & "(Action Register tab)", "Reads all ticked ACTION? rows from Meeting Minutes and writes them to the Action Register. Run at the end of every meeting. Preserves any status or notes already entered — only adds new rows."), _
        Array("Populate from Minutes" & Chr(10) & "(Risk Register tab)", "Reads all ticked RISK? rows from Meeting Minutes and writes them to the Risk Register. Run at the end of every meeting. Only adds new risks — never overwrites existing rows."), _
        Array("Hide / Show Done" & Chr(10) & "(Meeting Minutes tab)", "Toggles Done rows on and off in the Meeting Minutes tab. Click once to hide completed actions, click again to show them."), _
        Array("Hide / Show Done" & Chr(10) & "(Action Register tab)", "Toggles Done rows on and off in the Action Register. Useful for decluttering the view before a meeting."), _
        Array("Hide / Show Closed" & Chr(10) & "(Risk Register tab)", "Toggles Closed risks on and off in the Risk Register."))

    For i = 0 To UBound(buttonData)
        ws.Rows(r).rowHeight = 36
        With ws.Range("B" & r)
            .Value = buttonData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(46, 114, 182)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
        End With
        With ws.Range("C" & r)
            .Value = buttonData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignCenter
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        ' Override B cell colour set above
        ws.Range("B" & r).Interior.Color = RGB(46, 114, 182)
        r = r + 1
    Next i

    ' -- SPACER --
    ws.Rows(r).rowHeight = 8
    ws.Range("A" & r & ":D" & r).Interior.Color = RGB(240, 240, 240)
    r = r + 1

    ' -----------------------
    ' 07  TROUBLESHOOTING
    ' -----------------------
    ws.Rows(r).rowHeight = 22
    With ws.Range("A" & r & ":D" & r)
        .Merge
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.bold = True
        .Value = "  07   TROUBLESHOOTING"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    Dim troubleData As Variant
    troubleData = Array( _
        Array("Row not turning amber", "Check the ACTION? cell — it must contain TRUE (boolean), not the text 'TRUE'. Make sure the checkbox is linked to the correct cell via right-click > Format Control > Cell Link."), _
        Array("Checkbox showing TRUE/FALSE text", "Run the HideTrueFalseText macro. This sets the number format to ;;; which hides the text. Must be re-run if new rows are added below the last formatted row."), _
        Array("Action Register not updating", "Click Populate from Minutes on the Action Register tab — the register does not update automatically. Check that ACTION? cells contain TRUE (boolean) not text."), _
        Array("Risk Register not updating", "Click Populate from Minutes on the Risk Register tab. Check that RISK? checkboxes in col D are properly linked to their cells via Format Control > Cell Link."), _
        Array("New row has no checkbox", "Checkboxes must be added manually to new rows. Go to Developer tab > Insert > Form Controls > Checkbox. Draw inside the cell, delete the caption text, and set the Cell Link to the cell address via Format Control."), _
        Array("Populate picks up wrong rows", "Check that the checkbox is linked to the correct column. ACTION? must link to col C. RISK? must link to col D. If linked to the wrong column the macro will misread it."), _
        Array("Macro security warning", "Click Enable Content in the yellow bar. You may need to do this each time the file is opened from SharePoint."), _
        Array("File saving to wrong location", "Always open the file from the synced SharePoint folder in File Explorer — not from the browser URL. Browser opens save to personal OneDrive by default."), _
        Array("Status colour not showing", "Check Conditional Formatting rules via Home > Conditional Formatting > Manage Rules. CF fills must be set manually in Manage Rules using the hex codes in the style guide — they cannot be set reliably via VBA."))

    For i = 0 To UBound(troubleData)
        ws.Rows(r).rowHeight = 36
        With ws.Range("B" & r)
            .Value = troubleData(i)(0)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
            .WrapText = True
            .IndentLevel = 1
        End With
        With ws.Range("C" & r)
            .Value = troubleData(i)(1)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(26, 26, 46)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
        End With
        If i Mod 2 = 0 Then
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(250, 250, 250)
        Else
            ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
        End If
        r = r + 1
    Next i

    ' -- Final spacer --
    ws.Rows(r).rowHeight = 8
    ws.Range("A" & r & ":D" & r).Interior.Color = RGB(240, 240, 240)

    MsgBox "How To Use tab built successfully.", vbInformation

End Sub
Sub HideDecisionColumn()
    Dim ws As Worksheet
    Dim chk As CheckBox
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    
    ' Hide all checkboxes sitting in col E
    For Each chk In ws.CheckBoxes
        If chk.TopLeftCell.Column = 5 Then
            chk.Visible = False
        End If
    Next chk
    
    ' Hide the column
    ws.Columns("E").Hidden = True
    
    MsgBox "Decision column hidden.", vbInformation
End Sub

Sub ReplaceAttendeeCheckboxesWithDropdown()
    Dim ws As Worksheet
    Dim chk As CheckBox
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    
    ' Delete all existing checkboxes in col C rows 8-14
    For Each chk In ws.CheckBoxes
        If chk.TopLeftCell.Column = 3 And _
           chk.TopLeftCell.Row >= 8 And _
           chk.TopLeftCell.Row <= 14 Then
            chk.Delete
        End If
    Next chk
    
    ' Add Yes/No dropdown to C8:C14
    With ws.Range("C8:C14").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Yes,No"
        .ShowError = False
    End With
    
    ' Default all to No
    For i = 8 To 14
        ws.Cells(i, 3).Value = "No"
        ws.Cells(i, 3).HorizontalAlignment = xlCenter
        ws.Cells(i, 3).Font.Name = "Arial"
        ws.Cells(i, 3).Font.Size = 9
        ' Clear the ;;; number format from old checkbox
        ws.Cells(i, 3).NumberFormat = "General"
    Next i
    
    MsgBox "Attendee dropdowns created.", vbInformation
End Sub
Sub SyncAll()
    Call RefreshProjectDropdown
    Call PushNewActionsToMinutes
    Call SyncStatusBack
    Call PopulateActionRegister
    Call FixMMCFRanges
    Call RebuildARFormatting

    MsgBox "Sync complete.", vbInformation, "SyncAll Done"
End Sub
Sub RefreshProjectDropdown()
    Dim wsMM As Worksheet
    Dim wsAR As Worksheet
    Dim r As Long
    Dim cellVal As String
    Dim projName As String
    Dim seen As Object
    Dim writeRow As Long

    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    Set seen = CreateObject("Scripting.Dictionary")

    wsAR.Range("Z1:Z200").ClearContents

    writeRow = 1
    For r = 15 To 300
        cellVal = Trim(CStr(wsMM.Cells(r, 1).Value))
        If InStr(cellVal, "|") > 0 Then
            projName = Trim(Split(cellVal, "|")(0))
            If projName <> "" And Not seen.Exists(projName) Then
                seen(projName) = True
                wsAR.Cells(writeRow, 26).Value = projName
                writeRow = writeRow + 1
            End If
        End If
    Next r

    If writeRow = 1 Then
        MsgBox "No project sub-headers found in Meeting Minutes.", vbExclamation
        Exit Sub
    End If

    wsAR.Columns("Z").Hidden = True

    wsAR.Range("A4:A103").Validation.Delete
    wsAR.Range("A4:A103").Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="=$Z$1:$Z$" & (writeRow - 1)
End Sub
Sub DiagnoseValidation()
    Dim wsAR As Worksheet
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    
    Debug.Print "Sheet name: [" & wsAR.Name & "]"
    Debug.Print "Col Z row 1: [" & wsAR.Cells(1, 26).Value & "]"
    Debug.Print "Col Z row 2: [" & wsAR.Cells(2, 26).Value & "]"
    Debug.Print "Col Z row 3: [" & wsAR.Cells(3, 26).Value & "]"
    
    ' Check if col A already has validation
    On Error Resume Next
    Dim v As String
    v = wsAR.Range("A4").Validation.Formula1
    Debug.Print "Current A4 validation: [" & v & "]"
    On Error GoTo 0
    
    ' Check if AR is protected
    Debug.Print "Sheet protected: " & wsAR.ProtectContents
End Sub
Sub TestSimpleValidation()
    Dim wsAR As Worksheet
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    
    ' Try the simplest possible validation
    wsAR.Range("A4:A103").Validation.Delete
    wsAR.Range("A4:A103").Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="Apple,Banana,Cherry"
    
    MsgBox "Basic validation worked"
End Sub

Sub TestRangeValidation()
    Dim wsAR As Worksheet
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    
    wsAR.Range("A4:A103").Validation.Delete
    wsAR.Range("A4:A103").Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="=$Z$1:$Z$3"
    
    MsgBox "Range validation worked"
End Sub


Sub PushNewActionsToMinutes()
    Dim wsMM As Worksheet
    Dim wsAR As Worksheet
    Dim arRow As Long
    Dim minRow As Long
    Dim scanRow As Long
    Dim arDesc As String
    Dim arOwner As String
    Dim arProj As String
    Dim arDue As Variant
    Dim mmCellVal As String
    Dim mmProj As String
    Dim insertRow As Long
    Dim foundProject As Boolean
    Dim alreadyInMM As Boolean
    Dim pushed As Long
    Dim notMatched As Long
    Dim FAFAFA As Long: FAFAFA = RGB(250, 250, 250)
    Dim DARK As Long: DARK = RGB(26, 26, 46)
    Dim chk As CheckBox
    Dim cellC As Range
    Dim cellD As Range

    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    pushed = 0
    notMatched = 0

    Application.ScreenUpdating = False

    For arRow = 4 To 103
        arDesc = Trim(CStr(wsAR.Cells(arRow, 2).Value))
        If arDesc = "" Then GoTo NextAR

        arOwner = Trim(CStr(wsAR.Cells(arRow, 3).Value))
        arProj = Trim(CStr(wsAR.Cells(arRow, 1).Value))

        alreadyInMM = False
        For minRow = 15 To 300
            If InStr(Trim(CStr(wsMM.Cells(minRow, 1).Value)), "|") = 0 Then
                If Trim(CStr(wsMM.Cells(minRow, 2).Value)) = arDesc And _
                   Trim(CStr(wsMM.Cells(minRow, 6).Value)) = arOwner Then
                    alreadyInMM = True
                    Exit For
                End If
            End If
        Next minRow
        If alreadyInMM Then GoTo NextAR

        If arProj = "" Then
            notMatched = notMatched + 1
            GoTo NextAR
        End If

        foundProject = False
        insertRow = 0
        For minRow = 15 To 300
            mmCellVal = Trim(CStr(wsMM.Cells(minRow, 1).Value))
            If InStr(mmCellVal, "|") > 0 Then
                mmProj = Trim(Split(mmCellVal, "|")(0))
                If mmProj = arProj Then
                    foundProject = True
                    insertRow = minRow + 1
                    For scanRow = minRow + 1 To 300
                        If InStr(Trim(CStr(wsMM.Cells(scanRow, 1).Value)), "|") > 0 Then Exit For
                        If Trim(CStr(wsMM.Cells(scanRow, 1).Value)) <> "" And _
                           InStr(Trim(CStr(wsMM.Cells(scanRow, 1).Value)), "|") = 0 And _
                           Trim(CStr(wsMM.Cells(scanRow, 2).Value)) = "" Then Exit For
                        insertRow = scanRow + 1
                    Next scanRow
                    Exit For
                End If
            End If
        Next minRow

        If Not foundProject Then
            notMatched = notMatched + 1
            GoTo NextAR
        End If

        wsMM.Rows(insertRow).Insert Shift:=xlDown
        wsMM.Rows(insertRow).ClearContents
        wsMM.Rows(insertRow).ClearFormats
        wsMM.Rows(insertRow).rowHeight = 15

        With wsMM.Range("A" & insertRow & ":K" & insertRow)
            .Font.Name = "Arial": .Font.Size = 9
            .Font.Color = DARK: .Font.bold = False
            .VerticalAlignment = xlCenter
            .Interior.Pattern = xlSolid
            .Interior.Color = FAFAFA
        End With
        wsMM.Range("A" & insertRow).HorizontalAlignment = xlLeft
        wsMM.Range("B" & insertRow).HorizontalAlignment = xlLeft
        wsMM.Range("C" & insertRow & ":J" & insertRow).HorizontalAlignment = xlCenter
        wsMM.Range("K" & insertRow).HorizontalAlignment = xlLeft

        wsMM.Cells(insertRow, 2).Value = arDesc
        wsMM.Cells(insertRow, 6).Value = arOwner
        arDue = wsAR.Cells(arRow, 4).Value
        If arDue <> "" Then
            wsMM.Cells(insertRow, 7).Value = arDue
            wsMM.Cells(insertRow, 7).NumberFormat = "DD/MM/YYYY"
        End If
        wsMM.Cells(insertRow, 8).Value = Trim(CStr(wsAR.Cells(arRow, 6).Value))
        wsMM.Cells(insertRow, 9).Value = Trim(CStr(wsAR.Cells(arRow, 5).Value))
        wsMM.Cells(insertRow, 11).Value = Trim(CStr(wsAR.Cells(arRow, 8).Value))
        wsMM.Cells(insertRow, 12).Value = arProj

        wsMM.Range("C" & insertRow & ":E" & insertRow).NumberFormat = ";;;"

        ' Overdue formula col J
        wsMM.Cells(insertRow, 10).Formula = _
            "=IF(OR(H" & insertRow & "=""Done"",H" & insertRow & "=""On Hold"",H" & insertRow & "=""Waiting""),"""",IF(NOT(C" & insertRow & "),"""",IF(G" & insertRow & "="""","""",IF(TODAY()>G" & insertRow & ",TODAY()-G" & insertRow & ",""""))))"
        wsMM.Cells(insertRow, 10).NumberFormat = "0"
        wsMM.Cells(insertRow, 10).HorizontalAlignment = xlCenter

        ' Action checkbox — write TRUE to linked cell FIRST, then create box
        Set cellC = wsMM.Cells(insertRow, 3)
        cellC.Value = True
        Set chk = wsMM.CheckBoxes.Add(cellC.Left, cellC.Top, cellC.Width, cellC.Height)
        With chk
            .Caption = ""
            .LinkedCell = cellC.Address
            .Name = "chkAction_" & insertRow
            .Display3DShading = False
            .Value = xlOn
        End With

        ' Risk checkbox
        Set cellD = wsMM.Cells(insertRow, 4)
        cellD.Value = False
        Set chk = wsMM.CheckBoxes.Add(cellD.Left, cellD.Top, cellD.Width, cellD.Height)
        With chk
            .Caption = ""
            .LinkedCell = cellD.Address
            .Name = "chkRisk_" & insertRow
            .Display3DShading = False
            .Value = xlOff
        End With

        pushed = pushed + 1

NextAR:
    Next arRow

    Application.ScreenUpdating = True

    MsgBox pushed & " new action(s) pushed to Minutes." & Chr(10) & _
           notMatched & " row(s) skipped (no project or project not found).", _
           vbInformation, "PushNewActionsToMinutes"
End Sub

Sub FixMMCFRanges()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")

    Dim fc As FormatCondition
    Dim i As Long
    Dim f As String

    For i = 1 To ws.Cells.FormatConditions.Count
        Set fc = ws.Cells.FormatConditions(i)
        f = fc.Formula1

        ' Priority badges — col I
        If f = """Critical""" Or f = """High""" Or f = """Medium""" Or f = """Low""" Then
            fc.ModifyAppliesToRange ws.Range("$I$15:$I$300")

        ' Status — col H
        ElseIf f = """Open""" Or f = """In Progress""" Or f = """Done""" Or _
               f = """On Hold""" Or f = """Waiting""" Then
            fc.ModifyAppliesToRange ws.Range("$H$15:$H$300")

        ' Overdue days red — col J, >=8
        ElseIf InStr(f, ">=8") > 0 And InStr(f, "$J") > 0 Then
            fc.ModifyAppliesToRange ws.Range("$J$15:$J$300")

        ' Overdue days amber — col J, 1–7
        ElseIf InStr(f, ">=1") > 0 And InStr(f, "<=7") > 0 And InStr(f, "$J") > 0 Then
            fc.ModifyAppliesToRange ws.Range("$J$15:$J$300")

        ' Action row amber — =$C..=TRUE
        ElseIf f = "=$C15=TRUE" Then
            fc.ModifyAppliesToRange ws.Range("$A$15:$K$300")

        ' Overdue action row red — compound formula starting with AND($C
        ElseIf InStr(f, "AND($C15=TRUE") > 0 Then
            fc.ModifyAppliesToRange ws.Range("$A$15:$K$300")
        End If
    Next i
End Sub
Sub RebuildARFormatting()
    Dim ws As Worksheet
    Dim r As Long
    Dim FAFAFA As Long: FAFAFA = RGB(250, 250, 250)
    Dim F2F2F2 As Long: F2F2F2 = RGB(242, 242, 242)
    Dim DARK As Long: DARK = RGB(26, 26, 46)

    Set ws = ThisWorkbook.Sheets("Action Register")
    Application.ScreenUpdating = False

    ' Dropdowns across full range
    ws.Range("C4:C103").Validation.Delete
    ws.Range("C4:C103").Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Formula1:="TM1,TM2,TM3,TM4,TM5,TM6,ALL"

    ws.Range("E4:E103").Validation.Delete
    ws.Range("E4:E103").Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Formula1:="Critical,High,Medium,Low"

    ws.Range("F4:F103").Validation.Delete
    ws.Range("F4:F103").Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Formula1:="Open,In Progress,Done,On Hold,Waiting"

    ' Per-row fixes for rows with a description
    For r = 4 To 103
        If Trim(CStr(ws.Cells(r, 2).Value)) <> "" Then
            ' Formula
            ws.Cells(r, 9).Formula = _
                "=IF(OR(F" & r & "=""Done"",F" & r & "=""On Hold"",F" & r & "=""Waiting""),"""",IF(D" & r & "="""","""",IF(TODAY()>D" & r & ",TODAY()-D" & r & ","""")))"
            ws.Cells(r, 9).NumberFormat = "0"
            ws.Cells(r, 9).HorizontalAlignment = xlCenter

            ' Date formats
            ws.Cells(r, 4).NumberFormat = "DD/MM/YYYY"
            ws.Cells(r, 7).NumberFormat = "DD/MM/YYYY"

            ' Alignment
            ws.Rows(r).rowHeight = 15
            ws.Range("A" & r).HorizontalAlignment = xlLeft
            ws.Range("B" & r).HorizontalAlignment = xlLeft
            ws.Range("C" & r & ":G" & r).HorizontalAlignment = xlCenter
            ws.Range("H" & r).HorizontalAlignment = xlLeft

            ' Font
            With ws.Range("A" & r & ":I" & r)
                .Font.Name = "Arial"
                .Font.Size = 9
                .Font.Color = DARK
            End With

            ' Date Added default if blank
            If Trim(CStr(ws.Cells(r, 7).Value)) = "" Then
                ws.Cells(r, 7).Value = Date
            End If
        End If
    Next r

    Application.ScreenUpdating = True
End Sub
Sub EndMeeting()
    Dim response As VbMsgBoxResult
    response = MsgBox("END MEETING SYNC" & Chr(10) & Chr(10) & _
                     "This will:" & Chr(10) & _
                     "1. Overwrite matching AR rows with MM values" & Chr(10) & _
                     "2. Add any new MM actions to the AR" & Chr(10) & _
                     "3. Clean up formatting" & Chr(10) & Chr(10) & _
                     "Run this ONCE at the end of the meeting." & Chr(10) & Chr(10) & _
                     "Continue?", vbYesNo + vbExclamation, "End Meeting")
    If response <> vbYes Then Exit Sub

    Dim wsMin As Worksheet
    Dim wsAR As Worksheet
    Dim r As Long
    Dim arRow As Long
    Dim arDesc As String
    Dim arOwner As String
    Dim updated As Long

    Set wsMin = ThisWorkbook.Sheets("Meeting Minutes")
    Set wsAR = ThisWorkbook.Sheets("Action Register")
    updated = 0

    Application.ScreenUpdating = False

    ' Step 1: Overwrite matching AR rows with MM values
    For arRow = 4 To 103
        arDesc = Trim(CStr(wsAR.Cells(arRow, 2).Value))
        If arDesc = "" Then GoTo NextAR
        arOwner = Trim(CStr(wsAR.Cells(arRow, 3).Value))

        For r = 15 To 300
            If InStr(Trim(CStr(wsMin.Cells(r, 1).Value)), "|") > 0 Then GoTo NextMM
            If wsMin.Cells(r, 3).Value <> True Then GoTo NextMM
            If Trim(CStr(wsMin.Cells(r, 2).Value)) = arDesc And _
               Trim(CStr(wsMin.Cells(r, 6).Value)) = arOwner Then
                wsAR.Cells(arRow, 4).Value = wsMin.Cells(r, 7).Value
                wsAR.Cells(arRow, 4).NumberFormat = "DD/MM/YYYY"
                wsAR.Cells(arRow, 5).Value = wsMin.Cells(r, 9).Value
                wsAR.Cells(arRow, 6).Value = wsMin.Cells(r, 8).Value
                wsAR.Cells(arRow, 8).Value = wsMin.Cells(r, 11).Value
                updated = updated + 1
                Exit For
            End If
NextMM:
        Next r
NextAR:
    Next arRow

    ' Step 2: Pull any NEW MM actions into AR
    Call PopulateActionRegister

    ' Step 3: Clean up
    Call FixMMCFRanges
    Call RebuildARFormatting

    Application.ScreenUpdating = True

    MsgBox "End of meeting sync complete." & Chr(10) & Chr(10) & _
           updated & " existing action(s) overwritten from Minutes." & Chr(10) & _
           "New actions pulled in (see AR).", vbInformation, "End Meeting Done"
End Sub
Sub CleanupAndAlignCheckboxes()
    Dim wsMM As Worksheet
    Dim chk As CheckBox
    Dim cell As Range
    Dim deletedE As Long
    Dim alignedCD As Long
    
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    Application.ScreenUpdating = False
    
    deletedE = 0
    alignedCD = 0
    
    For Each chk In wsMM.CheckBoxes
        Set cell = chk.TopLeftCell
        
        If cell.Column = 5 Then
            chk.Delete
            deletedE = deletedE + 1
        ElseIf cell.Column = 3 Or cell.Column = 4 Then
            chk.Left = cell.Left + 2
            chk.Top = cell.Top + (cell.Height - chk.Height) / 2
            alignedCD = alignedCD + 1
        End If
    Next chk
    
    Application.ScreenUpdating = True
    
    MsgBox deletedE & " col E checkbox(es) deleted." & Chr(10) & _
           alignedCD & " col C/D checkbox(es) aligned.", vbInformation
End Sub

Sub BuildChangelogTab()
    Dim ws As Worksheet
    Dim pw As String
    pw = "1502"
    
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Changelog").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Changelog"
    ws.Tab.Color = RGB(31, 56, 100)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 70
    ws.Columns("F").ColumnWidth = 3
    
    ws.Rows(1).rowHeight = 36
    ws.Range("A1:F1").Merge
    With ws.Range("A1:F1")
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.bold = True
        .Value = "  Engineering Team Minutes  —  Changelog"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A1:F1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 173, 71)
    End With
    
    ws.Rows(2).rowHeight = 8
    ws.Rows(3).rowHeight = 22
    With ws.Range("B3:E3")
        .Merge
        .Interior.Color = RGB(238, 244, 251)
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.bold = True
        .Font.Color = RGB(31, 56, 100)
        .Value = "Creator:  Adina Lieblich  —  Project Manager, Engineering Team"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    
    ws.Rows(4).rowHeight = 20
    With ws.Range("B4:E4")
        .Merge
        .Interior.Color = RGB(238, 244, 251)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
        .Value = "Original author of this meeting minutes and action tracking system. All rights reserved."
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    
    ws.Rows(5).rowHeight = 8
    
    ws.Rows(6).rowHeight = 24
    Dim hdrs As Variant
    hdrs = Array("Version", "Date", "Author", "Changes")
    Dim c As Long
    For c = 0 To 3
        With ws.Cells(6, c + 2)
            .Value = hdrs(c)
            .Interior.Color = RGB(31, 56, 100)
            .Font.Color = RGB(255, 255, 255)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.bold = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
    Next c
    
    ws.Rows(7).rowHeight = 60
    With ws.Cells(7, 2)
        .Value = "v1.0"
        .Font.bold = True
        .Font.Color = RGB(31, 56, 100)
    End With
    With ws.Cells(7, 3)
        .Value = Date
        .NumberFormat = "DD/MM/YYYY"
    End With
    ws.Cells(7, 4).Value = "Adina Lieblich"
    ws.Cells(7, 5).Value = "First official release. Two-tab system: Meeting Minutes + Action Register + Risk Register + Summary + How To Use. Buttons: Start Meeting (Sync All), End Meeting, Hide/Show Done."
    
    With ws.Range("B7:E7")
        .Interior.Color = RGB(250, 250, 250)
        .Font.Name = "Arial"
        .Font.Size = 9
        .VerticalAlignment = xlTop
        .WrapText = True
        .IndentLevel = 1
    End With
    ws.Range("B7:E7").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("B7:E7").Borders(xlEdgeBottom).Color = RGB(220, 220, 220)
    
    ws.Rows(9).rowHeight = 8
    ws.Rows(10).rowHeight = 40
    With ws.Range("B10:E10")
        .Merge
        .Interior.Color = RGB(255, 248, 232)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(90, 61, 0)
        .Value = "To add a new version entry: unprotect the sheet (Review > Unprotect Sheet > enter password 1502), add a row below the last entry, re-protect the sheet. This sheet is protected to preserve version history integrity."
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("B10:E10").Borders
        .LineStyle = xlContinuous
        .Color = RGB(186, 133, 32)
        .Weight = xlThin
    End With
    
    ws.Protect Password:=pw, UserInterfaceOnly:=False, AllowFormattingCells:=False, _
               AllowFormattingColumns:=False, AllowFormattingRows:=False, _
               AllowInsertingRows:=False, AllowDeletingRows:=False
    
    MsgBox "Changelog tab built and protected with password 1502.", vbInformation
End Sub
