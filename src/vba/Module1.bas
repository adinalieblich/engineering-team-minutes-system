Attribute VB_Name = "Module1"
' =========================================================================
' MODULE 1 — Engineering Team Meeting Minutes System
' Live macros for daily use. "Option 3" architecture.
'
' Column layout (Meeting Minutes sheet):
'   A = Type        Dropdown: Action / Risk / blank
'   B = Description Free text
'   C = Owner       Dropdown: TM1..TM6, ALL
'   D = Due         Date DD/MM/YYYY
'   E = Status      Dropdown: Open / In Progress / Done / On Hold / Waiting
'   F = Priority    Dropdown: Critical / High / Medium / Low
'   G = Days OD     Formula (auto)
'   H = Notes       Free text
'   I = Project     Hidden, used for AR filter
'   J = Date Added  Hidden, auto-stamped on row creation, NEVER overwritten
'
' Row layout: rows 1-14 = header block (DO NOT TOUCH), row 15 = column
' headers, row 16+ = data.
'
' See docs/CLAUDE.md and docs/What-AI-Wont-Tell-You.html for the recurring
' VBA gotchas that shaped this file:
'   - Pattern = xlSolid MUST be set before Interior.Color on FormatConditions
'   - xlExpression with $Col absolute refs on contiguous sqrefs (never xlCellValue
'     on growing data)
'   - Plain quotes in Validation Formula1 CSV; triple quotes in CF formulas
'   - Module 1 is live code; old/build macros live in docs/Macro-Archive.html
' =========================================================================

Option Explicit

Private Const SHEET_MM     As String = "Meeting Minutes"
Private Const SHEET_RR     As String = "Risk Register"
Private Const SHEET_SUM    As String = "Summary"
Private Const SHEET_HTU    As String = "How To Use"
Private Const SHEET_LOG    As String = "Changelog"
Private Const DATA_START   As Long = 16
Private Const HEADER_ROW   As Long = 15
Private Const LAST_DATA    As Long = 1000

' Column indices on Meeting Minutes
Private Const COL_TYPE     As Long = 1   ' A
Private Const COL_DESC     As Long = 2   ' B
Private Const COL_OWNER    As Long = 3   ' C
Private Const COL_DUE      As Long = 4   ' D
Private Const COL_STATUS   As Long = 5   ' E
Private Const COL_PRIORITY As Long = 6   ' F
Private Const COL_OD       As Long = 7   ' G
Private Const COL_NOTES    As Long = 8   ' H
Private Const COL_PROJECT  As Long = 9   ' I
Private Const COL_DATE_ADD As Long = 10  ' J


' =========================================================================
' DAILY MACROS
' =========================================================================

' --- AddRow ---------------------------------------------------------------
' Ctrl+Shift+A. Inserts a data row immediately below the active row.
' Inherits Project (col I) from the row above (walks up past banners).
' Stamps Date Added (col J) with TODAY. Sets Days OD formula on col G.
Sub AddRow()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MM)

    Dim activeR As Long
    activeR = ActiveCell.Row
    If activeR < DATA_START Then
        MsgBox "Click a data row (row " & DATA_START & " or below) before running AddRow.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim newR As Long
    newR = activeR + 1
    ws.Rows(newR).Insert Shift:=xlDown
    ws.Rows(newR).ClearFormats

    ' Default row formatting
    With ws.Range("A" & newR & ":J" & newR)
        .Font.Name = "Arial"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .RowHeight = 15
    End With
    ws.Cells(newR, COL_DUE).NumberFormat = "DD/MM/YYYY"
    ws.Cells(newR, COL_DATE_ADD).NumberFormat = "DD/MM/YYYY"

    ' Inherit project from nearest non-banner row above
    Dim projVal As String
    projVal = NearestProjectAbove(ws, activeR)
    ws.Cells(newR, COL_PROJECT).Value = projVal

    ' Stamp Date Added (today's date — never overwritten)
    ws.Cells(newR, COL_DATE_ADD).Value = Date

    ' Days OD formula
    ws.Cells(newR, COL_OD).Formula = OD_Formula(newR)
    ws.Cells(newR, COL_OD).NumberFormat = "0"

    ' Cursor to Description
    ws.Cells(newR, COL_DESC).Activate

    Application.ScreenUpdating = True
End Sub

' --- AddProject -----------------------------------------------------------
' Ctrl+Shift+P. Inserts a project banner row below the active row, then a
' fresh data row beneath it primed for typing.
Sub AddProject()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MM)

    Dim activeR As Long
    activeR = ActiveCell.Row
    If activeR < DATA_START - 1 Then
        MsgBox "Click in the data area before running AddProject.", vbExclamation
        Exit Sub
    End If

    Dim projName As String
    projName = InputBox("Project name (e.g. ""Project A""):", "Add Project")
    If Trim(projName) = "" Then Exit Sub

    Application.ScreenUpdating = False

    Dim bannerR As Long: bannerR = activeR + 1
    Dim dataR As Long:   dataR = activeR + 2

    ws.Rows(bannerR).Insert Shift:=xlDown
    ws.Rows(bannerR).Insert Shift:=xlDown
    ws.Rows(bannerR).ClearFormats
    ws.Rows(dataR).ClearFormats

    ' Project banner — merged A:J, light blue fill, dark blue text
    ws.Range("A" & bannerR & ":J" & bannerR).Merge
    With ws.Range("A" & bannerR & ":J" & bannerR)
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(181, 212, 244)   ' #B5D4F4
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(4, 44, 83)           ' #042C53
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 2
        .Value = projName
        .RowHeight = 18
    End With

    ' Fresh data row beneath
    With ws.Range("A" & dataR & ":J" & dataR)
        .Font.Name = "Arial"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .RowHeight = 15
    End With
    ws.Cells(dataR, COL_DUE).NumberFormat = "DD/MM/YYYY"
    ws.Cells(dataR, COL_DATE_ADD).NumberFormat = "DD/MM/YYYY"
    ws.Cells(dataR, COL_PROJECT).Value = projName
    ws.Cells(dataR, COL_DATE_ADD).Value = Date
    ws.Cells(dataR, COL_OD).Formula = OD_Formula(dataR)
    ws.Cells(dataR, COL_OD).NumberFormat = "0"

    ws.Cells(dataR, COL_DESC).Activate
    Application.ScreenUpdating = True
End Sub

' --- ShowActionRegisterView -----------------------------------------------
' Filter MM to Type = "Action", unhide Project col so the AR view shows
' which project each action belongs to.
Sub ShowActionRegisterView()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MM)

    Application.ScreenUpdating = False

    ' Unhide Project col
    ws.Columns(COL_PROJECT).Hidden = False

    ' Ensure AutoFilter exists on the header row, then apply
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(HEADER_ROW, COL_TYPE), ws.Cells(HEADER_ROW, COL_DATE_ADD)).AutoFilter
    ws.Range(ws.Cells(HEADER_ROW, COL_TYPE), ws.Cells(HEADER_ROW, COL_DATE_ADD)).AutoFilter _
        Field:=COL_TYPE, Criteria1:="Action"

    Application.ScreenUpdating = True
End Sub

' --- RestoreFullView ------------------------------------------------------
' Clear the AR filter and re-hide Project + Date Added.
Sub RestoreFullView()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MM)

    Application.ScreenUpdating = False
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Columns(COL_PROJECT).Hidden = True
    ws.Columns(COL_DATE_ADD).Hidden = True
    Application.ScreenUpdating = True
End Sub

' --- ToggleDoneRows -------------------------------------------------------
' Hide/show rows on Meeting Minutes where Status (col E) = "Done".
Sub ToggleDoneRows()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MM)

    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, COL_DESC).End(xlUp).Row
    If lastR < DATA_START Then Exit Sub

    ' Decide direction: if any Done row is visible, we hide. Otherwise show.
    Dim hideMode As Boolean
    hideMode = False
    Dim r As Long
    For r = DATA_START To lastR
        If ws.Cells(r, COL_STATUS).Value = "Done" And Not ws.Rows(r).Hidden Then
            hideMode = True
            Exit For
        End If
    Next r

    Application.ScreenUpdating = False
    For r = DATA_START To lastR
        If ws.Cells(r, COL_STATUS).Value = "Done" Then
            ws.Rows(r).Hidden = hideMode
        End If
    Next r
    Application.ScreenUpdating = True

    ' Reflect state on any button bound to this macro
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.OnAction = "ToggleDoneRows" Then
            If shp.HasTextFrame Then
                shp.TextFrame.Characters.Text = IIf(hideMode, "Show Done", "Hide Done")
            End If
        End If
    Next shp
End Sub

' --- ToggleClosedRisks ----------------------------------------------------
' Hide/show rows on Risk Register where Status = "Closed".
Sub ToggleClosedRisks()
    Const RR_STATUS_COL As Long = 9   ' I on Risk Register

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_RR)

    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastR < 4 Then Exit Sub

    Dim hideMode As Boolean
    hideMode = False
    Dim r As Long
    For r = 4 To lastR
        If ws.Cells(r, RR_STATUS_COL).Value = "Closed" And Not ws.Rows(r).Hidden Then
            hideMode = True
            Exit For
        End If
    Next r

    Application.ScreenUpdating = False
    For r = 4 To lastR
        If ws.Cells(r, RR_STATUS_COL).Value = "Closed" Then
            ws.Rows(r).Hidden = hideMode
        End If
    Next r
    Application.ScreenUpdating = True

    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.OnAction = "ToggleClosedRisks" Then
            If shp.HasTextFrame Then
                shp.TextFrame.Characters.Text = IIf(hideMode, "Show Closed", "Hide Closed")
            End If
        End If
    Next shp
End Sub

' --- PopulateRiskRegister -------------------------------------------------
' Additive copy of MM rows where Type = "Risk" into the Risk Register sheet.
' Matched by compound key (Description + Owner). Existing RR entries are
' NEVER overwritten — RR owns the row once it's been classified.
Sub PopulateRiskRegister()
    Const RR_PROJECT     As Long = 2
    Const RR_DATE        As Long = 3
    Const RR_DESCRIPTION As Long = 4
    Const RR_OWNER       As Long = 7
    Const RR_NOTES       As Long = 10

    Dim wsMM As Worksheet, wsRR As Worksheet
    Set wsMM = ThisWorkbook.Sheets(SHEET_MM)
    Set wsRR = ThisWorkbook.Sheets(SHEET_RR)

    Dim existing As Object
    Set existing = CreateObject("Scripting.Dictionary")

    ' Build set of existing keys in RR
    Dim rrLast As Long
    rrLast = wsRR.Cells(wsRR.Rows.Count, RR_DESCRIPTION).End(xlUp).Row
    Dim r As Long
    For r = 4 To rrLast
        If Trim(CStr(wsRR.Cells(r, RR_DESCRIPTION).Value)) <> "" Then
            existing(CompoundKey(wsRR.Cells(r, RR_DESCRIPTION).Value, _
                                 wsRR.Cells(r, RR_OWNER).Value)) = True
        End If
    Next r

    ' Find first empty row in RR
    Dim writeR As Long
    writeR = rrLast + 1
    If writeR < 4 Then writeR = 4

    ' Scan MM for Risk rows
    Dim mmLast As Long
    mmLast = wsMM.Cells(wsMM.Rows.Count, COL_DESC).End(xlUp).Row
    Dim added As Long: added = 0

    Application.ScreenUpdating = False
    For r = DATA_START To mmLast
        If wsMM.Cells(r, COL_TYPE).Value = "Risk" Then
            Dim k As String
            k = CompoundKey(wsMM.Cells(r, COL_DESC).Value, wsMM.Cells(r, COL_OWNER).Value)
            If Not existing.exists(k) Then
                wsRR.Cells(writeR, 1).Value = "R-" & Format(writeR - 3, "000")
                wsRR.Cells(writeR, RR_PROJECT).Value = wsMM.Cells(r, COL_PROJECT).Value
                wsRR.Cells(writeR, RR_DATE).Value = Date
                wsRR.Cells(writeR, RR_DATE).NumberFormat = "DD/MM/YYYY"
                wsRR.Cells(writeR, RR_DESCRIPTION).Value = wsMM.Cells(r, COL_DESC).Value
                wsRR.Cells(writeR, RR_OWNER).Value = wsMM.Cells(r, COL_OWNER).Value
                wsRR.Cells(writeR, RR_NOTES).Value = wsMM.Cells(r, COL_NOTES).Value
                existing(k) = True
                writeR = writeR + 1
                added = added + 1
            End If
        End If
    Next r
    Application.ScreenUpdating = True

    MsgBox added & " new risk(s) added to Risk Register. Existing entries untouched.", vbInformation
End Sub


' =========================================================================
' ONE-SHOT MACROS — kept in Module1 for re-runs
' =========================================================================

' --- WriteODDaysFormula ---------------------------------------------------
' Writes the Days OD formula into G16:G<LAST_DATA>.
' Blank if Status is Done/On Hold/Waiting, blank if not an Action, blank if
' no Due date, blank if not yet overdue.
Sub WriteODDaysFormula()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MM)

    Application.ScreenUpdating = False
    Dim r As Long
    For r = DATA_START To LAST_DATA
        If Not ws.Cells(r, COL_TYPE).MergeCells Then
            ws.Cells(r, COL_OD).Formula = OD_Formula(r)
            ws.Cells(r, COL_OD).NumberFormat = "0"
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox "Days OD formula written to G" & DATA_START & ":G" & LAST_DATA & ".", vbInformation
End Sub

' --- FixSummaryOverdueFormula ---------------------------------------------
' Rebuilds the Summary tab overdue counts. Assumes Summary col A = owner
' name, col H = overdue count. Range A4:A<n> holds owner rows.
Sub FixSummaryOverdueFormula()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SUM)

    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim totR As Long: totR = 0
    Dim r As Long
    For r = lastR To 1 Step -1
        If ws.Cells(r, 1).Value = "TOTAL" Then
            totR = r
            Exit For
        End If
    Next r
    If totR = 0 Then totR = lastR + 1

    ' Reference live MM data
    Dim ref As String
    ref = "'" & SHEET_MM & "'!"

    For r = 4 To totR - 1
        Dim ownerRef As String
        ownerRef = "A" & r
        ws.Cells(r, 8).Formula = _
            "=COUNTIFS(" & ref & "$A$" & DATA_START & ":$A$" & LAST_DATA & ",""Action""," & _
            ref & "$C$" & DATA_START & ":$C$" & LAST_DATA & "," & ownerRef & "," & _
            ref & "$D$" & DATA_START & ":$D$" & LAST_DATA & ",""<""&TODAY()," & _
            ref & "$E$" & DATA_START & ":$E$" & LAST_DATA & ",""<>Done""," & _
            ref & "$E$" & DATA_START & ":$E$" & LAST_DATA & ",""<>On Hold""," & _
            ref & "$E$" & DATA_START & ":$E$" & LAST_DATA & ",""<>Waiting"")"
        ws.Cells(r, 8).NumberFormat = "0"
        ws.Cells(r, 8).HorizontalAlignment = xlCenter
    Next r

    ' CF: red when > 0
    ws.Range("H4:H" & totR).FormatConditions.Delete
    With ws.Range("H4:H" & totR).FormatConditions.Add(Type:=xlCellValue, _
            Operator:=xlGreater, Formula1:="0")
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .StopIfTrue = False
    End With

    MsgBox "Summary overdue counts updated.", vbInformation
End Sub

' --- BuildSummaryTab ------------------------------------------------------
' Rebuilds the Summary tab from scratch. Owners hard-coded to TM1..TM6, ALL.
Sub BuildSummaryTab()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(SHEET_SUM).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(SHEET_MM))
    ws.Name = SHEET_SUM

    ' Banner
    ws.Range("A1:H1").Merge
    With ws.Range("A1:H1")
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Name = "Arial"
        .Value = "Engineering Team — Summary"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .RowHeight = 28
    End With

    ' Column headers row 3
    Dim heads As Variant
    heads = Array("Owner", "Action total", "Open", "In Progress", "Done", "On Hold", "Waiting", "Overdue")
    Dim c As Long
    For c = 0 To UBound(heads)
        With ws.Cells(3, c + 1)
            .Value = heads(c)
            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(31, 56, 100)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Name = "Arial"
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
    ws.Rows(3).RowHeight = 22
    ws.Columns("A").ColumnWidth = 14
    ws.Range("B:H").ColumnWidth = 13

    ' Owner rows
    Dim owners As Variant
    owners = Array("TM1", "TM2", "TM3", "TM4", "TM5", "TM6", "ALL")
    Dim ref As String
    ref = "'" & SHEET_MM & "'!"
    Dim r As Long, i As Long
    For i = 0 To UBound(owners)
        r = 4 + i
        ws.Cells(r, 1).Value = owners(i)
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 2).Formula = _
            "=COUNTIFS(" & ref & "$A$" & DATA_START & ":$A$" & LAST_DATA & ",""Action""," & _
            ref & "$C$" & DATA_START & ":$C$" & LAST_DATA & ",""" & owners(i) & """)"
        Dim statuses As Variant
        statuses = Array("Open", "In Progress", "Done", "On Hold", "Waiting")
        Dim sIdx As Long
        For sIdx = 0 To UBound(statuses)
            ws.Cells(r, 3 + sIdx).Formula = _
                "=COUNTIFS(" & ref & "$A$" & DATA_START & ":$A$" & LAST_DATA & ",""Action""," & _
                ref & "$C$" & DATA_START & ":$C$" & LAST_DATA & ",""" & owners(i) & """," & _
                ref & "$E$" & DATA_START & ":$E$" & LAST_DATA & ",""" & statuses(sIdx) & """)"
        Next sIdx
        With ws.Range("A" & r & ":H" & r)
            .Font.Name = "Arial"
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        ws.Cells(r, 1).HorizontalAlignment = xlLeft
    Next i

    ' TOTAL row
    Dim totR As Long: totR = 4 + UBound(owners) + 1
    ws.Cells(totR, 1).Value = "TOTAL"
    ws.Cells(totR, 1).Font.Bold = True
    For c = 2 To 8
        ws.Cells(totR, c).Formula = "=SUM(" & ws.Cells(4, c).Address(False, False) & _
                                   ":" & ws.Cells(totR - 1, c).Address(False, False) & ")"
    Next c
    With ws.Range("A" & totR & ":H" & totR)
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(221, 235, 247)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Cells(totR, 1).HorizontalAlignment = xlLeft

    ' Overdue column (H) — calls dedicated builder so logic stays in one place
    FixSummaryOverdueFormula

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    MsgBox "Summary tab rebuilt.", vbInformation
End Sub

' --- BuildHowToUseTab -----------------------------------------------------
' Builds an embedded cheat sheet. Brief version — the canonical reference
' is the HTML in docs/How-To-Use.html.
Sub BuildHowToUseTab()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(SHEET_HTU).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(SHEET_MM))
    ws.Name = SHEET_HTU

    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 80

    ' Banner
    ws.Range("A1:B1").Merge
    With ws.Range("A1:B1")
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Name = "Arial"
        .Value = "How To Use — Meeting Minutes"
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .RowHeight = 28
        .VerticalAlignment = xlCenter
    End With

    Dim r As Long: r = 3
    WriteHTUSection ws, r, "Keyboard shortcuts"
    WriteHTURow ws, r, "Ctrl + Shift + A", "AddRow — insert a new row below the cursor. Project inherits, Date Added auto-stamps."
    WriteHTURow ws, r, "Ctrl + Shift + P", "AddProject — insert a new project banner below the cursor."

    r = r + 1
    WriteHTUSection ws, r, "Buttons"
    WriteHTURow ws, r, "AR View", "Filter Meeting Minutes to Action rows. Click Restore Full View to undo."
    WriteHTURow ws, r, "Hide / Show Done", "Toggle Done rows on Meeting Minutes."
    WriteHTURow ws, r, "Hide / Show Closed", "Toggle Closed risks on the Risk Register."

    r = r + 1
    WriteHTUSection ws, r, "Columns"
    WriteHTURow ws, r, "A — Type", "Action / Risk / blank. Blank = a general note."
    WriteHTURow ws, r, "B — Description", "What was discussed. One point per row."
    WriteHTURow ws, r, "C — Owner", "TM1..TM6, ALL."
    WriteHTURow ws, r, "D — Due", "DD/MM/YYYY."
    WriteHTURow ws, r, "E — Status", "Open / In Progress / Done / On Hold / Waiting."
    WriteHTURow ws, r, "F — Priority", "Critical / High / Medium / Low."
    WriteHTURow ws, r, "G — Days OD", "Auto-calculated days overdue."
    WriteHTURow ws, r, "H — Notes", "Free text for context or updates."
    WriteHTURow ws, r, "I — Project (hidden)", "Auto-filled by AddRow / AddProject."
    WriteHTURow ws, r, "J — Date Added (hidden)", "Stamped on row creation. Never overwritten."

    r = r + 1
    WriteHTUSection ws, r, "Troubleshooting"
    WriteHTURow ws, r, "Row not turning amber", "Set col A = Action exactly (from the dropdown)."
    WriteHTURow ws, r, "Ctrl+Shift+A not working", "Alt+F8 → AddRow → Options → assign Ctrl+Shift+A."
    WriteHTURow ws, r, "Macro security warning", "Click ""Enable Content"" in the yellow bar."

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    MsgBox "How To Use tab rebuilt.", vbInformation
End Sub

' --- BuildChangelogTab ----------------------------------------------------
' Builds a locked amendment-record sheet. Subsequent edits land in fresh
' rows; the macro never overwrites prior entries on re-run.
Sub BuildChangelogTab()
    Dim ws As Worksheet
    Dim isNew As Boolean: isNew = False

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_LOG
        isNew = True
    End If

    If isNew Then
        ws.Columns("A").ColumnWidth = 14
        ws.Columns("B").ColumnWidth = 10
        ws.Columns("C").ColumnWidth = 16
        ws.Columns("D").ColumnWidth = 16
        ws.Columns("E").ColumnWidth = 60

        ws.Range("A1:E1").Merge
        With ws.Range("A1:E1")
            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(31, 56, 100)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = "Arial"
            .Value = "Changelog — Amendment Record"
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            .RowHeight = 28
            .VerticalAlignment = xlCenter
        End With

        Dim heads As Variant
        heads = Array("Date", "Version", "Author", "Section", "Change")
        Dim c As Long
        For c = 0 To UBound(heads)
            With ws.Cells(3, c + 1)
                .Value = heads(c)
                .Interior.Pattern = xlSolid
                .Interior.Color = RGB(221, 235, 247)
                .Font.Color = RGB(31, 56, 100)
                .Font.Bold = True
                .Font.Name = "Arial"
                .Font.Size = 10
            End With
        Next c
        ws.Rows(3).RowHeight = 22
    End If

    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub


' =========================================================================
' PRIVATE HELPERS
' =========================================================================

' Days OD formula for a given row. Triple-quoted CF-style strings — Excel
' needs the literal double-quote to survive into the formula.
Private Function OD_Formula(ByVal r As Long) As String
    OD_Formula = _
        "=IF(OR($E" & r & "=""Done"",$E" & r & "=""On Hold"",$E" & r & "=""Waiting""),""""," & _
        "IF($A" & r & "<>""Action"",""""," & _
        "IF($D" & r & "="""",""""," & _
        "IF(TODAY()>$D" & r & ",TODAY()-$D" & r & ","""")" & _
        ")))"
End Function

' Walk up from row r looking for the nearest data row's Project value. Skips
' banner rows (merged) and blank rows.
Private Function NearestProjectAbove(ws As Worksheet, ByVal r As Long) As String
    Dim i As Long
    For i = r To DATA_START Step -1
        If Not ws.Cells(i, COL_TYPE).MergeCells Then
            Dim v As String
            v = Trim(CStr(ws.Cells(i, COL_PROJECT).Value))
            If Len(v) > 0 Then
                NearestProjectAbove = v
                Exit Function
            End If
        End If
    Next i
    NearestProjectAbove = ""
End Function

' Compound key for verification / matching. Description alone is not unique —
' duplicate descriptions with different owners must remain distinct.
Private Function CompoundKey(d As Variant, o As Variant) As String
    CompoundKey = Trim(CStr(d)) & "||" & Trim(CStr(o))
End Function

' Build a section banner row on the How To Use sheet.
Private Sub WriteHTUSection(ws As Worksheet, ByRef r As Long, ByVal title As String)
    ws.Range("A" & r & ":B" & r).Merge
    With ws.Range("A" & r & ":B" & r)
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(46, 117, 182)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Arial"
        .Font.Size = 11
        .Value = title
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .RowHeight = 22
        .VerticalAlignment = xlCenter
    End With
    r = r + 1
End Sub

' Build a labeled row on the How To Use sheet.
Private Sub WriteHTURow(ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal text As String)
    With ws.Cells(r, 1)
        .Value = label
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .IndentLevel = 1
    End With
    With ws.Cells(r, 2)
        .Value = text
        .Font.Name = "Arial"
        .Font.Size = 10
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With
    ws.Rows(r).RowHeight = 18
    r = r + 1
End Sub
