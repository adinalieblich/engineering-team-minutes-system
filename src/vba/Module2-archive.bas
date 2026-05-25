Sub AlignCheckboxesLeft()
    Dim ws As Worksheet
    Dim chk As CheckBox
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    
    For Each chk In ws.CheckBoxes
        Set cell = chk.TopLeftCell
        
        ' Only touch checkboxes in col C or col D
        If cell.Column = 3 Or cell.Column = 4 Then
            chk.Left = cell.Left + 2
            chk.Top = cell.Top + (cell.Height - chk.Height) / 2
        End If
    Next chk
    
    MsgBox "Checkboxes aligned left.", vbInformation
End Sub

Sub BuildSummaryTab()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsAR As Worksheet
    Dim col As Integer
    Dim i As Integer
    Dim j As Integer
    Dim r As Long
    Dim rowBG As Long
    Dim totRow As Long
    Dim headers As Variant
    Dim hColors As Variant
    Dim hFonts As Variant
    Dim owners As Variant
    Dim statuses As Variant
    Dim statBGs As Variant
    Dim statFGs As Variant

    Set wb = ThisWorkbook
    Set wsAR = wb.Sheets("Action Register")

    On Error Resume Next
    wb.Sheets("Summary").Delete
    On Error GoTo 0

    Set ws = wb.Sheets.Add(After:=wb.Sheets("Action Register"))
    ws.Name = "Summary"
    ws.Tab.Color = RGB(31, 56, 100)
ws.Activate
ActiveWindow.DisplayGridlines = False
    ws.Columns(1).ColumnWidth = 16
    ws.Columns(2).ColumnWidth = 10
    ws.Columns(3).ColumnWidth = 14
    ws.Columns(4).ColumnWidth = 14
    ws.Columns(5).ColumnWidth = 14
    ws.Columns(6).ColumnWidth = 14
    ws.Columns(7).ColumnWidth = 14

    ws.Rows(1).rowHeight = 28
    ws.Range("A1:G1").Merge
    With ws.Cells(1, 1)
        .Value = "Action Register — Summary by Team Member"
        .Font.Name = "Arial": .Font.Size = 12: .Font.bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With

    ws.Rows(2).rowHeight = 20
    ws.Range("A2:G2").Merge
    With ws.Cells(2, 1)
        .Value = "Live counts from Action Register. Run Refresh Action Register macro first to get latest data."
        .Font.Name = "Arial": .Font.Size = 9: .Font.Italic = True: .Font.Color = RGB(46, 114, 182)
        .Interior.Color = RGB(238, 244, 251)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With

    headers = Array("Owner", "Total", "Open", "In Progress", "Done", "On Hold", "Waiting")
    hColors = Array(RGB(31, 56, 100), RGB(31, 56, 100), RGB(242, 242, 242), RGB(255, 243, 205), RGB(226, 239, 218), RGB(214, 228, 247), RGB(252, 228, 214))
    hFonts = Array(RGB(255, 255, 255), RGB(255, 255, 255), RGB(89, 89, 89), RGB(125, 90, 0), RGB(55, 86, 35), RGB(31, 56, 100), RGB(131, 60, 0))

    ws.Rows(3).rowHeight = 26
    For col = 1 To 7
        With ws.Cells(3, col)
            .Value = headers(col - 1)
            .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True
            .Font.Color = hFonts(col - 1)
            .Interior.Color = hColors(col - 1)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin: .Borders(xlEdgeLeft).Color = RGB(68, 114, 196)
            .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin: .Borders(xlEdgeRight).Color = RGB(68, 114, 196)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlMedium: .Borders(xlEdgeBottom).Color = RGB(46, 114, 182)
        End With
    Next col

    owners = Array("TM1", "TM2", "TM3", "TM4", "TM5", "TM6", "TM7", "TM8", "ALL")
    statuses = Array("Open", "In Progress", "Done", "On Hold", "Waiting")
    statBGs = Array(RGB(242, 242, 242), RGB(255, 243, 205), RGB(226, 239, 218), RGB(214, 228, 247), RGB(252, 228, 214))
    statFGs = Array(RGB(89, 89, 89), RGB(125, 90, 0), RGB(55, 86, 35), RGB(31, 56, 100), RGB(131, 60, 0))

    For i = 0 To UBound(owners)
        r = 4 + i
        ws.Rows(r).rowHeight = 22
        If i Mod 2 = 0 Then rowBG = RGB(242, 242, 242) Else rowBG = RGB(250, 250, 250)

        With ws.Cells(r, 1)
            .Value = owners(i)
            .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True: .Font.Color = RGB(31, 56, 100)
            .Interior.Color = rowBG
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlMedium: .Borders(xlEdgeLeft).Color = RGB(214, 214, 214)
            .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlHairline: .Borders(xlEdgeRight).Color = RGB(208, 208, 208)
            .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlHairline: .Borders(xlEdgeTop).Color = RGB(208, 208, 208)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlHairline: .Borders(xlEdgeBottom).Color = RGB(208, 208, 208)
        End With

        With ws.Cells(r, 2)
            .Formula = "=COUNTIF('Action Register'!C4:C103,A" & r & ")"
            .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True: .Font.Color = RGB(31, 56, 100)
            .Interior.Color = rowBG
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlHairline: .Borders(xlEdgeLeft).Color = RGB(208, 208, 208)
            .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlHairline: .Borders(xlEdgeRight).Color = RGB(208, 208, 208)
            .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlHairline: .Borders(xlEdgeTop).Color = RGB(208, 208, 208)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlHairline: .Borders(xlEdgeBottom).Color = RGB(208, 208, 208)
        End With

        For j = 0 To 4
            With ws.Cells(r, 3 + j)
                .Formula = "=COUNTIFS('Action Register'!C4:C103,A" & r & ",'Action Register'!E4:E103,""" & statuses(j) & """)"
                .Font.Name = "Arial": .Font.Size = 9: .Font.Color = statFGs(j)
                .Interior.Color = statBGs(j)
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlHairline: .Borders(xlEdgeLeft).Color = RGB(208, 208, 208)
                .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlHairline: .Borders(xlEdgeRight).Color = RGB(208, 208, 208)
                .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlHairline: .Borders(xlEdgeTop).Color = RGB(208, 208, 208)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlHairline: .Borders(xlEdgeBottom).Color = RGB(208, 208, 208)
            End With
        Next j
    Next i

    totRow = 4 + UBound(owners) + 1
    ws.Rows(totRow).rowHeight = 22
    ws.Range("A" & totRow & ":G" & totRow).Interior.Color = RGB(31, 56, 100)

    With ws.Cells(totRow, 1)
        .Value = "TOTAL"
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True: .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlMedium: .Borders(xlEdgeLeft).Color = RGB(68, 114, 196)
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlHairline: .Borders(xlEdgeRight).Color = RGB(68, 114, 196)
    End With

    For col = 2 To 7
        With ws.Cells(totRow, col)
            .Formula = "=SUM(" & Chr(64 + col) & "4:" & Chr(64 + col) & (totRow - 1) & ")"
            .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True: .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlHairline: .Borders(xlEdgeLeft).Color = RGB(68, 114, 196)
            .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlHairline: .Borders(xlEdgeRight).Color = RGB(68, 114, 196)
        End With
    Next col

    MsgBox "Summary tab built successfully!", vbInformation
End Sub


Sub AddRefreshButton()
    Dim ws As Worksheet
    Dim btn As Shape
    Set ws = Sheets("Action Register")

    Dim existingShape As Shape
    For Each existingShape In ws.Shapes
        If existingShape.Name = "btnRefreshAR" Then
            existingShape.Delete
        End If
    Next existingShape

    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double

    btnWidth = 160
    btnHeight = 24
    btnLeft = ws.Cells(1, 7).Left + ws.Cells(1, 7).Width - btnWidth - 4
    btnTop = ws.Rows(1).Top + 4

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)

    With btn
        .Name = "btnRefreshAR"
        .Fill.ForeColor.RGB = RGB(46, 114, 182)
        .Fill.Solid
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "Refresh from Minutes"
        .TextFrame.Characters.Font.Name = "Arial"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = "PopulateActionRegister"
    End With

    MsgBox "Refresh button added to Action Register tab!", vbInformation
End Sub
Sub AddAllButtons()
    Dim wsMin As Worksheet
    Dim wsAR As Worksheet
    Dim btn As Shape

    Set wsMin = Sheets("Meeting Minutes")
    Set wsAR = Sheets("Action Register")

    ' Clear existing buttons on both sheets
    Dim s As Shape
    For Each s In wsMin.Shapes
        If s.Name = "btnToggleDoneMin" Then s.Delete
    Next s
    For Each s In wsAR.Shapes
        If s.Name = "btnRefreshAR" Or s.Name = "btnToggleDoneAR" Or s.Name = "btnSyncBack" Then s.Delete
    Next s

    ' -- MINUTES TAB — Toggle Done button ------------------------------
    Dim minBtnLeft As Double
    Dim minBtnTop As Double
    minBtnLeft = wsMin.Cells(3, 8).Left + 4
    minBtnTop = wsMin.Rows(3).Top + 4

    Set btn = wsMin.Shapes.AddShape(msoShapeRoundedRectangle, minBtnLeft, minBtnTop, 150, 22)
    With btn
        .Name = "btnToggleDoneMin"
        .Fill.ForeColor.RGB = RGB(46, 114, 182)
        .Fill.Solid
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "Hide / Show Done"
        .TextFrame.Characters.Font.Name = "Arial"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = "ToggleDoneMinutes"
    End With

    ' -- ACTION REGISTER TAB — three buttons ---------------------------
    Dim arBtnLeft As Double
    Dim arBtnTop As Double
    arBtnTop = wsAR.Rows(1).Top + 4

    ' Button 1 — Refresh from Minutes
    arBtnLeft = wsAR.Cells(1, 7).Left + wsAR.Cells(1, 7).Width - 484
    Set btn = wsAR.Shapes.AddShape(msoShapeRoundedRectangle, arBtnLeft, arBtnTop, 155, 22)
    With btn
        .Name = "btnRefreshAR"
        .Fill.ForeColor.RGB = RGB(31, 56, 100)
        .Fill.Solid
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "Refresh from Minutes"
        .TextFrame.Characters.Font.Name = "Arial"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = "PopulateActionRegister"
    End With

    ' Button 2 — Sync Status Back to Minutes
    arBtnLeft = wsAR.Cells(1, 7).Left + wsAR.Cells(1, 7).Width - 322
    Set btn = wsAR.Shapes.AddShape(msoShapeRoundedRectangle, arBtnLeft, arBtnTop, 155, 22)
    With btn
        .Name = "btnSyncBack"
        .Fill.ForeColor.RGB = RGB(31, 56, 100)
        .Fill.Solid
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "Sync Status to Minutes"
        .TextFrame.Characters.Font.Name = "Arial"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = "SyncStatusBack"
    End With

    ' Button 3 — Hide / Show Done
    arBtnLeft = wsAR.Cells(1, 7).Left + wsAR.Cells(1, 7).Width - 160
    Set btn = wsAR.Shapes.AddShape(msoShapeRoundedRectangle, arBtnLeft, arBtnTop, 155, 22)
    With btn
        .Name = "btnToggleDoneAR"
        .Fill.ForeColor.RGB = RGB(46, 114, 182)
        .Fill.Solid
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "Hide / Show Done"
        .TextFrame.Characters.Font.Name = "Arial"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = "ToggleDoneRegister"
    End With

    MsgBox "All buttons added!", vbInformation
End Sub

Sub BuildHowToUseTab()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook

    On Error Resume Next
    wb.Sheets("How To Use").Delete
    On Error GoTo 0

    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "How To Use"
    ws.Tab.Color = RGB(68, 114, 196)
    ws.Activate
ActiveWindow.DisplayGridlines = False

    Dim i As Integer
    Dim r As Long

    ws.Columns(1).ColumnWidth = 3
    ws.Columns(2).ColumnWidth = 28
    ws.Columns(3).ColumnWidth = 22
    ws.Columns(4).ColumnWidth = 48
    ws.Columns(5).ColumnWidth = 3

    Dim NAVY As Long, BLUE As Long, MIDBLUE As Long, LBLUE As Long, SLBLUE As Long
    Dim WHITE As Long, OFFWHITE As Long, GREY As Long, DGREY As Long, INK As Long
    Dim AMBER As Long, AMBERFONT As Long
    Dim DONEBG As Long, DONEFG As Long, INPROGBG As Long, INPROGFG As Long
    Dim ONHOLDBG As Long, ONHOLDFG As Long, WAITBG As Long, WAITFG As Long
    Dim OPENBG As Long, OPENFG As Long

    NAVY = RGB(31, 56, 100): BLUE = RGB(46, 114, 182): MIDBLUE = RGB(68, 114, 196)
    LBLUE = RGB(214, 228, 247): SLBLUE = RGB(238, 244, 251)
    WHITE = RGB(255, 255, 255): OFFWHITE = RGB(250, 250, 250): GREY = RGB(242, 242, 242)
    DGREY = RGB(89, 89, 89): INK = RGB(26, 26, 46)
    AMBER = RGB(255, 243, 205): AMBERFONT = RGB(193, 123, 0)
    DONEBG = RGB(226, 239, 218): DONEFG = RGB(55, 86, 35)
    INPROGBG = RGB(255, 243, 205): INPROGFG = RGB(125, 90, 0)
    ONHOLDBG = RGB(214, 228, 247): ONHOLDFG = RGB(31, 56, 100)
    WAITBG = RGB(252, 228, 214): WAITFG = RGB(131, 60, 0)
    OPENBG = RGB(242, 242, 242): OPENFG = RGB(89, 89, 89)

    r = 1
    ws.Rows(r).rowHeight = 54
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Merge
    With ws.Cells(r, 1)
        .Value = "ENGINEERING TEAM MEETING MINUTES SYSTEM"
        .Font.Name = "Arial": .Font.Size = 18: .Font.bold = True: .Font.Color = WHITE
        .Interior.Color = NAVY
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    r = r + 1

    ws.Rows(r).rowHeight = 4
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Merge
    ws.Cells(r, 1).Interior.Color = MIDBLUE
    r = r + 1

    ws.Rows(r).rowHeight = 24
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Merge
    With ws.Cells(r, 1)
        .Value = "Engineering Team — Engineering Team   |   How To Use This Document"
        .Font.Name = "Arial": .Font.Size = 11: .Font.Color = NAVY: .Font.Italic = True
        .Interior.Color = SLBLUE
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    r = r + 1

    ws.Rows(r).rowHeight = 8
    r = r + 1

    Call WriteSection(ws, r, "01  OVERVIEW", NAVY, MIDBLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1
    Call WriteOverview(ws, r, NAVY, BLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY, AMBER, AMBERFONT)
    r = ws.UsedRange.Rows.Count + 1

    Call WriteSection(ws, r, "02  THE THREE TABS", NAVY, MIDBLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1
    Call WriteThreeTabs(ws, r, NAVY, BLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1

    Call WriteSection(ws, r, "03  MEETING MINUTES — COLUMN GUIDE", NAVY, MIDBLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1
    Call WriteColumnGuide(ws, r, NAVY, BLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY, AMBER, AMBERFONT)
    r = ws.UsedRange.Rows.Count + 1

    Call WriteSection(ws, r, "04  COLOUR CODE GUIDE", NAVY, MIDBLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1
    Call WriteColourGuide(ws, r, NAVY, BLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY, AMBER, AMBERFONT, DONEBG, DONEFG, INPROGBG, INPROGFG, ONHOLDBG, ONHOLDFG, WAITBG, WAITFG, OPENBG, OPENFG)
    r = ws.UsedRange.Rows.Count + 1

    Call WriteSection(ws, r, "05  WEEKLY WORKFLOW", NAVY, MIDBLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1
    Call WriteWorkflow(ws, r, NAVY, BLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY, AMBER, AMBERFONT)
    r = ws.UsedRange.Rows.Count + 1

    Call WriteSection(ws, r, "06  KEYBOARD SHORTCUTS & BUTTONS", NAVY, MIDBLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1
    Call WriteShortcuts(ws, r, NAVY, BLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1

    Call WriteSection(ws, r, "07  ACTION REGISTER", NAVY, MIDBLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1
    Call WriteActionRegisterGuide(ws, r, NAVY, BLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1

    Call WriteSection(ws, r, "08  SUMMARY TAB", NAVY, MIDBLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1
    Call WriteSummaryGuide(ws, r, NAVY, BLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1

    Call WriteSection(ws, r, "09  TROUBLESHOOTING", NAVY, MIDBLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)
    r = ws.UsedRange.Rows.Count + 1
    Call WriteTroubleshooting(ws, r, NAVY, BLUE, SLBLUE, GREY, OFFWHITE, WHITE, INK, DGREY)

    MsgBox "How To Use tab built successfully!", vbInformation, "Done"
End Sub

Sub WriteSection(ws As Worksheet, r As Long, title As String, _
    NAVY As Long, MIDBLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long)
    ws.Rows(r).rowHeight = 6
    r = r + 1
    ws.Rows(r).rowHeight = 26
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Merge
    With ws.Cells(r, 1)
        .Value = title
        .Font.Name = "Arial": .Font.Size = 11: .Font.bold = True: .Font.Color = WHITE
        .Interior.Color = NAVY
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = MIDBLUE
    End With
End Sub

Sub writeRow(ws As Worksheet, r As Long, label As String, content As String, _
    labelBG As Long, contentBG As Long, labelFG As Long, contentFG As Long, _
    Optional bold As Boolean = False, Optional rowHeight As Double = 28)
    ws.Rows(r).rowHeight = rowHeight
    ws.Cells(r, 1).Interior.Color = labelBG
    ws.Cells(r, 2).Value = label
    With ws.Cells(r, 2)
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True: .Font.Color = labelFG
        .Interior.Color = labelBG
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlTop: .WrapText = True: .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlHairline: .Borders(xlEdgeBottom).Color = RGB(220, 220, 220)
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlHairline: .Borders(xlEdgeRight).Color = RGB(220, 220, 220)
    End With
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Merge
    ws.Cells(r, 3).Value = content
    With ws.Cells(r, 3)
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = bold: .Font.Color = contentFG
        .Interior.Color = contentBG
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlTop: .WrapText = True: .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlHairline: .Borders(xlEdgeBottom).Color = RGB(220, 220, 220)
    End With
    ws.Cells(r, 5).Interior.Color = contentBG
    ws.Cells(r, 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(r, 5).Borders(xlEdgeBottom).Weight = xlHairline
    ws.Cells(r, 5).Borders(xlEdgeBottom).Color = RGB(220, 220, 220)
End Sub

Sub WriteSubHeader(ws As Worksheet, r As Long, title As String, BLUE As Long, WHITE As Long)
    ws.Rows(r).rowHeight = 20
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Merge
    With ws.Cells(r, 1)
        .Value = "  " & title
        .Font.Name = "Arial": .Font.Size = 9: .Font.bold = True: .Font.Color = WHITE
        .Interior.Color = BLUE
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
End Sub

Sub WriteOverview(ws As Worksheet, r As Long, _
    NAVY As Long, BLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long, AMBER As Long, AMBERFONT As Long)
    Dim alt As Boolean: alt = False
    Call writeRow(ws, r, "What is this?", "A fully automated meeting minutes and action tracking system built in Microsoft Excel. Designed for project-based teams, it captures meeting notes, flags action items, and automatically populates an Action Register — all without leaving Excel.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Who is it for?", "Project managers, engineers, and team leads who run regular status meetings and need a simple system to track what was discussed, what needs to be done, and who is doing it.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "What does it need?", "Microsoft Excel only. No plugins, no licences, no admin rights required. Works within SharePoint and standard corporate IT environments.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "How does it work?", "You take minutes during the meeting. Anything requiring follow-up gets ticked as an ACTION. At the end of the meeting, one click populates the Action Register. The team updates statuses during the week. Before the next meeting, one click syncs everything back. The Summary tab always shows who has what outstanding.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=44): r = r + 1
End Sub

Sub WriteThreeTabs(ws As Worksheet, r As Long, _
    NAVY As Long, BLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long)
    Dim alt As Boolean: alt = False
    Call writeRow(ws, r, "Meeting Minutes", "Where you take notes during the meeting. One row per agenda item. Use section headers (dark navy) and project sub-headers (mid blue) to structure the meeting. Tick the ACTION? checkbox for anything requiring follow-up.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Action Register", "Automatically populated from Meeting Minutes. Shows all ticked actions with owner, due date and status. The team updates statuses here throughout the week. Use the filter arrows to view by owner, project or status.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Summary", "A live count of actions by team member and status — Total, Open, In Progress, Done, On Hold and Waiting. Updates automatically when the Action Register is refreshed. Review at the start of each meeting.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "How To Use", "This tab. Read it once, refer back as needed.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK): r = r + 1
End Sub

Sub WriteColumnGuide(ws As Worksheet, r As Long, _
    NAVY As Long, BLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long, AMBER As Long, AMBERFONT As Long)
    Dim alt As Boolean: alt = False
    Call writeRow(ws, r, "A  —  Section / Project", "The project or agenda section this row relates to. Only fill on the FIRST row of each project. Leave blank on continuation rows — this keeps the minutes clean and readable.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "B  —  Description / Notes", "What was discussed, decided, or what needs to be done. Keep to one clear point per row. For action rows write it as an instruction: e.g. 'Email Western Power re works date'.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "C  —  ACTION?", "Tick this checkbox if the row requires someone to do something. The row immediately turns amber to stand out from general notes. Unticked rows are general notes — no action required.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "D  —  Owner", "Who is responsible for completing this action. Select from the dropdown. Required for all action rows.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=28): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "E  —  Due Date", "When the action must be completed. Always use DD/MM/YYYY format — e.g. 07/04/2026. Never use short form (7/4/26) or long form (7 April 2026). Required for all action rows.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "F  —  Status", "Only fill when ACTION? is ticked. Options: Open, In Progress, Done, On Hold, Waiting. Grey and inactive on non-action rows. Cell colour updates automatically to match the status.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "G  —  Notes / Update", "Additional context, weekly updates, or relevant information. Use to capture progress notes without changing the status.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=28): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "H  —  Project (hidden)", "A hidden column storing the project name for every row. Used by the Action Register macro to group actions by project. Do not edit or unhide — maintained automatically by macros.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), DGREY, rowHeight:=36): r = r + 1
End Sub

Sub WriteColourGuide(ws As Worksheet, r As Long, _
    NAVY As Long, BLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long, _
    AMBER As Long, AMBERFONT As Long, DONEBG As Long, DONEFG As Long, _
    INPROGBG As Long, INPROGFG As Long, ONHOLDBG As Long, ONHOLDFG As Long, _
    WAITBG As Long, WAITFG As Long, OPENBG As Long, OPENFG As Long)
    Call WriteSubHeader(ws, r, "Row colours — Meeting Minutes", BLUE, WHITE): r = r + 1
    Call writeRow(ws, r, "Dark navy header", "Main section header (e.g. 3.0 ROADS / CYCLEWAYS). Do not edit these rows.", NAVY, NAVY, WHITE, WHITE): r = r + 1
    Call writeRow(ws, r, "Mid blue header", "Project sub-header (e.g. Project A | Budget: $X.XXM | PM: TM1/TM4). Do not edit these rows.", BLUE, BLUE, WHITE, WHITE): r = r + 1
    Call writeRow(ws, r, "Amber row", "ACTION? is ticked — action required. Fill in Owner and Due Date.", AMBER, AMBER, AMBERFONT, AMBERFONT, True): r = r + 1
    Call writeRow(ws, r, "Green row", "ACTION ticked + Status = Done. Task is complete.", DONEBG, DONEBG, DONEFG, DONEFG, True): r = r + 1
    Call writeRow(ws, r, "Amber/gold row", "ACTION ticked + Status = In Progress. Someone is actively working on this.", INPROGBG, INPROGBG, INPROGFG, INPROGFG, True): r = r + 1
    Call writeRow(ws, r, "Blue row", "ACTION ticked + Status = On Hold. Paused — waiting on a decision or external factor.", ONHOLDBG, ONHOLDBG, ONHOLDFG, ONHOLDFG, True): r = r + 1
    Call writeRow(ws, r, "Pink row", "ACTION ticked + Status = Waiting. Waiting on someone else to respond or act.", WAITBG, WAITBG, WAITFG, WAITFG, True): r = r + 1
    Call writeRow(ws, r, "White / grey row", "General note or update — ACTION? not ticked. No follow-up required.", OFFWHITE, OFFWHITE, INK, INK): r = r + 1
    Call WriteSubHeader(ws, r, "Status cell colours — col F", BLUE, WHITE): r = r + 1
    Call writeRow(ws, r, "Open", "Task acknowledged — not yet started.", OPENBG, OPENBG, OPENFG, OPENFG): r = r + 1
    Call writeRow(ws, r, "In Progress", "Someone is actively working on it.", INPROGBG, INPROGBG, INPROGFG, INPROGFG): r = r + 1
    Call writeRow(ws, r, "Done", "Completed. Will be hidden when Hide Done is toggled.", DONEBG, DONEBG, DONEFG, DONEFG): r = r + 1
    Call writeRow(ws, r, "On Hold", "Paused — blocked or deferred.", ONHOLDBG, ONHOLDBG, ONHOLDFG, ONHOLDFG): r = r + 1
    Call writeRow(ws, r, "Waiting", "Waiting on another person or team to act.", WAITBG, WAITBG, WAITFG, WAITFG): r = r + 1
    Call writeRow(ws, r, "Grey (no status)", "Not an action row — status is inactive.", RGB(235, 235, 235), RGB(235, 235, 235), DGREY, DGREY): r = r + 1
End Sub
Sub WriteWorkflow(ws As Worksheet, r As Long, _
    NAVY As Long, BLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long, AMBER As Long, AMBERFONT As Long)

    Call WriteSubHeader(ws, r, "Tuesday — During the meeting", BLUE, WHITE): r = r + 1
    Call writeRow(ws, r, "Step 1", "Open this week's file from SharePoint.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 2", "Update the Date cell in row 4 to today's date.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 3", "Update the attendees table — tick Attended? for each person present.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 4", "Take minutes row by row. Use Ctrl+Shift+A to add rows, Ctrl+Shift+S for new projects.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 5", "For any item requiring action — tick the ACTION? checkbox. Fill in Owner and Due Date.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 6", "At the end of the meeting — click Refresh from Minutes on the Action Register tab.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 7", "Save the file. SharePoint will sync automatically.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1

    ws.Rows(r).rowHeight = 6: r = r + 1

    Call WriteSubHeader(ws, r, "Tuesday to Monday — During the week", BLUE, WHITE): r = r + 1
    Call writeRow(ws, r, "Step 1", "Team members open the file from SharePoint.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 2", "Go to the Action Register tab.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 3", "Filter by your name using the Owner dropdown filter.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 4", "Update the Status of your actions as you complete them.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 5", "Add notes in the Notes column if needed.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 6", "Save the file. Changes sync to SharePoint automatically.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK): r = r + 1

    ws.Rows(r).rowHeight = 6: r = r + 1

    Call WriteSubHeader(ws, r, "Next Tuesday — Before the meeting", BLUE, WHITE): r = r + 1
    Call writeRow(ws, r, "Step 1", "Open the file from SharePoint.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 2", "On the Action Register tab — click Sync Status to Minutes to push updated statuses back into the Meeting Minutes tab.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1
    Call writeRow(ws, r, "Step 3", "Use Hide / Show Done to declutter the view — hides completed actions.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 4", "Review the Summary tab — check who has outstanding actions before the meeting starts.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 5", "Take this week's minutes on top of the existing document.", GREY, GREY, RGB(31, 56, 100), INK): r = r + 1
    Call writeRow(ws, r, "Step 6", "At the end of the meeting — click Refresh from Minutes again to update the register with new actions.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1
End Sub

Sub WriteShortcuts(ws As Worksheet, r As Long, _
    NAVY As Long, BLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long)

    Call WriteSubHeader(ws, r, "Keyboard shortcuts — Meeting Minutes tab", BLUE, WHITE): r = r + 1
    Call writeRow(ws, r, "Ctrl + Shift + A", "Add Row — inserts a blank formatted row at the end of the current project subsection. Click any row in the project first.", GREY, GREY, RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1
    Call writeRow(ws, r, "Ctrl + Shift + S", "Add Subsection — adds a new project sub-header and blank row within the current section. Prompts for project name.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1
    Call writeRow(ws, r, "Ctrl + Shift + R", "Add Section — adds a full new section header, sub-header and blank row. Prompts for section name, sub-heading and project name.", GREY, GREY, RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1

    ws.Rows(r).rowHeight = 6: r = r + 1

    Call WriteSubHeader(ws, r, "Buttons — Action Register tab", BLUE, WHITE): r = r + 1
    Call writeRow(ws, r, "Refresh from Minutes", "Reads all ticked ACTION? rows from Meeting Minutes and writes them to the Action Register. Run this at the end of every meeting.", GREY, GREY, RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1
    Call writeRow(ws, r, "Sync Status to Minutes", "Reads statuses from the Action Register and pushes them back into the Meeting Minutes tab. Run this before each meeting.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1
    Call writeRow(ws, r, "Hide / Show Done", "Toggles Done rows on and off. Click once to hide completed actions, click again to show them.", GREY, GREY, RGB(31, 56, 100), INK, rowHeight:=28): r = r + 1

    ws.Rows(r).rowHeight = 6: r = r + 1

    Call WriteSubHeader(ws, r, "Button — Meeting Minutes tab", BLUE, WHITE): r = r + 1
    Call writeRow(ws, r, "Hide / Show Done", "Hides all rows where Status = Done in the Meeting Minutes tab. Click again to show them.", OFFWHITE, OFFWHITE, RGB(31, 56, 100), INK, rowHeight:=28): r = r + 1
End Sub

Sub WriteActionRegisterGuide(ws As Worksheet, r As Long, _
    NAVY As Long, BLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long)
    Dim alt As Boolean: alt = False
    Call writeRow(ws, r, "What it contains", "Every row from Meeting Minutes where ACTION? is ticked. Automatically populated by the Refresh macro.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Columns", "Project  |  Description  |  Owner  |  Due Date  |  Status  |  Date Added  |  Notes", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Filtering", "Click the filter arrow on any column header to filter by Owner, Status or Project. Use this during the meeting to quickly review one person's actions.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Updating status", "Team members update the Status column directly in the Action Register during the week. Before the next meeting, run Sync Status to Minutes to push changes back.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Source of truth", "The Action Register is the source of truth for STATUS during the week. The Meeting Minutes is the source of truth for everything else — description, owner, due date.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Capacity", "Holds up to 100 action rows. If you exceed this — mark completed ones as Done and use Hide Done to declutter.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=28): r = r + 1
End Sub

Sub WriteSummaryGuide(ws As Worksheet, r As Long, _
    NAVY As Long, BLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long)
    Dim alt As Boolean: alt = False
    Call writeRow(ws, r, "What it shows", "A count of actions by team member broken down by status: Total, Open, In Progress, Done, On Hold, Waiting.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "How to refresh", "Uses live COUNTIFS formulas reading from the Action Register. Updates automatically when the Action Register is refreshed. No extra step required.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Adding team members", "Insert a row in the Summary tab and type their initials in col A. Copy the formula from an adjacent cell in cols B-G.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Removing team members", "Delete their row in the Summary tab. Also remove them from the Owner dropdown via Data tab ? Data Validation on any Owner cell.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=32): r = r + 1
End Sub

Sub WriteTroubleshooting(ws As Worksheet, r As Long, _
    NAVY As Long, BLUE As Long, SLBLUE As Long, GREY As Long, OFFWHITE As Long, _
    WHITE As Long, INK As Long, DGREY As Long)
    Dim alt As Boolean: alt = False
    Call writeRow(ws, r, "Row not turning amber", "Check the ACTION? cell — it must contain TRUE (boolean), not the text 'TRUE'. Make sure the checkbox is linked to the correct cell via Format Control ? Cell Link.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Status colour not showing", "Check Conditional Formatting rules (Home ? Conditional Formatting ? Manage Rules). Status rules on F22:F180 must be separate from row rules on A22:G180. Check priority order.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Add Row inserts in wrong place", "Click a DATA row (not a header) before pressing Ctrl+Shift+A. The macro scans down from your cursor to find the next merged header row and inserts just before it.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Action Register not updating", "Run the Refresh from Minutes macro — the register does not update automatically. Check that ACTION? cells contain TRUE (boolean) not text.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Sync Status not matching", "The sync matches by Description AND Owner. If either has been edited since the last refresh the row will not match. Check the 'not matched' count in the message box and update manually.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "File saving to wrong location", "Always open the file from the synced SharePoint folder in File Explorer — not from the browser URL. Browser opens save to personal OneDrive by default.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1: alt = Not alt
    Call writeRow(ws, r, "Macro security warning", "If you see a yellow bar saying macros are disabled — click Enable Content. You may need to do this each time the file is opened from SharePoint.", IIf(alt, GREY, OFFWHITE), IIf(alt, GREY, OFFWHITE), RGB(31, 56, 100), INK, rowHeight:=36): r = r + 1
End Sub


Sub FixARDateFormat()
    Dim wsAR As Worksheet
    Set wsAR = Sheets("Action Register")
    
    wsAR.Range("D4:D103").NumberFormat = "DD/MM/YYYY"
    wsAR.Range("F4:F103").NumberFormat = "DD/MM/YYYY"
    
    MsgBox "Done — dates now showing DD/MM/YYYY", vbInformation
End Sub

Sub RestoreAttendeeCheckboxes()
    Dim ws As Worksheet
    Dim chk As CheckBox
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Meeting Minutes")
    
    For Each chk In ws.CheckBoxes
        Set cell = chk.TopLeftCell
        
        ' Only touch checkboxes in col C rows 8-14
        If cell.Column = 3 And cell.Row >= 8 And cell.Row <= 14 Then
            chk.Left = cell.Left + 2
            chk.Top = cell.Top + (cell.Height - chk.Height) / 2
            chk.Width = 15
            chk.Height = 15
            chk.Visible = True
        End If
    Next chk
    
    MsgBox "Attendee checkboxes restored.", vbInformation
End Sub

Sub FixARDropdowns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Action Register")
    
    ' Fix Status dropdown col F rows 4-103
    With ws.Range("F4:F103").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Open,In Progress,Done,On Hold,Waiting"
        .ShowError = False
    End With
    
    ' Fix Priority dropdown col E rows 4-103
    With ws.Range("E4:E103").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Critical,High,Medium,Low"
        .ShowError = False
    End With
    
    ' Fix Owner dropdown col C rows 4-103
    With ws.Range("C4:C103").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="TM1,TM2,TM3,TM4,TM5,TM6,ALL"
        .ShowError = False
    End With
    
    MsgBox "Action Register dropdowns fixed.", vbInformation
End Sub
Sub BuildRiskRegisterCF()
    Dim ws As Worksheet
    Dim rng As Range
    Dim fc As FormatCondition
    Set ws = ThisWorkbook.Sheets("Risk Register")
    
    ' -- Clear all existing CF on data rows --
    ws.Range("A4:J103").FormatConditions.Delete
    
    ' --------------------------------------
    ' LIKELIHOOD col E — font colour only
    ' --------------------------------------
    Set rng = ws.Range("E4:E103")
    
    ' High — white bold
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="High")
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.bold = True
    
    ' Medium — dark amber bold
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Medium")
    fc.Font.Color = RGB(125, 82, 0)
    fc.Font.bold = True
    
    ' Low — dark green bold
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Low")
    fc.Font.Color = RGB(55, 86, 35)
    fc.Font.bold = True
    
    ' --------------------------------------
    ' RISK RATING col F — font colour only
    ' --------------------------------------
    Set rng = ws.Range("F4:F103")
    
    ' High — white bold
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="High")
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.bold = True
    
    ' Medium — dark amber bold
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Medium")
    fc.Font.Color = RGB(125, 82, 0)
    fc.Font.bold = True
    
    ' Low — dark green bold
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Low")
    fc.Font.Color = RGB(55, 86, 35)
    fc.Font.bold = True
    
    ' --------------------------------------
    ' STATUS col I — font colour only
    ' --------------------------------------
    Set rng = ws.Range("I4:I103")
    
    ' Open — grey
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Open")
    fc.Font.Color = RGB(89, 89, 89)
    fc.Font.bold = False
    
    ' Monitored — amber
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Monitored")
    fc.Font.Color = RGB(125, 82, 0)
    fc.Font.bold = True
    
    ' Closed — green
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Closed")
    fc.Font.Color = RGB(55, 86, 35)
    fc.Font.bold = True
    
    ' --------------------------------------
    ' CLOSED ROWS — grey out entire row
    ' Formula: if col I = Closed, grey all text A:J
    ' --------------------------------------
    Set rng = ws.Range("A4:J103")
    
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$I4=""Closed""")
    fc.Font.Color = RGB(166, 166, 166)
    fc.Font.bold = False
    fc.Font.Italic = True
    
    MsgBox "Risk Register CF rules applied. Now set fills manually in Manage Rules.", vbInformation

End Sub
Sub ResizeMMButtons()
    Dim wsMM As Worksheet
    Dim shp As Shape
    Dim btnText As String
    Dim leftStart As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    Dim btnGap As Double
    Dim positioned As Long
    
    Set wsMM = ThisWorkbook.Sheets("Meeting Minutes")
    
    btnWidth = 140
    btnHeight = 22
    btnGap = 6
    leftStart = wsMM.Cells(2, 5).Left
    btnTop = wsMM.Cells(2, 1).Top + 4
    positioned = 0
    
    Application.ScreenUpdating = False
    
    For Each shp In wsMM.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoFormControl Then
            btnText = ""
            On Error Resume Next
            btnText = LCase(Trim(shp.TextFrame.Characters.Text))
            On Error GoTo 0
            
            shp.Width = btnWidth
            shp.Height = btnHeight
            shp.Top = btnTop
            
            If InStr(btnText, "start") > 0 Then
                shp.Left = leftStart
                shp.Fill.ForeColor.RGB = RGB(46, 114, 182)
                shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
                shp.TextFrame.Characters.Font.bold = True
                positioned = positioned + 1
            ElseIf InStr(btnText, "end meeting") > 0 Then
                shp.Left = leftStart + btnWidth + btnGap
                shp.Fill.ForeColor.RGB = RGB(237, 125, 49)
                shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
                shp.TextFrame.Characters.Font.bold = True
                positioned = positioned + 1
            ElseIf InStr(btnText, "hide") > 0 Or InStr(btnText, "show done") > 0 Then
                shp.Left = leftStart + (btnWidth + btnGap) * 2
                shp.Fill.ForeColor.RGB = RGB(155, 187, 220)
                shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
                shp.TextFrame.Characters.Font.bold = True
                positioned = positioned + 1
            End If
        End If
    Next shp
    
    Application.ScreenUpdating = True
    
    MsgBox positioned & " button(s) repositioned and resized.", vbInformation
End Sub
