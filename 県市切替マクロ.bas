Attribute VB_Name = "KenShiSwitch"
Option Explicit

' ================================================================
' Ken/Shi sheet switch macro - VBE import safe version
'
' This module intentionally avoids Japanese identifiers and Japanese
' string literals. Japanese sheet names and UI labels are built with
' ChrW code points so the code will not break when imported by VBE
' on older Excel environments.
'
' User workflow:
'   1. Import this .bas file into VBE.
'   2. Run Setup_KenShi once from Alt+F8.
'   3. Use D3 on the input sheet.
'
' D3 = Ken  : show prefecture sheets, hide city sheets.
' D3 = Shi  : show city sheets, hide prefecture sheets.
' D3 = blank: show both prefecture and city sheets.
' ================================================================

Private Const INPUT_CELL As String = "D3"
Private Const NAV_RANGE As String = "H3:K30"

' ----- Unicode helpers -----
Private Function W(ByVal code As Long) As String
    If code > 32767 Then code = code - 65536
    W = ChrW$(code)
End Function

Private Function U(ParamArray codes() As Variant) As String
    Dim i As Long
    Dim s As String
    For i = LBound(codes) To UBound(codes)
        s = s & W(CLng(codes(i)))
    Next i
    U = s
End Function

' ----- Japanese text values built safely -----
Private Function InputSheetName() As String
    InputSheetName = U(&H65BD, &H5DE5, &H8A08, &H753B, &H66F8, &HFF08, &H306F, &H3058, &H3081, &H306B, &H8A18, &H5165, &HFF09)
End Function

Private Function KenText() As String
    KenText = U(&H770C)
End Function

Private Function ShiText() As String
    ShiText = U(&H5E02)
End Function

Private Function KenDobokuText() As String
    KenDobokuText = U(&H770C, &H571F, &H6728)
End Function

Private Function ShiyakushoText() As String
    ShiyakushoText = U(&H5E02, &H5F79, &H6240)
End Function

Private Function NavTitleText() As String
    NavTitleText = U(&H63D0, &H51FA, &H66F8, &H985E, &H30CA, &H30D3)
End Function

Private Function CurrentModePrefix() As String
    CurrentModePrefix = U(&H73FE, &H5728, &H306E, &H63D0, &H51FA, &H5148, &H003A, &H0020)
End Function

Private Function BlankModeText() As String
    BlankModeText = U(&H672A, &H9078, &H629E, &HFF08, &H770C, &H30FB, &H5E02, &H0020, &H4E21, &H65B9, &H8868, &H793A, &HFF09)
End Function

Private Function KenHeaderText() As String
    KenHeaderText = U(&H3010, &H770C, &H571F, &H6728, &H3011)
End Function

Private Function ShiHeaderText() As String
    ShiHeaderText = U(&H3010, &H5E02, &H5F79, &H6240, &H3011)
End Function

' ----- Sheet lists -----
Public Function KenSheetNames() As Variant
    KenSheetNames = Array( _
        U(&H76EE, &H6B21, &HFF08, &H65BD, &H5DE5, &H8A08, &H753B, &H66F8, &H002E, &H770C, &H571F, &H6728, &HFF09), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H6559, &H80B2, &H4E88, &H5B9A, &H8868), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &H7D44, &H7E54, &H8868), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &HFF11), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &HFF12), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &HFF13), _
        U(&H0039, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &HFF13, &H0020, &H0028, &H6025, &H50BE, &H659C, &H0020, &H0029), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &HFF13, &H002D, &H0032), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &H0034) _
    )
End Function

Public Function ShiSheetNames() As Variant
    ShiSheetNames = Array( _
        U(&H76EE, &H6B21, &HFF08, &H65BD, &H5DE5, &H8A08, &H753B, &H002E, &H5E02, &H5F79, &H6240, &HFF09), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H6559, &H80B2, &H4E88, &H5B9A, &H8868, &H0020, &H0028, &H0032, &H0029), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &H7D44, &H7E54, &H8868, &H0020, &H0028, &H0032, &H0029), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &HFF11, &H0020, &H0028, &H0032, &H0029), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &HFF12, &H0020, &H0028, &H0032, &H0029), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &HFF13, &H0020, &H0028, &H0032, &H0029), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &HFF13, &H002D, &H0032, &H0020, &H0028, &H0032, &H0029), _
        U(&H0031, &H0030, &HFF08, &H0039, &HFF09, &H002E, &H5B89, &H5168, &H7BA1, &H7406, &HFF30, &H0034, &H0020, &H0028, &H0032, &H0029) _
    )
End Function

' ================================================================
' Public macros
' ================================================================
Public Sub Setup_KenShi()
    Dim ws As Worksheet
    Set ws = SheetByName(InputSheetName())
    If ws Is Nothing Then
        MsgBox "Input sheet was not found: " & InputSheetName(), vbExclamation, "Setup_KenShi"
        Exit Sub
    End If

    If HasNonNavigationData(ws) Then
        If MsgBox("The navigation area " & NAV_RANGE & " already has data." & vbCrLf & _
                  "This macro will overwrite that area." & vbCrLf & vbCrLf & _
                  "Continue setup?", vbYesNo + vbExclamation, "Setup_KenShi") <> vbYes Then
            Exit Sub
        End If
    End If

    On Error Resume Next
    ws.Range(INPUT_CELL).Validation.Delete
    On Error GoTo 0

    With ws.Range(INPUT_CELL).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=KenText() & "," & ShiText()
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Input error"
        .ErrorMessage = "Choose Ken or Shi. Leave blank to show both."
        .ShowError = True
    End With

    Dim eventResult As String
    eventResult = WriteWorksheetChange(ws)

    KenShiSheetSwitch CStr(ws.Range(INPUT_CELL).Value)

    MsgBox "Setup completed." & vbCrLf & _
           "Event: " & eventResult & vbCrLf & vbCrLf & _
           "D3 = Ken  : prefecture sheets only" & vbCrLf & _
           "D3 = Shi  : city sheets only" & vbCrLf & _
           "D3 = blank: show both", vbInformation, "Setup_KenShi"
End Sub

Public Sub Check_KenShiSettings()
    Dim problems As String
    Dim ws As Worksheet
    Set ws = SheetByName(InputSheetName())

    If ws Is Nothing Then
        problems = problems & "- Input sheet not found: " & InputSheetName() & vbCrLf
    Else
        Dim mode As String
        mode = NormalizeMode(CStr(ws.Range(INPUT_CELL).Value))
        If Not (mode = "" Or mode = KenText() Or mode = ShiText()) Then
            problems = problems & "- D3 value is invalid. Use Ken, Shi, or blank." & vbCrLf
        End If

        If ws.Visible <> xlSheetVisible Then
            problems = problems & "- Input sheet is hidden." & vbCrLf
        End If

        problems = problems & CheckWorksheetChange(ws)
    End If

    problems = problems & CheckSheetList("Ken", KenSheetNames())
    problems = problems & CheckSheetList("Shi", ShiSheetNames())
    problems = problems & CheckDuplicateSheetNames(KenSheetNames(), ShiSheetNames())

    If Len(problems) = 0 Then
        MsgBox "Check completed. No problems were found.", vbInformation, "Check_KenShiSettings"
    Else
        MsgBox "Problems found:" & vbCrLf & vbCrLf & problems, vbExclamation, "Check_KenShiSettings"
    End If
End Sub

Public Sub KenShiSheetSwitch(ByVal valueText As String)
    On Error GoTo CleanFail

    Dim mode As String
    mode = NormalizeMode(valueText)

    Dim showList As Variant
    Dim hideList As Variant
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    Select Case mode
        Case KenText()
            showList = KenSheetNames()
            hideList = ShiSheetNames()
        Case ShiText()
            showList = ShiSheetNames()
            hideList = KenSheetNames()
        Case ""
            ' Blank means show both.
        Case Else
            MsgBox "Choose Ken or Shi. Leave blank to show both.", vbExclamation, "KenShiSheetSwitch"
            Exit Sub
    End Select

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If mode = "" Then
        ShowSheets KenSheetNames()
        ShowSheets ShiSheetNames()
        BuildNavigation ""
        GoTo CleanExit
    End If

    ShowSheets showList

    Dim visibleBefore As Long
    Dim hideExpected As Long
    visibleBefore = VisibleSheetCount()
    hideExpected = HideCandidateCount(hideList)

    If visibleBefore - hideExpected < 1 Then
        MsgBox "Switch was stopped because all sheets would become hidden.", vbExclamation, "KenShiSheetSwitch"
        GoTo CleanExit
    End If

    HideSheets hideList
    BuildNavigation mode

CleanExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

CleanFail:
    MsgBox "An error occurred while switching sheets:" & vbCrLf & Err.Description, vbExclamation, "KenShiSheetSwitch"
    Resume CleanExit
End Sub

' ================================================================
' Internal helpers
' ================================================================
Private Function NormalizeMode(ByVal valueText As String) As String
    Dim s As String
    s = valueText
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, ChrW$(&H3000), "")
    NormalizeMode = Trim$(s)
End Function

Private Function SheetByName(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function HasNonNavigationData(ByVal ws As Worksheet) As Boolean
    Dim navArea As Range
    Set navArea = ws.Range(NAV_RANGE)
    If Application.WorksheetFunction.CountA(navArea) = 0 Then
        HasNonNavigationData = False
    Else
        HasNonNavigationData = (CStr(navArea.Cells(1, 1).Value) <> NavTitleText())
    End If
End Function

Private Function WriteWorksheetChange(ByVal ws As Worksheet) As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object
    Dim codeText As String

    codeText = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
               "    If Intersect(Target, Me.Range(""" & INPUT_CELL & """)) Is Nothing Then Exit Sub" & vbCrLf & _
               "    If Target.Cells.Count > 1 Then Exit Sub" & vbCrLf & _
               "    KenShiSheetSwitch CStr(Target.Value)" & vbCrLf & _
               "End Sub"

    On Error GoTo TrustNG
    Set vbProj = ThisWorkbook.VBProject
    Set vbComp = vbProj.VBComponents(ws.CodeName)
    Set codeMod = vbComp.CodeModule

    On Error Resume Next
    Dim startLine As Long
    startLine = codeMod.ProcStartLine("Worksheet_Change", 0)
    If startLine > 0 Then
        Dim lineCount As Long
        lineCount = codeMod.ProcCountLines("Worksheet_Change", 0)
        codeMod.DeleteLines startLine, lineCount
    End If
    On Error GoTo TrustNG

    codeMod.InsertLines codeMod.CountOfLines + 1, codeText
    WriteWorksheetChange = "OK"
    Exit Function

TrustNG:
    MsgBox "Could not write Worksheet_Change automatically." & vbCrLf & _
           "If D3 does not switch sheets, paste this code into the input sheet module:" & vbCrLf & vbCrLf & _
           codeText, vbExclamation, "Setup_KenShi"
    WriteWorksheetChange = "Manual setup required"
End Function

Private Function CheckWorksheetChange(ByVal ws As Worksheet) As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object
    Dim startLine As Long
    Dim lineCount As Long
    Dim text As String

    On Error GoTo TrustNG
    Set vbProj = ThisWorkbook.VBProject
    Set vbComp = vbProj.VBComponents(ws.CodeName)
    Set codeMod = vbComp.CodeModule

    On Error Resume Next
    startLine = codeMod.ProcStartLine("Worksheet_Change", 0)
    Err.Clear
    On Error GoTo TrustNG

    If startLine <= 0 Then
        CheckWorksheetChange = "- Worksheet_Change was not found. Run Setup_KenShi again." & vbCrLf
        Exit Function
    End If

    lineCount = codeMod.ProcCountLines("Worksheet_Change", 0)
    text = codeMod.Lines(startLine, lineCount)

    If InStr(text, "KenShiSheetSwitch") = 0 Then
        CheckWorksheetChange = "- Worksheet_Change does not call KenShiSheetSwitch." & vbCrLf
    End If
    Exit Function

TrustNG:
    CheckWorksheetChange = "- Could not check Worksheet_Change. Trust access may be disabled." & vbCrLf
End Function

Private Function CheckSheetList(ByVal label As String, ByVal names As Variant) As String
    Dim i As Long
    Dim missing As String

    For i = LBound(names) To UBound(names)
        If SheetByName(CStr(names(i))) Is Nothing Then
            missing = missing & "    " & CStr(names(i)) & vbCrLf
        End If
    Next i

    If Len(missing) > 0 Then
        CheckSheetList = "- Missing " & label & " sheets:" & vbCrLf & missing
    End If
End Function

Private Function CheckDuplicateSheetNames(ByVal a As Variant, ByVal b As Variant) As String
    Dim i As Long
    Dim j As Long
    Dim dup As String

    For i = LBound(a) To UBound(a)
        For j = LBound(b) To UBound(b)
            If StrComp(CStr(a(i)), CStr(b(j)), vbTextCompare) = 0 Then
                dup = dup & "    " & CStr(a(i)) & vbCrLf
                Exit For
            End If
        Next j
    Next i

    If Len(dup) > 0 Then
        CheckDuplicateSheetNames = "- Duplicate sheet names in both lists:" & vbCrLf & dup
    End If
End Function

Private Sub BuildNavigation(ByVal mode As String)
    Dim ws As Worksheet
    Set ws = SheetByName(InputSheetName())
    If ws Is Nothing Then Exit Sub

    Dim navArea As Range
    Set navArea = ws.Range(NAV_RANGE)
    navArea.ClearContents

    Dim idx As Long
    Dim link As Hyperlink
    For idx = ws.Hyperlinks.Count To 1 Step -1
        Set link = ws.Hyperlinks(idx)
        If Not Intersect(link.Range, navArea) Is Nothing Then link.Delete
    Next idx

    Dim topRow As Long
    Dim leftCol As Long
    Dim bottomRow As Long
    topRow = navArea.Row
    leftCol = navArea.Column
    bottomRow = navArea.Row + navArea.Rows.Count - 1

    ws.Cells(topRow, leftCol).Value = NavTitleText()

    Dim row As Long
    row = topRow + 3

    Select Case mode
        Case KenText()
            ws.Cells(topRow + 1, leftCol).Value = CurrentModePrefix() & KenDobokuText()
            row = AddNavList(ws, row, leftCol, bottomRow, KenSheetNames())
        Case ShiText()
            ws.Cells(topRow + 1, leftCol).Value = CurrentModePrefix() & ShiyakushoText()
            row = AddNavList(ws, row, leftCol, bottomRow, ShiSheetNames())
        Case ""
            ws.Cells(topRow + 1, leftCol).Value = CurrentModePrefix() & BlankModeText()
            ws.Cells(row, leftCol).Value = KenHeaderText()
            row = row + 1
            row = AddNavList(ws, row, leftCol, bottomRow, KenSheetNames())
            If row <= bottomRow Then
                ws.Cells(row, leftCol).Value = ShiHeaderText()
                row = row + 1
                row = AddNavList(ws, row, leftCol, bottomRow, ShiSheetNames())
            End If
    End Select
End Sub

Private Function AddNavList(ByVal ws As Worksheet, ByVal startRow As Long, ByVal leftCol As Long, ByVal bottomRow As Long, ByVal names As Variant) As Long
    Dim row As Long
    Dim i As Long
    Dim target As Worksheet
    row = startRow

    For i = LBound(names) To UBound(names)
        If row > bottomRow Then Exit For
        Set target = SheetByName(CStr(names(i)))
        If Not target Is Nothing Then
            ws.Hyperlinks.Add _
                Anchor:=ws.Cells(row, leftCol), _
                Address:="", _
                SubAddress:="'" & Replace(target.Name, "'", "''") & "'!A1", _
                TextToDisplay:=target.Name
            row = row + 1
        End If
    Next i

    AddNavList = row
End Function

Private Function ShowSheets(ByVal names As Variant) As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim missing As Long

    For i = LBound(names) To UBound(names)
        Set ws = SheetByName(CStr(names(i)))
        If ws Is Nothing Then
            missing = missing + 1
        Else
            ws.Visible = xlSheetVisible
        End If
    Next i

    ShowSheets = missing
End Function

Private Function HideSheets(ByVal names As Variant) As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim missing As Long

    For i = LBound(names) To UBound(names)
        Set ws = SheetByName(CStr(names(i)))
        If ws Is Nothing Then
            missing = missing + 1
        ElseIf ws.Name <> InputSheetName() Then
            ws.Visible = xlSheetHidden
        End If
    Next i

    HideSheets = missing
End Function

Private Function HideCandidateCount(ByVal names As Variant) As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim count As Long

    For i = LBound(names) To UBound(names)
        Set ws = SheetByName(CStr(names(i)))
        If Not ws Is Nothing Then
            If ws.Name <> InputSheetName() And ws.Visible = xlSheetVisible Then
                count = count + 1
            End If
        End If
    Next i

    HideCandidateCount = count
End Function

Private Function VisibleSheetCount() As Long
    Dim ws As Worksheet
    Dim count As Long

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then count = count + 1
    Next ws

    VisibleSheetCount = count
End Function
