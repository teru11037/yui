Attribute VB_Name = "県市切替"
'==================================================================
'  県/市シート 自動表示切替マクロ
'  ────────────────────────────────────────────────
'  最初のシート「施工計画書（はじめに記入）」のD3セルに
'  ▼ドロップダウンで「県」または「市」を選ぶと、
'  下で指定したシートが自動で 表示/非表示 されます。
'
'  【取り込み手順】
'  1) Excelで対象ブックを開き Alt+F11 で VBE を開く
'  2) メニュー「ファイル」→「ファイルのインポート」で
'     この 県市切替マクロ.bas を選択
'  3) プロジェクト内の「Sheet1 (施工計画書（はじめに記入）)」を
'     ダブルクリックして開き、末尾に下の Worksheet_Change を貼り付け
'  4) VBE のイミディエイトウィンドウ (Ctrl+G) で
'        Call 初期設定
'     と入力して Enter を押す（D3に▼が付きます）
'  5) ブックを「マクロ有効ブック (.xlsm) または .xls」で保存
'==================================================================
Option Explicit

' 入力セル（最初のシートのこのセルが切替トリガー）
Private Const 入力セル As String = "D3"


'------------------------------------------------------------------
'  ▼ ここを編集してください ▼
'   「県」と入力したときに「表示」するシート名を列挙します。
'   （同時に「市」のときは ここに書かれたシートは「非表示」）
'------------------------------------------------------------------
Public Function 県シート一覧() As Variant
    県シート一覧 = Array( _
        "目次（施工計画書.県土木）", _
        "10（9）.安全教育予定表", _
        "10（9）.安全管理組織表", _
        "10（9）.安全管理Ｐ１", _
        "10（9）.安全管理Ｐ２", _
        "10（9）.安全管理Ｐ３", _
        "9.安全管理Ｐ３ (急傾斜 )", _
        "10（9）.安全管理Ｐ３-2", _
        "10（9）.安全管理Ｐ4" _
    )
End Function

'------------------------------------------------------------------
'  ▼ ここを編集してください ▼
'   「市」と入力したときに「表示」するシート名を列挙します。
'   （同時に「県」のときは ここに書かれたシートは「非表示」）
'------------------------------------------------------------------
Public Function 市シート一覧() As Variant
    市シート一覧 = Array( _
        "目次（施工計画.市役所）", _
        "10（9）.安全教育予定表 (2)", _
        "10（9）.安全管理組織表 (2)", _
        "10（9）.安全管理Ｐ１ (2)", _
        "10（9）.安全管理Ｐ２ (2)", _
        "10（9）.安全管理Ｐ３ (2)", _
        "10（9）.安全管理Ｐ３-2 (2)", _
        "10（9）.安全管理Ｐ4 (2)" _
    )
End Function


'==================================================================
'  ▼ 初期設定（D3にドロップダウンを付ける。一度だけ実行）
'==================================================================
Public Sub 初期設定()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(1)   ' 最初のシート

    With sh.Range(入力セル).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="県,市"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "「県」または「市」を選択してください。"
        .ShowError = True
    End With

    MsgBox "セットアップ完了！" & vbCrLf & _
           "「" & sh.Name & "」シートの " & 入力セル & " で「県/市」を選択してください。", _
           vbInformation
End Sub


'==================================================================
'  ▼ 切替本体（基本的に編集不要）
'==================================================================
Public Sub 県市シート切替(ByVal 値 As String)
    Dim 表示 As Variant, 非表示 As Variant

    Select Case Trim$(値)
        Case "県"
            表示 = 県シート一覧()
            非表示 = 市シート一覧()
        Case "市"
            表示 = 市シート一覧()
            非表示 = 県シート一覧()
        Case ""
            Exit Sub  ' 空欄は何もしない
        Case Else
            MsgBox "「県」または「市」を選択してください。", vbExclamation, "入力エラー"
            Exit Sub
    End Select

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    シート表示切替 表示, True
    シート表示切替 非表示, False

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub シート表示切替(ByVal 名前一覧 As Variant, ByVal 表示する As Boolean)
    Dim i As Long, sh As Worksheet
    For i = LBound(名前一覧) To UBound(名前一覧)
        Set sh = Nothing
        On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(CStr(名前一覧(i)))
        On Error GoTo 0
        If Not sh Is Nothing Then
            If 表示する Then
                sh.Visible = xlSheetVisible
            Else
                ' 最後の1枚を非表示にしようとするとエラーになるので保護
                If 表示中シート数() > 1 Then
                    sh.Visible = xlSheetHidden
                End If
            End If
        End If
    Next i
End Sub

Private Function 表示中シート数() As Long
    Dim sh As Worksheet, c As Long
    For Each sh In ThisWorkbook.Worksheets
        If sh.Visible = xlSheetVisible Then c = c + 1
    Next sh
    表示中シート数 = c
End Function


'==================================================================
'  【シートモジュール用イベント】
'  ────────────────────────────────────────────────
'  ※ 下のコードは「標準モジュール」ではなく、
'    Sheet1（施工計画書（はじめに記入））のコードに貼り付けます。
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     If Intersect(Target, Me.Range("D3")) Is Nothing Then Exit Sub
'     If Target.Cells.CountLarge > 1 Then Exit Sub
'     県市シート切替 CStr(Target.Value)
' End Sub
'==================================================================
