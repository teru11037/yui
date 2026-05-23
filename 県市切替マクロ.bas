Attribute VB_Name = "県市切替"
'==================================================================
'  県/市シート 自動表示切替マクロ
'  ────────────────────────────────────────────────
'  最初のシート「施工計画書（はじめに記入）」の D3 セルに
'  ▼ドロップダウンで「県」または「市」を選ぶと、
'  下で指定したシートが自動で 表示/非表示 されます。
'
'  ─ 取り込み手順（たった2ステップ） ─
'   ① Excelで対象ブックを開き Alt+F11 → メニュー「ファイル」
'      →「ファイルのインポート」で この 県市切替マクロ.bas を選択
'   ② Excel側に戻り、Alt+F8 で「初期設定」を選んで実行
'        （またはVBEのイミディエイトウィンドウで Call 初期設定）
'
'   ★ ②で「VBAプロジェクトへのアクセス信頼設定」が必要な場合は
'     画面の指示に従って一度だけ設定してください。
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
'  ▼ 初期設定 ― 一度だけ実行してください
'    (1) 最初のシートのD3にドロップダウン(県/市)を設置
'    (2) 最初のシートの Worksheet_Change イベントを自動で書き込み
'==================================================================
Public Sub 初期設定()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(1)   ' 最初のシート

    ' --- (1) D3 に「県/市」のドロップダウンを設定 ---
    On Error Resume Next
    sh.Range(入力セル).Validation.Delete
    On Error GoTo 0
    With sh.Range(入力セル).Validation
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

    ' --- (2) Worksheet_Change イベントをシートモジュールに自動追加 ---
    Dim 結果 As String
    結果 = イベントコード書込(sh)

    Dim msg As String
    msg = "▼ ドロップダウン設置: 完了 (" & sh.Name & "!" & 入力セル & ")" & vbCrLf
    msg = msg & "▼ 自動切替イベント: " & 結果 & vbCrLf & vbCrLf
    msg = msg & "→ D3 の▼から「県」または「市」を選んでください。"
    MsgBox msg, vbInformation, "セットアップ"
End Sub


'------------------------------------------------------------------
'  シートモジュールに Worksheet_Change を書き込む
'   ※ VBAプロジェクトへのアクセス許可がないと例外になる
'      その場合は手動貼付の手順をメッセージで案内する
'------------------------------------------------------------------
Private Function イベントコード書込(ByVal sh As Worksheet) As String
    Dim VBProj As Object, VBComp As Object, CodeMod As Object
    Dim コード As String, 行 As Long

    コード = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
            "    If Intersect(Target, Me.Range(""" & 入力セル & """)) Is Nothing Then Exit Sub" & vbCrLf & _
            "    If Target.Cells.CountLarge > 1 Then Exit Sub" & vbCrLf & _
            "    県市シート切替 CStr(Target.Value)" & vbCrLf & _
            "End Sub"

    On Error GoTo TrustNG
    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(sh.CodeName)
    Set CodeMod = VBComp.CodeModule

    ' すでに同じ手続きがあれば消してから挿入
    On Error Resume Next
    Dim 既存 As Long
    既存 = CodeMod.ProcStartLine("Worksheet_Change", 0)
    If 既存 > 0 Then
        Dim 行数 As Long
        行数 = CodeMod.ProcCountLines("Worksheet_Change", 0)
        CodeMod.DeleteLines 既存, 行数
    End If
    On Error GoTo TrustNG

    行 = CodeMod.CountOfLines + 1
    CodeMod.InsertLines 行, コード
    イベントコード書込 = "完了（自動）"
    Exit Function

TrustNG:
    ' VBAプロジェクトへの信頼アクセスが無効 → 手動貼付の案内
    MsgBox "Excelの設定で『VBAプロジェクトへのアクセスを信頼する』を" & vbCrLf & _
           "有効化する必要があります。" & vbCrLf & vbCrLf & _
           "  [ファイル]→[オプション]→[トラストセンター]" & vbCrLf & _
           "  →[トラストセンターの設定]→[マクロの設定]" & vbCrLf & _
           "  →『VBAプロジェクト オブジェクト モデルへのアクセスを信頼する』にチェック" & vbCrLf & vbCrLf & _
           "設定後、再度『初期設定』マクロを実行してください。" & vbCrLf & vbCrLf & _
           "── 手動で貼り付ける場合 ──" & vbCrLf & _
           "VBEで Sheet1 (" & sh.Name & ") を開き、以下を貼り付けてください:" & vbCrLf & vbCrLf & _
           "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
           "    If Intersect(Target, Me.Range(""D3"")) Is Nothing Then Exit Sub" & vbCrLf & _
           "    If Target.Cells.CountLarge > 1 Then Exit Sub" & vbCrLf & _
           "    県市シート切替 CStr(Target.Value)" & vbCrLf & _
           "End Sub", vbExclamation, "VBA信頼アクセスが必要です"
    イベントコード書込 = "未設定（上記の案内を参照）"
End Function


'==================================================================
'  ▼ 切替本体（基本的に編集不要）
'==================================================================
Public Sub 県市シート切替(ByVal 値 As String)
    On Error GoTo CleanFail

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

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "シート切替中にエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "切替エラー"
    Resume CleanExit
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
