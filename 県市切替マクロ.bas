Attribute VB_Name = "県市切替"
'==================================================================
'  県/市シート 自動表示切替マクロ
'  ────────────────────────────────────────────────
'  入力シート「施工計画書（はじめに記入）」の D3 セルに
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

' 入力シート（この名前のシートを対象とする）
Private Const 入力シート名 As String = "施工計画書（はじめに記入）"

' 入力セル（上記シートのこのセルが切替トリガー）
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
'    (1) 入力シートのD3にドロップダウン(県/市)を設置
'    (2) 入力シートの Worksheet_Change イベントを自動で書き込み
'==================================================================
Public Sub 初期設定()
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(入力シート名)
    On Error GoTo 0
    If sh Is Nothing Then
        MsgBox "対象シート「" & 入力シート名 & "」が見つかりません。" & vbCrLf & _
               "シート名を確認するか、定数『入力シート名』を実際のシート名に合わせて変更してください。", _
               vbExclamation, "シートが見つかりません"
        Exit Sub
    End If

    ' --- (0) H3:K30 に既存データがあれば上書き確認（既存ナビは除外） ---
    Dim ナビ領域 As Range
    Set ナビ領域 = sh.Range("H3:K30")
    If Application.WorksheetFunction.CountA(ナビ領域) > 0 Then
        If CStr(sh.Range("H3").Value) <> "提出書類ナビ" Then
            If MsgBox("入力シートの H3:K30 に既存データがあります。" & vbCrLf & _
                      "このマクロは H3:K30 を『提出書類ナビ』用に使うため、" & vbCrLf & _
                      "切替の度に上書きされます。" & vbCrLf & vbCrLf & _
                      "セットアップを続行しますか？", _
                      vbYesNo + vbExclamation, "H3:K30 上書きの確認") <> vbYes Then
                Exit Sub
            End If
        End If
    End If

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

    ' D3 が既に県/市なら ナビも初回生成
    Dim 現値 As String
    現値 = 値正規化(CStr(sh.Range(入力セル).Value))
    ' D3 が空欄/県/市 のいずれでもナビ初回生成
    ナビ再生成 現値

    Dim msg As String
    msg = "▼ ドロップダウン設置: 完了 (" & sh.Name & "!" & 入力セル & ")" & vbCrLf
    msg = msg & "▼ 自動切替イベント: " & 結果 & vbCrLf & vbCrLf
    msg = msg & "→ D3 の▼から「県」または「市」を選んでください。" & vbCrLf & vbCrLf
    msg = msg & "（提出前・配布前に、Alt+F8 から『設定チェック』を実行してシート設定の妥当性を確認してください）"
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
            "    If Target.Cells.Count > 1 Then Exit Sub" & vbCrLf & _
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
    ' Worksheet_Change の自動書込みに失敗 → 原因を断定せず情報提示
    Dim エラー番号 As Long, エラー説明 As String
    エラー番号 = Err.Number
    エラー説明 = Err.Description
    MsgBox "Worksheet_Change イベントの自動書込みに失敗しました。" & vbCrLf & _
           "エラー番号: " & エラー番号 & vbCrLf & _
           "詳細: " & エラー説明 & vbCrLf & vbCrLf & _
           "原因として多いのは『VBAプロジェクト オブジェクト モデルへのアクセスを信頼する』" & vbCrLf & _
           "が無効な場合です。以下を確認してください:" & vbCrLf & _
           "  [ファイル]→[オプション]→[トラストセンター]" & vbCrLf & _
           "  →[トラストセンターの設定]→[マクロの設定]" & vbCrLf & _
           "  →『VBAプロジェクト オブジェクト モデルへのアクセスを信頼する』にチェック" & vbCrLf & vbCrLf & _
           "上記を確認しても解決しない場合は、以下を Sheet1 (" & sh.Name & ") に" & vbCrLf & _
           "手動で貼り付けてください:" & vbCrLf & vbCrLf & _
           "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
           "    If Intersect(Target, Me.Range(""D3"")) Is Nothing Then Exit Sub" & vbCrLf & _
           "    If Target.Cells.Count > 1 Then Exit Sub" & vbCrLf & _
           "    県市シート切替 CStr(Target.Value)" & vbCrLf & _
           "End Sub", vbExclamation, "イベント書込みに失敗"
    イベントコード書込 = "未設定（上記の案内を参照）"
End Function


'------------------------------------------------------------------
'  入力値の正規化
'   半角/全角スペース・改行・タブを除去して比較しやすくする
'------------------------------------------------------------------
Private Function 値正規化(ByVal 値 As String) As String
    Dim s As String
    s = 値
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, "　", "")  ' 全角スペース
    値正規化 = Trim$(s)
End Function

'==================================================================
'  ▼ 設定チェック ― 提出前/配布前に実行して設定の妥当性を確認
'==================================================================
Public Sub 設定チェック()
    Dim 問題 As String
    Dim 入力sh As Worksheet
    Dim sh As Worksheet
    Dim i As Long, j As Long

    ' (1) 入力シートの存在
    On Error Resume Next
    Set 入力sh = ThisWorkbook.Worksheets(入力シート名)
    On Error GoTo 0
    If 入力sh Is Nothing Then
        問題 = 問題 & "・入力シート「" & 入力シート名 & "」が存在しません。" & vbCrLf
    Else
        ' (2) 入力セル と その値
        Dim セル As Range
        On Error Resume Next
        Set セル = 入力sh.Range(入力セル)
        On Error GoTo 0
        If セル Is Nothing Then
            問題 = 問題 & "・入力セル「" & 入力セル & "」が解釈できません。" & vbCrLf
        Else
            Dim 値 As String
            値 = 値正規化(CStr(セル.Value))
            Select Case 値
                Case "", "県", "市"
                    ' OK
                Case Else
                    問題 = 問題 & "・入力セル " & 入力セル & " の値「" & セル.Value & _
                             "」が想定外です（空欄/県/市 のいずれかにしてください）。" & vbCrLf
            End Select
        End If

        ' (6) 入力シートが表示状態か
        If 入力sh.Visible <> xlSheetVisible Then
            問題 = 問題 & "・入力シート「" & 入力シート名 & "」が非表示です。表示状態にしてください。" & vbCrLf
        End If
    End If

    ' (3) 県シート一覧 の存在チェック
    Dim 県名 As Variant, 県不在 As String
    県名 = 県シート一覧()
    For i = LBound(県名) To UBound(県名)
        Set sh = Nothing
        On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(CStr(県名(i)))
        On Error GoTo 0
        If sh Is Nothing Then
            県不在 = 県不在 & "    " & CStr(県名(i)) & vbCrLf
        End If
    Next i
    If Len(県不在) > 0 Then
        問題 = 問題 & "・県シート一覧 に実在しないシート:" & vbCrLf & 県不在
    End If

    ' (4) 市シート一覧 の存在チェック
    Dim 市名 As Variant, 市不在 As String
    市名 = 市シート一覧()
    For i = LBound(市名) To UBound(市名)
        Set sh = Nothing
        On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(CStr(市名(i)))
        On Error GoTo 0
        If sh Is Nothing Then
            市不在 = 市不在 & "    " & CStr(市名(i)) & vbCrLf
        End If
    Next i
    If Len(市不在) > 0 Then
        問題 = 問題 & "・市シート一覧 に実在しないシート:" & vbCrLf & 市不在
    End If

    ' (5) 県/市 両方に重複するシート名
    Dim 重複 As String
    For i = LBound(県名) To UBound(県名)
        For j = LBound(市名) To UBound(市名)
            If StrComp(CStr(県名(i)), CStr(市名(j)), vbTextCompare) = 0 Then
                重複 = 重複 & "    " & CStr(県名(i)) & vbCrLf
                Exit For
            End If
        Next j
    Next i
    If Len(重複) > 0 Then
        問題 = 問題 & "・県/市 両方の一覧に存在するシート:" & vbCrLf & 重複
    End If

    ' (7) 入力シートの Worksheet_Change イベント確認
    If Not 入力sh Is Nothing Then
        問題 = 問題 & イベント存在確認(入力sh)
    End If

    ' 結果表示
    If Len(問題) = 0 Then
        MsgBox "設定チェック完了。問題は見つかりませんでした。", vbInformation, "設定チェック"
    Else
        MsgBox "以下の問題が見つかりました:" & vbCrLf & vbCrLf & 問題, vbExclamation, "設定チェック"
    End If
End Sub

'------------------------------------------------------------------
'  入力シートの Worksheet_Change イベントを VBProject 経由で確認
'   戻り値: 問題があればその記述（複数行可）、無ければ ""
'   VBProject にアクセスできない場合は『未確認』として警告に含める
'------------------------------------------------------------------
Private Function イベント存在確認(ByVal sh As Worksheet) As String
    Dim VBProj As Object, VBComp As Object, CodeMod As Object
    Dim 開始 As Long, 行数 As Long, 本文 As String
    Dim 結果 As String

    On Error GoTo TrustNG
    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(sh.CodeName)
    Set CodeMod = VBComp.CodeModule

    On Error Resume Next
    開始 = CodeMod.ProcStartLine("Worksheet_Change", 0)
    Err.Clear
    On Error GoTo TrustNG

    If 開始 <= 0 Then
        結果 = "・入力シートに Worksheet_Change イベントが見つかりません。『初期設定』を再実行してください。" & vbCrLf
    Else
        On Error Resume Next
        行数 = CodeMod.ProcCountLines("Worksheet_Change", 0)
        本文 = CodeMod.Lines(開始, 行数)
        Err.Clear
        On Error GoTo TrustNG
        If InStr(本文, "県市シート切替") = 0 Then
            結果 = "・Worksheet_Change イベント内に 県市シート切替 の呼び出しが見つかりません。" & vbCrLf
        End If
    End If

    イベント存在確認 = 結果
    Exit Function

TrustNG:
    イベント存在確認 = "・Worksheet_Change の確認に失敗（Err " & Err.Number & ": " & Err.Description & "）。" & vbCrLf & _
                      "    『VBA プロジェクト オブジェクト モデルへのアクセスを信頼する』が無効の可能性があります（イベント有無は未確認）。" & vbCrLf
End Function


'==================================================================
'  ▼ 提出書類ナビ ― 入力シート H3:K30 にジャンプ用一覧を再生成
'==================================================================
Private Sub ナビ再生成(ByVal モード As String)
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(入力シート名)
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    ' 既存ナビ領域をクリア
    Dim 範囲 As Range
    Set 範囲 = sh.Range("H3:K30")
    範囲.ClearContents
    Dim hl As Hyperlink, k As Long
    For k = sh.Hyperlinks.Count To 1 Step -1
        Set hl = sh.Hyperlinks(k)
        If Not Intersect(hl.Range, 範囲) Is Nothing Then hl.Delete
    Next k

    ' モードに応じてナビ内容を生成
    Dim 行 As Long
    Select Case モード
        Case "県"
            sh.Range("H3").Value = "提出書類ナビ"
            sh.Range("H4").Value = "現在の提出先: 県土木"
            ナビリスト書込 sh, 6, 県シート一覧()
        Case "市"
            sh.Range("H3").Value = "提出書類ナビ"
            sh.Range("H4").Value = "現在の提出先: 市役所"
            ナビリスト書込 sh, 6, 市シート一覧()
        Case ""
            ' 空欄 = 県・市 両方を見出し付きで表示
            sh.Range("H3").Value = "提出書類ナビ"
            sh.Range("H4").Value = "現在の提出先: 未選択（県・市 両方表示）"
            行 = 6
            sh.Cells(行, 8).Value = "【県土木】"
            行 = ナビリスト書込(sh, 行 + 1, 県シート一覧())
            If 行 <= 30 Then
                sh.Cells(行, 8).Value = "【市役所】"
                ナビリスト書込 sh, 行 + 1, 市シート一覧()
            End If
        Case Else
            ' 想定外はクリアだけで終了
    End Select
End Sub

'------------------------------------------------------------------
'  ナビへハイパーリンク一覧を書き込む共通処理
'   - 実在しないシートは無視
'   - 行 30 を超えたら打ち切り
'   - 戻り値: 次に書き込める行番号
'------------------------------------------------------------------
Private Function ナビリスト書込(ByVal sh As Worksheet, ByVal 開始行 As Long, ByVal 名前一覧 As Variant) As Long
    Dim 行 As Long, i As Long, 対象 As Worksheet
    行 = 開始行
    For i = LBound(名前一覧) To UBound(名前一覧)
        If 行 > 30 Then Exit For
        Set 対象 = Nothing
        On Error Resume Next
        Set 対象 = ThisWorkbook.Worksheets(CStr(名前一覧(i)))
        On Error GoTo 0
        If Not 対象 Is Nothing Then
            sh.Hyperlinks.Add _
                Anchor:=sh.Cells(行, 8), _
                Address:="", _
                SubAddress:="'" & 対象.Name & "'!A1", _
                TextToDisplay:=対象.Name
            行 = 行 + 1
        End If
    Next i
    ナビリスト書込 = 行
End Function


'==================================================================
'  ▼ 切替本体（基本的に編集不要）
'==================================================================
Public Sub 県市シート切替(ByVal 値 As String)
    On Error GoTo CleanFail

    Dim 表示 As Variant, 非表示 As Variant

    Select Case 値正規化(値)
        Case "県"
            表示 = 県シート一覧()
            非表示 = 市シート一覧()
        Case "市"
            表示 = 市シート一覧()
            非表示 = 県シート一覧()
        Case ""
            ' 空欄 = 県・市 両方表示
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            Call シート表示(県シート一覧())
            Call シート表示(市シート一覧())
            ナビ再生成 ""
            GoTo CleanExit
        Case Else
            MsgBox "「県」または「市」を選択してください。", vbExclamation, "入力エラー"
            Exit Sub
    End Select

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' (1) 表示対象を先にすべて可視化（hide 前にカウントを確定させる）
    Call シート表示(表示)

    ' (2) 非表示後に最低1枚の可視シートが残るかを一括判定
    Dim 可視前 As Long, 非表示見込 As Long
    可視前 = 表示中シート数()
    非表示見込 = 非表示有効数(非表示)
    If 可視前 - 非表示見込 < 1 Then
        MsgBox "全シートが非表示になってしまうため、切替を中断しました。" & vbCrLf & _
               "県/市シート一覧の設定を確認してください。", vbExclamation, "切替中断"
        GoTo CleanExit
    End If

    ' (3) 非表示対象を hide（入力シートは常に保護）
    Call シート非表示(非表示)

    ' (4) 提出書類ナビを最新状態に再生成
    ナビ再生成 値正規化(値)

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "シート切替中にエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "切替エラー"
    Resume CleanExit
End Sub

' 名前一覧 のシートを可視化。戻り値は未検出シート数。
Private Function シート表示(ByVal 名前一覧 As Variant) As Long
    Dim i As Long, sh As Worksheet, 未検出 As Long
    For i = LBound(名前一覧) To UBound(名前一覧)
        Set sh = Nothing
        On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(CStr(名前一覧(i)))
        On Error GoTo 0
        If sh Is Nothing Then
            未検出 = 未検出 + 1
        Else
            sh.Visible = xlSheetVisible
        End If
    Next i
    シート表示 = 未検出
End Function

' 名前一覧 のシートを非表示化。入力シート(入力シート名)は常に保護。戻り値は未検出シート数。
Private Function シート非表示(ByVal 名前一覧 As Variant) As Long
    Dim i As Long, sh As Worksheet, 未検出 As Long
    For i = LBound(名前一覧) To UBound(名前一覧)
        Set sh = Nothing
        On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(CStr(名前一覧(i)))
        On Error GoTo 0
        If sh Is Nothing Then
            未検出 = 未検出 + 1
        ElseIf sh.Name <> 入力シート名 Then
            sh.Visible = xlSheetHidden
        End If
    Next i
    シート非表示 = 未検出
End Function

' 非表示候補のうち、実在しかつ入力シートでなく かつ 現在可視 のシート数
Private Function 非表示有効数(ByVal 名前一覧 As Variant) As Long
    Dim i As Long, sh As Worksheet, c As Long
    For i = LBound(名前一覧) To UBound(名前一覧)
        Set sh = Nothing
        On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(CStr(名前一覧(i)))
        On Error GoTo 0
        If Not sh Is Nothing Then
            If sh.Name <> 入力シート名 And sh.Visible = xlSheetVisible Then
                c = c + 1
            End If
        End If
    Next i
    非表示有効数 = c
End Function

Private Function 表示中シート数() As Long
    Dim sh As Worksheet, c As Long
    For Each sh In ThisWorkbook.Worksheets
        If sh.Visible = xlSheetVisible Then c = c + 1
    Next sh
    表示中シート数 = c
End Function
