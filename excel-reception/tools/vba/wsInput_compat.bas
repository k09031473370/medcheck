Attribute VB_Name = "wsInput"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' ========================================================================================
' 選択パネル設定（会社名・氏名を大きな文字で一覧表示し、クリックで選択）
' ========================================================================================
Private Const PANEL_FIRST_COL As Long = 9      ' I列から
Private Const PANEL_COLS As Long = 5           ' 5列（I～M列）
Private Const PANEL_HEADER_ROW As Long = 5
Private Const PANEL_FIRST_ROW As Long = 6
Private Const PANEL_LAST_ROW As Long = 200
Private Const PANEL_MODE_CELL As String = "AB1"
Private Const PANEL_FONT_SIZE As Long = 16

' ========================================================================================
' 各種ボタン
' ========================================================================================

' ------------------------------------------------------------------------------
' 受付登録ボタンクリック時の処理
' メインロジックはmdReceptionモジュールに委譲
' ------------------------------------------------------------------------------
Public Sub ClickRegisterButton()
    Call RegisterReception
End Sub

' ------------------------------------------------------------------------------
' クリアボタンクリック時の処理
' ユーザーに確認メッセージを表示してから入力内容をクリア
' ------------------------------------------------------------------------------
Public Sub ClickClearButton()
    Dim result As VbMsgBoxResult
    
    ' 確認メッセージ
    result = MsgBox("入力内容をクリアしますか？", vbQuestion + vbYesNo, "確認")
    
    If result = vbYes Then
        Call ClearInputFields
        Me.Range("会社名").Select
    End If
    
    Application.EnableEvents = True
End Sub

' ------------------------------------------------------------------------------
' 本日のシートボタンクリック時の処理
' 今日の日付のシート（yyyymmdd形式）にジャンプ
' シートが存在しない場合は何もしない
' ------------------------------------------------------------------------------
Public Sub ClickTodaySheetButton()
    Dim todaySheetName As String
    Dim ws As Worksheet
    
    ' 本日の日付でシート名を生成（例：20251010）
    todaySheetName = Format(Date, "yyyymmdd")
    
    ' シートが存在するかチェック
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(todaySheetName)
    On Error GoTo 0
    
    ' シートが存在する場合のみアクティブ化
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A1").Select
    End If
End Sub

' ------------------------------------------------------------------------------
' 受診者番号(フリー3)で検索ボタンクリック時の処理
' InputBoxで番号を入力し、該当する受診者情報を会社名・氏名に設定
' ------------------------------------------------------------------------------
Public Sub ClickSearchByFree3Button()
    Dim inputValue As String
    Dim free3Value As String
    Dim foundCompany As String
    Dim foundPerson As String
    Dim result As VbMsgBoxResult
    
    ' InputBoxで番号入力を促す
    inputValue = InputBox( _
        "「フリー3」を入力してください。" & vbCrLf & _
        vbCrLf & _
        "※受診者マスタの「フリー3」列に登録されている番号（文字列）です。", _
        "フリー3で検索", _
        "")
    
    ' キャンセルまたは空白の場合は終了
    If inputValue = "" Then Exit Sub
    
    free3Value = Trim(inputValue)
    
    ' 受診者マスタから検索
    If Not SearchPersonByFree3(free3Value, foundCompany, foundPerson) Then
        MsgBox "フリー3「" & free3Value & "」は見つかりませんでした。" & vbCrLf & _
               vbCrLf & _
               "受診者マスタの「フリー3」列を確認してください。", _
               vbExclamation, "該当者なし"
        Exit Sub
    End If
    
    ' 会社名・氏名が既に入力されている場合は確認
    If Trim(Me.Range("会社名").Value) <> "" Or Trim(Me.Range("氏名").Value) <> "" Then
        result = MsgBox("現在の入力内容を上書きしてよろしいですか？" & vbCrLf & _
                       vbCrLf & _
                       "検索結果:" & vbCrLf & _
                       "  会社名: " & foundCompany & vbCrLf & _
                       "  氏名: " & foundPerson, _
                       vbQuestion + vbYesNo, "上書き確認")
        
        If result <> vbYes Then Exit Sub
    End If
    
    ' イベントを一時停止
    Application.EnableEvents = False
    
    ' 会社名と氏名を設定
    Me.Range("会社名").Value = foundCompany
    Me.Range("氏名").Value = foundPerson
    
    ' イベントを再開
    Application.EnableEvents = True
    
    ' 連動プルダウンを手動で更新(カナ氏名等が自動入力される)
    Call UpdatePersonNameList(foundCompany, Me.Range("氏名"))
    
    ' 完了メッセージは不要(自然に入力されているので)
End Sub

' ------------------------------------------------------------------------------
' フリー3の値で受診者マスタを検索
' @param free3Value 検索するフリー3の値
' @param foundCompany 見つかった会社名(ByRef)
' @param foundPerson 見つかった氏名(ByRef)
' @return 見つかった場合True
' ------------------------------------------------------------------------------
Private Function SearchPersonByFree3(free3Value As String, _
                                     ByRef foundCompany As String, _
                                     ByRef foundPerson As String) As Boolean
    Dim masterTbl As ListObject
    Dim oRow As ListRow
    Dim rowFree3 As String
    
    SearchPersonByFree3 = False
    foundCompany = ""
    foundPerson = ""
    
    ' 受診者マスタを取得
    On Error Resume Next
    Set masterTbl = ws受診者マスタ.ListObjects(1)
    On Error GoTo 0
    
    If masterTbl Is Nothing Then Exit Function
    
    ' フリー3列が存在するかチェック
    On Error Resume Next
    Dim free3ColIndex As Long
    free3ColIndex = masterTbl.ListColumns("フリー3").Index
    On Error GoTo 0
    
    If free3ColIndex = 0 Then Exit Function
    
    ' テーブルを検索
    For Each oRow In masterTbl.ListRows
        rowFree3 = Trim(CStr(oRow.Range(free3ColIndex).Value))
        
        If rowFree3 = free3Value Then
            foundCompany = Trim(oRow.Range(masterTbl.ListColumns("会社名").Index).Value)
            foundPerson = Trim(oRow.Range(masterTbl.ListColumns("氏名").Index).Value)
            SearchPersonByFree3 = True
            Exit Function
        End If
    Next oRow
End Function


' ========================================================================================
' プルダウン連動処理関連
' ========================================================================================

' ------------------------------------------------------------------------------
' ワークシート変更イベント(連動プルダウン制御)
' 会社名変更時:氏名のプルダウンを絞り込み、必要に応じて氏名をクリア
' ------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim companyRange As Range
    Dim personRange As Range
    Dim selectedCompany As String
    Dim currentPerson As String
    
    ' 名前付き範囲の取得
    On Error Resume Next
    Set companyRange = Me.Range("会社名")
    Set personRange = Me.Range("氏名")
    On Error GoTo 0
    
    ' === 会社名変更時の処理 ===
    If Not companyRange Is Nothing And Not personRange Is Nothing Then
        If Not Intersect(Target, companyRange) Is Nothing Then
            Application.EnableEvents = False
            On Error GoTo CleanExit
            
            selectedCompany = Trim(companyRange.Value)
            currentPerson = Trim(personRange.Value)
            
            ' 会社名がクリアされた場合
            If selectedCompany = "" Then
                On Error Resume Next
                personRange.Validation.Delete
                personRange.ClearContents
                On Error GoTo 0
                Call ShowCompanyPanel
                GoTo CleanExit
            End If
            
            ' 氏名のプルダウンリストを更新(数式ベース)
            Call UpdatePersonNameList(selectedCompany, personRange)
            
            ' 既に氏名が選択されている場合、新しい会社に存在するかチェック
            If currentPerson <> "" Then
                If Not CheckPersonExistsInCompany(currentPerson, selectedCompany) Then
                    personRange.ClearContents
                    Call UpdatePersonNameList(selectedCompany, personRange)
                End If
            End If
            
            GoTo CleanExit
        End If
    End If
    
    ' === 氏名変更時の処理（選択パネルの強調表示を更新） ===
    If Not companyRange Is Nothing And Not personRange Is Nothing Then
        If Not Intersect(Target, personRange) Is Nothing Then
            Call UpdatePersonNameList(Trim(CStr(companyRange.Value)), personRange)
        End If
    End If
    
CleanExit:
    Application.EnableEvents = True
End Sub

' ------------------------------------------------------------------------------
' ワークシートアクティブ時イベント
' 氏名リストを最新状態に更新（受診者マスタの変更を反映）
' ------------------------------------------------------------------------------
Private Sub Worksheet_Activate()
    Dim companyRange As Range
    Dim personRange As Range
    Dim selectedCompany As String
    
    ' 名前付き範囲の取得
    On Error Resume Next
    Set companyRange = Me.Range("会社名")
    Set personRange = Me.Range("氏名")
    On Error GoTo 0
    
    If companyRange Is Nothing Or personRange Is Nothing Then Exit Sub
    
    ' 会社が1社だけなら自動で入力
    Call AutoFillSingleCompany
    
    selectedCompany = Trim(companyRange.Value)
    
    Application.EnableEvents = False
    On Error GoTo CleanExit
    
    ' 会社名の有無に関わらずリスト更新（空ならクリーンアップ）
    Call UpdatePersonNameList(selectedCompany, personRange)
    
    ' 次の受付番号を更新
    Call UpdateNextReceptionNumber
    
CleanExit:
    Application.EnableEvents = True
End Sub

' ------------------------------------------------------------------------------
' 氏名のプルダウンリストを更新(VBA抽出版・Excel2016/2019対応)
' 受診者マスタから会社別に抽出しカナ氏名順でAA列に書き出す
' @param companyName 選択された会社名(空文字の場合は入力規則削除)
' @param targetRange 氏名セル
' ------------------------------------------------------------------------------
Private Sub UpdatePersonNameList(companyName As String, targetRange As Range)
    ' 旧Excel(2016/2019)互換版: FILTER/SORTBY を使わず、VBAで名簿から該当者を抽出する
    ' ・AA列に書き出して氏名セルの入力規則(リスト)に設定
    ' ・同じ一覧を右側の選択パネルに大きな文字で表示（受付済みは灰色）
    Dim isProtected As Boolean
    Dim prevEvents As Boolean
    Dim masterTbl As ListObject
    Dim oRow As ListRow
    Dim nameArr() As String
    Dim kanaArr() As String
    Dim doneArr() As Boolean
    Dim n As Long, i As Long, j As Long
    Dim doneCount As Long
    Dim tmpN As String, tmpK As String, tmpD As Boolean
    Dim listRange As Range
    Dim compCol As Long, nameCol As Long, kanaCol As Long, doneCol As Long
    Dim selectedName As String
    Dim headerText As String
    Const LIST_COL As String = "AA"
    
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    
    isProtected = Me.ProtectContents
    If isProtected Then Me.Unprotect
    
    On Error GoTo CleanExit
    
    ' 既存のリストと入力規則をクリア
    Me.Columns(LIST_COL).ClearContents
    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo CleanExit
    
    ' 会社名が空なら会社の選択パネルを表示
    If companyName = "" Then
        Call ShowCompanyPanel
        GoTo CleanExit
    End If
    
    selectedName = Trim(CStr(targetRange.Value))
    
    ' 受診者マスタから該当会社の氏名・カナ氏名・受付済みを収集
    Set masterTbl = ws受診者マスタ.ListObjects(1)
    compCol = masterTbl.ListColumns("会社名").Index
    nameCol = masterTbl.ListColumns("氏名").Index
    kanaCol = masterTbl.ListColumns("カナ氏名").Index
    doneCol = 0
    On Error Resume Next
    doneCol = masterTbl.ListColumns("受付済み").Index
    On Error GoTo CleanExit
    
    ReDim nameArr(1 To masterTbl.ListRows.Count + 1)
    ReDim kanaArr(1 To masterTbl.ListRows.Count + 1)
    ReDim doneArr(1 To masterTbl.ListRows.Count + 1)
    n = 0
    doneCount = 0
    For Each oRow In masterTbl.ListRows
        If Trim(CStr(oRow.Range(compCol).Value)) = companyName And _
           Trim(CStr(oRow.Range(nameCol).Value)) <> "" Then
            n = n + 1
            nameArr(n) = Trim(CStr(oRow.Range(nameCol).Value))
            kanaArr(n) = Trim(CStr(oRow.Range(kanaCol).Value))
            doneArr(n) = False
            If doneCol > 0 Then
                If Trim(CStr(oRow.Range(doneCol).Value)) = "済" Then doneArr(n) = True
            End If
            If doneArr(n) Then doneCount = doneCount + 1
        End If
    Next oRow
    
    ' カナ氏名順に並べ替え(挿入ソート)
    For i = 2 To n
        tmpN = nameArr(i)
        tmpK = kanaArr(i)
        tmpD = doneArr(i)
        j = i - 1
        Do While j >= 1
            If StrComp(kanaArr(j), tmpK, vbTextCompare) <= 0 Then Exit Do
            nameArr(j + 1) = nameArr(j)
            kanaArr(j + 1) = kanaArr(j)
            doneArr(j + 1) = doneArr(j)
            j = j - 1
        Loop
        nameArr(j + 1) = tmpN
        kanaArr(j + 1) = tmpK
        doneArr(j + 1) = tmpD
    Next i
    
    If n > 0 Then
        ' AA列に書き出し
        Set listRange = Me.Range(LIST_COL & "1").Resize(n, 1)
        For i = 1 To n
            listRange.Cells(i, 1).Value = nameArr(i)
        Next i
        
        ' 入力規則(リスト)を設定
        On Error Resume Next
        With targetRange.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="=" & listRange.Address(True, True)
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        On Error GoTo CleanExit
        
        headerText = "氏名をクリックして選択　（全" & n & "名／受付済 " & doneCount & "名）"
    Else
        headerText = "この会社の名簿がありません（受診者マスタを確認）"
    End If
    
    Call RenderPanelItems(nameArr, doneArr, n, headerText, "person", selectedName)
    
CleanExit:
    If isProtected Then
        Me.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
    End If
    Application.EnableEvents = prevEvents
End Sub

' ------------------------------------------------------------------------------
' 指定された氏名が指定された会社に存在するかチェック(AA列リスト版)
' @param personName 氏名
' @param companyName 会社名
' @return 存在する場合True
' ------------------------------------------------------------------------------
Private Function CheckPersonExistsInCompany(personName As String, companyName As String) As Boolean
    ' AA列に書き出された氏名リストの中に存在するか
    Dim cell As Range
    Dim lastRow As Long
    
    CheckPersonExistsInCompany = False
    
    On Error Resume Next
    lastRow = Me.Cells(Me.Rows.Count, "AA").End(xlUp).Row
    On Error GoTo 0
    If lastRow < 1 Then Exit Function
    
    For Each cell In Me.Range("AA1:AA" & lastRow)
        If Trim(CStr(cell.Value)) = personName Then
            CheckPersonExistsInCompany = True
            Exit Function
        End If
    Next cell
End Function

' ========================================================================================
' Clear処理関連
' ========================================================================================

' ------------------------------------------------------------------------------
' 入力フィールドをクリア（メッセージなし）
' クリアボタンや受付登録後の自動クリアから呼ばれる共通処理
' 結合セル対応版
' ------------------------------------------------------------------------------
Public Sub ClearInputFields()
    Dim targetNames As Variant
    Dim i As Long
    Dim targetRange As Range
    Dim isProtected As Boolean
    
    ' シートの保護状態を記録
    isProtected = Me.ProtectContents
    
    ' 保護されている場合は一時的に解除
    If isProtected Then
        Me.Unprotect
    End If
    
    ' イベントの無限ループを防ぐため一時停止
    Application.EnableEvents = False
    
    On Error GoTo CleanExit
    
    ' クリア対象の名前付き範囲リスト
    targetNames = Array("会社名", "氏名", "便潜血", _
                       "オプション1", "オプション2", "オプション3", "オプション4", _
                       "オプション5", "オプション6", "オプション7", "オプション8", _
                       "備考")
    
    ' 各セルをクリア（結合セル対応）
    For i = LBound(targetNames) To UBound(targetNames)
        On Error Resume Next
        Set targetRange = Me.Range(targetNames(i))
        
        If Not targetRange Is Nothing Then
            Call ClearRangeSafely(targetRange)
        End If
        
        Set targetRange = Nothing
        On Error GoTo CleanExit
    Next i
    
CleanExit:
    Application.EnableEvents = True
    
    ' 元々保護されていた場合は再度保護
    If isProtected Then
        Me.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
    End If
    
    ' 会社が1社だけなら会社名を入れ直し、選択パネルを最新にする
    Call AutoFillSingleCompany
    Call UpdatePersonNameList(Trim(CStr(Me.Range("会社名").Value)), Me.Range("氏名"))
End Sub

' ========================================================================================
' 選択パネル（大きな文字の一覧からクリックで選択）
' ========================================================================================

' ------------------------------------------------------------------------------
' パネル内のセルがクリックされたら、その会社名・氏名を入力する
' ------------------------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim v As String
    Dim mode As String
    Dim panelArea As Range
    
    On Error GoTo SafeExit
    If Target.CountLarge <> 1 Then Exit Sub
    
    mode = CStr(Me.Range(PANEL_MODE_CELL).Value)
    
    ' 見出し右端の「会社を選び直す」
    If Target.Row = PANEL_HEADER_ROW And Target.Column = PANEL_FIRST_COL + PANEL_COLS - 1 Then
        If mode = "person" Then Call ShowCompanyPanel
        Exit Sub
    End If
    
    Set panelArea = Me.Range(Me.Cells(PANEL_FIRST_ROW, PANEL_FIRST_COL), _
                             Me.Cells(PANEL_LAST_ROW, PANEL_FIRST_COL + PANEL_COLS - 1))
    If Intersect(Target, panelArea) Is Nothing Then Exit Sub
    
    v = Trim(CStr(Target.Value))
    If v = "" Then Exit Sub
    
    If mode = "company" Then
        ' 会社名を入れると Worksheet_Change で氏名の一覧に切り替わる
        Application.EnableEvents = True
        Me.Range("会社名").Value = v
        Me.Range("氏名").Select
    ElseIf mode = "person" Then
        Application.EnableEvents = False
        Me.Range("氏名").Value = v
        Application.EnableEvents = True
        Call UpdatePersonNameList(Trim(CStr(Me.Range("会社名").Value)), Me.Range("氏名"))
        Me.Range("氏名").Select
    End If
    Exit Sub
    
SafeExit:
    Application.EnableEvents = True
End Sub

' ------------------------------------------------------------------------------
' 会社マスタに会社が1社だけ登録されていて、会社名が空なら自動で入力する
' ------------------------------------------------------------------------------
Private Sub AutoFillSingleCompany()
    Dim tbl As ListObject
    Dim oRow As ListRow
    Dim v As String
    Dim found As String
    Dim cnt As Long
    Dim prevEvents As Boolean
    
    On Error GoTo SafeExit
    If Trim(CStr(Me.Range("会社名").Value)) <> "" Then Exit Sub
    
    Set tbl = ws会社マスタ.ListObjects(1)
    cnt = 0
    For Each oRow In tbl.ListRows
        v = Trim(CStr(oRow.Range(1).Value))
        If v <> "" Then
            cnt = cnt + 1
            found = v
        End If
    Next oRow
    
    If cnt = 1 Then
        prevEvents = Application.EnableEvents
        Application.EnableEvents = False
        Me.Range("会社名").Value = found
        Application.EnableEvents = prevEvents
    End If
SafeExit:
End Sub

' ------------------------------------------------------------------------------
' 会社マスタの会社名を選択パネルに表示する
' ------------------------------------------------------------------------------
Private Sub ShowCompanyPanel()
    Dim tbl As ListObject
    Dim oRow As ListRow
    Dim items() As String
    Dim flags() As Boolean
    Dim n As Long
    Dim v As String
    Dim headerText As String
    
    On Error GoTo SafeExit
    Set tbl = ws会社マスタ.ListObjects(1)
    ReDim items(1 To tbl.ListRows.Count + 1)
    ReDim flags(1 To tbl.ListRows.Count + 1)
    n = 0
    For Each oRow In tbl.ListRows
        v = Trim(CStr(oRow.Range(1).Value))
        If v <> "" Then
            n = n + 1
            items(n) = v
            flags(n) = False
        End If
    Next oRow
    
    If n > 0 Then
        headerText = "会社名をクリックして選択"
    Else
        headerText = "会社マスタに会社名が登録されていません"
    End If
    Call RenderPanelItems(items, flags, n, headerText, "company", Trim(CStr(Me.Range("会社名").Value)))
SafeExit:
End Sub

' ------------------------------------------------------------------------------
' 選択パネルを描画する
' @param items 表示する文字列(1～n)
' @param doneFlags Trueの項目は灰色（受付済み）
' @param mode "company" または "person"
' @param selectedValue 強調表示する項目
' ------------------------------------------------------------------------------
Private Sub RenderPanelItems(items() As String, doneFlags() As Boolean, n As Long, _
                             headerText As String, mode As String, selectedValue As String)
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim rowsList As Collection
    Dim r As Long
    Dim perCol As Long
    Dim i As Long
    Dim colIdx As Long
    Dim rowIdx As Long
    Dim cell As Range
    Dim lastCol As Long
    
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo SafeExit
    
    lastCol = PANEL_FIRST_COL + PANEL_COLS - 1
    
    ' パネル全体を初期化
    With Me.Range(Me.Cells(PANEL_HEADER_ROW, PANEL_FIRST_COL), Me.Cells(PANEL_LAST_ROW, lastCol))
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        .Font.Name = "メイリオ"
        .Font.Size = PANEL_FONT_SIZE
        .Font.Bold = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' 見出し
    With Me.Range(Me.Cells(PANEL_HEADER_ROW, PANEL_FIRST_COL), Me.Cells(PANEL_HEADER_ROW, lastCol))
        .Interior.Color = RGB(44, 95, 158)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
    End With
    With Me.Cells(PANEL_HEADER_ROW, PANEL_FIRST_COL)
        .Value = headerText
        .HorizontalAlignment = xlLeft
    End With
    If mode = "person" Then
        With Me.Cells(PANEL_HEADER_ROW, lastCol)
            .Value = "［会社を選び直す］"
            .HorizontalAlignment = xlRight
            .Font.Size = 11
            .Font.Underline = xlUnderlineStyleSingle
            .Font.Color = RGB(255, 235, 156)
        End With
    End If
    Me.Range(PANEL_MODE_CELL).Value = mode
    
    If n <= 0 Then GoTo SafeExit
    
    ' 表示に使える行（非表示行・細い行は除く）
    Set rowsList = New Collection
    For r = PANEL_FIRST_ROW To PANEL_LAST_ROW
        If Not Me.Rows(r).Hidden Then
            If Me.Rows(r).RowHeight >= 20 Then rowsList.Add r
        End If
    Next r
    If rowsList.Count = 0 Then GoTo SafeExit
    
    ' 縦に並べてから次の列へ（カナ順が上から下に読める）
    perCol = (n + PANEL_COLS - 1) \ PANEL_COLS
    If perCol > rowsList.Count Then perCol = rowsList.Count
    
    For i = 1 To n
        colIdx = (i - 1) \ perCol
        rowIdx = ((i - 1) Mod perCol) + 1
        If colIdx >= PANEL_COLS Then Exit For
        Set cell = Me.Cells(rowsList(rowIdx), PANEL_FIRST_COL + colIdx)
        cell.Value = items(i)
        cell.Borders.LineStyle = xlContinuous
        cell.Borders.Weight = xlThin
        cell.Borders.Color = RGB(180, 200, 220)
        If items(i) = selectedValue Then
            cell.Interior.Color = RGB(255, 235, 156)
            cell.Font.Bold = True
        ElseIf doneFlags(i) Then
            cell.Interior.Color = RGB(225, 225, 225)
            cell.Font.Color = RGB(150, 150, 150)
        Else
            cell.Interior.Color = RGB(235, 243, 252)
        End If
    Next i
    
SafeExit:
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
End Sub
