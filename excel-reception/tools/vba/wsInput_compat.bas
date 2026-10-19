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
                GoTo CleanExit
            End If
            
            ' 氏名のプルダウンリストを更新(数式ベース)
            Call UpdatePersonNameList(selectedCompany, personRange)
            
            ' 既に氏名が選択されている場合、新しい会社に存在するかチェック
            If currentPerson <> "" Then
                If Not CheckPersonExistsInCompany(currentPerson, selectedCompany) Then
                    personRange.ClearContents
                End If
            End If
            
            GoTo CleanExit
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
    ' 旧Excel(2016/2019)互換版: FILTER/SORTBY を使わず、VBAで名簿から該当者を抽出し
    ' AA列に書き出して、その範囲を氏名セルの入力規則(リスト)に設定する
    Dim isProtected As Boolean
    Dim prevEvents As Boolean
    Dim masterTbl As ListObject
    Dim oRow As ListRow
    Dim nameArr() As String
    Dim kanaArr() As String
    Dim n As Long, i As Long, j As Long
    Dim tmpN As String, tmpK As String
    Dim listRange As Range
    Dim compCol As Long, nameCol As Long, kanaCol As Long
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
    
    If companyName = "" Then GoTo CleanExit
    
    ' 受診者マスタから該当会社の氏名・カナ氏名を収集
    Set masterTbl = ws受診者マスタ.ListObjects(1)
    compCol = masterTbl.ListColumns("会社名").Index
    nameCol = masterTbl.ListColumns("氏名").Index
    kanaCol = masterTbl.ListColumns("カナ氏名").Index
    
    n = 0
    If masterTbl.ListRows.Count > 0 Then
        ReDim nameArr(1 To masterTbl.ListRows.Count)
        ReDim kanaArr(1 To masterTbl.ListRows.Count)
        For Each oRow In masterTbl.ListRows
            If Trim(CStr(oRow.Range(compCol).Value)) = companyName And _
               Trim(CStr(oRow.Range(nameCol).Value)) <> "" Then
                n = n + 1
                nameArr(n) = Trim(CStr(oRow.Range(nameCol).Value))
                kanaArr(n) = Trim(CStr(oRow.Range(kanaCol).Value))
            End If
        Next oRow
    End If
    
    If n = 0 Then GoTo CleanExit
    
    ' カナ氏名順に並べ替え(挿入ソート)
    For i = 2 To n
        tmpN = nameArr(i)
        tmpK = kanaArr(i)
        j = i - 1
        Do While j >= 1
            If StrComp(kanaArr(j), tmpK, vbTextCompare) <= 0 Then Exit Do
            nameArr(j + 1) = nameArr(j)
            kanaArr(j + 1) = kanaArr(j)
            j = j - 1
        Loop
        nameArr(j + 1) = tmpN
        kanaArr(j + 1) = tmpK
    Next i
    
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
    On Error GoTo 0
    
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
End Sub
