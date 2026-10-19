import re
src = open('vba/wsInput_new.bas','rb').read().decode('cp932').replace('\r\n','\n')

def cut_proc(s, start_sig, end_kw):
    a = s.index(start_sig); b = s.index(end_kw, a) + len(end_kw)
    return a, b

# 1) remove zoom block (header comment + var + SelectionChange)
a = s0 = src.index("' ========================================================================================\n' プルダウン拡大表示")
b = src.index("End Sub", src.index("Private Sub Worksheet_SelectionChange")) + len("End Sub\n")
consts = """' ========================================================================================
' 選択パネル設定（会社名・氏名を大きな文字で一覧表示し、クリックで選択）
' ========================================================================================
Private Const PANEL_FIRST_COL As Long = 9      ' I列から
Private Const PANEL_COLS As Long = 5           ' 5列（I～M列）
Private Const PANEL_HEADER_ROW As Long = 5
Private Const PANEL_FIRST_ROW As Long = 6
Private Const PANEL_LAST_ROW As Long = 200
Private Const PANEL_MODE_CELL As String = "AB1"
Private Const PANEL_FONT_SIZE As Long = 16
"""
src = src[:a] + consts + src[b:]

# 2) Worksheet_Change: company cleared -> company panel; re-render after clearing person; person change -> highlight
wc_a, wc_b = cut_proc(src, "Private Sub Worksheet_Change(", "End Sub")
wc = src[wc_a:wc_b]
old = """                personRange.Validation.Delete
                personRange.ClearContents
                On Error GoTo 0
                GoTo CleanExit"""
assert wc.count(old) == 1
wc = wc.replace(old, """                personRange.Validation.Delete
                personRange.ClearContents
                On Error GoTo 0
                Call ShowCompanyPanel
                GoTo CleanExit""")
old = """                If Not CheckPersonExistsInCompany(currentPerson, selectedCompany) Then
                    personRange.ClearContents
                End If"""
assert wc.count(old) == 1
wc = wc.replace(old, """                If Not CheckPersonExistsInCompany(currentPerson, selectedCompany) Then
                    personRange.ClearContents
                    Call UpdatePersonNameList(selectedCompany, personRange)
                End If""")
old = "    End If\n    \nCleanExit:\n    Application.EnableEvents = True\nEnd Sub"
assert wc.count(old) == 1
wc = wc.replace(old, """    End If
    
    ' === 氏名変更時の処理（選択パネルの強調表示を更新） ===
    If Not companyRange Is Nothing And Not personRange Is Nothing Then
        If Not Intersect(Target, personRange) Is Nothing Then
            Call UpdatePersonNameList(Trim(CStr(companyRange.Value)), personRange)
        End If
    End If
    
CleanExit:
    Application.EnableEvents = True
End Sub""")
src = src[:wc_a] + wc + src[wc_b:]

# 3) Worksheet_Activate: auto-fill single company
wa_a, wa_b = cut_proc(src, "Private Sub Worksheet_Activate(", "End Sub")
wa = src[wa_a:wa_b]
old = "    selectedCompany = Trim(companyRange.Value)\n"
assert wa.count(old) == 1
wa = wa.replace(old, "    ' 会社が1社だけなら自動で入力\n    Call AutoFillSingleCompany\n    \n" + old)
src = src[:wa_a] + wa + src[wa_b:]

# 4) UpdatePersonNameList: full rewrite
up_a, up_b = cut_proc(src, "Private Sub UpdatePersonNameList(", "End Sub")
src = src[:up_a] + """Private Sub UpdatePersonNameList(companyName As String, targetRange As Range)
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
End Sub""" + src[up_b:]

# 5) ClearInputFields: refill single company + refresh panel
cl_a, cl_b = cut_proc(src, "Public Sub ClearInputFields()", "End Sub")
cl = src[cl_a:cl_b]
old = """    If isProtected Then
        Me.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
    End If
End Sub"""
assert cl.endswith(old)
cl = cl[:-len(old)] + """    If isProtected Then
        Me.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
    End If
    
    ' 会社が1社だけなら会社名を入れ直し、選択パネルを最新にする
    Call AutoFillSingleCompany
    Call UpdatePersonNameList(Trim(CStr(Me.Range("会社名").Value)), Me.Range("氏名"))
End Sub"""
src = src[:cl_a] + cl + src[cl_b:]

# 6) panel procedures appended at end
src = src.rstrip('\n') + """

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
    perCol = (n + PANEL_COLS - 1) \\ PANEL_COLS
    If perCol > rowsList.Count Then perCol = rowsList.Count
    
    For i = 1 To n
        colIdx = (i - 1) \\ perCol
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
"""
src = src.replace('\n', '\r\n')
assert 'mSavedZoom' not in src and 'DROPDOWN_ZOOM' not in src
open('vba/wsInput_panel.bas','wb').write(src.encode('cp932'))
s = src
print("bytes", len(src.encode('cp932')), "lines", s.count('\r\n'))
print("Sub", len(re.findall(r'^(?:Public |Private )?Sub ', s, re.M)), "End Sub", len(re.findall(r'^End Sub', s, re.M)),
      "| Function", len(re.findall(r'^(?:Public |Private )?Function ', s, re.M)), "End Function", len(re.findall(r'^End Function', s, re.M)))
# declarations must precede first procedure
first_proc = re.search(r'^(?:Public |Private )?(?:Sub|Function) ', s, re.M).start()
print("all Const before first proc:", all(m.start() < first_proc for m in re.finditer(r'^Private Const', s, re.M)))
