' ===============================================
' プロシージャ名：BasicProcessor（基本処理4ステップ）
' 作成者：関西のおばちゃん
' 作成日：2025/06/16
' 概要：PDF取得後の基本データ整理処理（4ステップのみ）
'       1. 文字列削除：(内作)(別注)(全ネジ) → ""
'       2. 空白行削除：A~E列が空欄の行を削除
'       3. データ統合：Table002 → Table001に統合
'       4. シート名変更：Table001 → "原価リスト"
' ※変数名は英語、コメントは関西弁で初心者にもわかりやすく♪
' ===============================================

Option Explicit

' ===============================================
' メイン処理：基本処理（4ステップ）
' ===============================================

Sub 基本処理_4ステップ()
    ' -----------------------------------------------
    ' PDF取得後の基本データ整理処理
    ' シンプルに4ステップだけ実行するで〜♪
    ' -----------------------------------------------
    
    Dim response As VbMsgBoxResult
    
    ' 実行確認
    response = MsgBox("基本処理（4ステップ）を実行するで〜" & vbCrLf & vbCrLf & _
                      "1. (内作)(別注)(全ネジ) 文字列削除" & vbCrLf & _
                      "2. A~E列空欄行を削除" & vbCrLf & _
                      "3. Table002 → Table001に統合" & vbCrLf & _
                      "4. Table001 → 「原価リスト」に名前変更" & vbCrLf & vbCrLf & _
                      "実行してもええ？", _
                      vbYesNo + vbQuestion, "基本処理4ステップ")
    
    If response = vbNo Then
        MsgBox "処理をキャンセルしたで〜", vbInformation, "キャンセル"
        Exit Sub
    End If
    
    ' 画面更新を止めて処理を早くする
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    MsgBox "基本処理を開始するで〜♪", vbInformation, "処理開始"
    
    ' ステップ実行（順番大事やで〜）
    Call Step1_文字列削除処理()
    Call Step2_空白行削除処理()
    Call Step3_データ統合処理()
    Call Step4_シート名変更処理()
    
    ' 画面更新を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "基本処理（4ステップ）完了〜♪" & vbCrLf & _
           "「原価リスト」シートを確認してや〜", vbInformation, "処理完了"
    
End Sub

' ===============================================
' ステップ1：文字列削除処理
' ===============================================

Sub Step1_文字列削除処理()
    ' -----------------------------------------------
    ' (内作)(別注)(全ネジ)の文字列を削除
    ' 行は削除せず、文字列だけ ""に置換するで〜
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim replaceCount As Integer
    
    Debug.Print "=== ステップ1: 文字列削除開始 ==="
    
    ' Table001の処理
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        replaceCount = 0
        With ws1
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            ' 各セルをチェックして文字列削除
            For i = 1 To lastRow
                For col = 1 To 5  ' A~E列
                    cellValue = CStr(.Cells(i, col).Value)
                    
                    ' (内作)を削除
                    If InStr(cellValue, "(内作)") > 0 Then
                        cellValue = Replace(cellValue, "(内作)", "")
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                    End If
                    
                    ' (別注)を削除
                    If InStr(cellValue, "(別注)") > 0 Then
                        cellValue = Replace(cellValue, "(別注)", "")
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                    End If
                    
                    ' (全ネジ)を削除
                    If InStr(cellValue, "(全ネジ)") > 0 Then
                        cellValue = Replace(cellValue, "(全ネジ)", "")
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table001: " & replaceCount & "個の文字列を削除"
    Else
        Debug.Print "Table001シートが見つからない"
    End If
    
    ' Table002の処理
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        replaceCount = 0
        With ws2
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            ' 各セルをチェックして文字列削除
            For i = 1 To lastRow
                For col = 1 To 5  ' A~E列
                    cellValue = CStr(.Cells(i, col).Value)
                    
                    ' (内作)を削除
                    If InStr(cellValue, "(内作)") > 0 Then
                        cellValue = Replace(cellValue, "(内作)", "")
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                    End If
                    
                    ' (別注)を削除
                    If InStr(cellValue, "(別注)") > 0 Then
                        cellValue = Replace(cellValue, "(別注)", "")
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                    End If
                    
                    ' (全ネジ)を削除
                    If InStr(cellValue, "(全ネジ)") > 0 Then
                        cellValue = Replace(cellValue, "(全ネジ)", "")
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table002: " & replaceCount & "個の文字列を削除"
    Else
        Debug.Print "Table002シートが見つからない"
    End If
    
    Debug.Print "=== ステップ1完了 ==="
    
End Sub

' ===============================================
' ステップ2：空白行削除処理
' ===============================================

Sub Step2_空白行削除処理()
    ' -----------------------------------------------
    ' A~E列がすべて空欄の行を削除するで〜
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    
    Debug.Print "=== ステップ2: 空白行削除開始 ==="
    
    ' Table001の空白行削除
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        Call Delete空白行(ws1)
    End If
    
    ' Table002の空白行削除
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        Call Delete空白行(ws2)
    End If
    
    Debug.Print "=== ステップ2完了 ==="
    
End Sub

Sub Delete空白行(ws As Worksheet)
    ' -----------------------------------------------
    ' 指定されたワークシートの空白行を削除
    ' A~E列がすべて空欄の行が対象やで〜
    ' -----------------------------------------------
    
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim isEmpty As Boolean
    Dim deleteCount As Integer
    
    With ws
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        deleteCount = 0
        
        ' 下から上に向かって処理（行削除で番号がズレるのを防ぐため）
        For i = lastRow To 2 Step -1  ' 1行目はヘッダーなので除外
            isEmpty = True
            
            ' A~E列をすべてチェック
            For col = 1 To 5
                If Trim(CStr(.Cells(i, col).Value)) <> "" Then
                    isEmpty = False  ' 何かデータがあったら空白行ではない
                    Exit For
                End If
            Next col
            
            ' すべての列が空白なら行を削除
            If isEmpty Then
                .Rows(i).Delete Shift:=xlUp
                deleteCount = deleteCount + 1
            End If
        Next i
        
        Debug.Print ws.Name & ": " & deleteCount & "行の空白行を削除"
    End With
    
End Sub

' ===============================================
' ステップ3：データ統合処理
' ===============================================

Sub Step3_データ統合処理()
    ' -----------------------------------------------
    ' Table002のデータをTable001の最終行に追加
    ' テーブルは使わず、普通のコピペやで〜
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, col As Integer
    Dim copyCount As Integer
    Dim hasData As Boolean
    
    Debug.Print "=== ステップ3: データ統合開始 ==="
    
    ' シートを取得
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    ' 両方のシートが存在するかチェック
    If ws1 Is Nothing Then
        Debug.Print "Table001シートが見つからない"
        Exit Sub
    End If
    
    If ws2 Is Nothing Then
        Debug.Print "Table002シートが見つからない"
        Exit Sub
    End If
    
    ' 最終行を取得
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    
    copyCount = 0
    
    ' Table002のデータをTable001にコピー（ヘッダー行は除く）
    For i = 2 To lastRow2  ' 2行目から開始（1行目はヘッダー）
        hasData = False
        
        ' その行にデータがあるかチェック
        For col = 1 To 5
            If Trim(CStr(ws2.Cells(i, col).Value)) <> "" Then
                hasData = True
                Exit For
            End If
        Next col
        
        ' データがある行のみコピー
        If hasData Then
            lastRow1 = lastRow1 + 1  ' Table001の次の行
            
            ' A~E列をコピ