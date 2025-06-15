' ===============================================
' PDFデータプロセッサー：PDF読み込み・ワークシート作成
' 作成者：関西のおばちゃん
' 作成日：2025/06/15
' 概要：パワークエリでPDFを読み込んでワークシートを作成するで〜♪
' 
' 処理フロー：
' 1. パワークエリで任意のPDFファイルを開く
' 2. Table001 (Page 1) とTable002 (Page 1)が出来上がる
' 3. データの変換で各テーブルを別々のワークシートに作成
' ===============================================

Option Explicit

Sub PDF更新()
    ' -----------------------------------------------
    ' PDFからパワークエリでデータを更新する
    ' 既存のパワークエリ接続を使ってデータを最新に更新するで〜
    ' -----------------------------------------------
    
    Dim wb As Workbook
    Dim conn As WorkbookConnection
    Dim queryName As String
    Dim errorMessage As String
    
    ' エラーハンドリング設定
    On Error GoTo UpdateError
    
    Set wb = ThisWorkbook
    
    MsgBox "PDFからデータを更新するで〜" & vbCrLf & _
           "少し時間がかかるかもしれんから、お茶でも飲んで待っててや♪", _
           vbInformation, "PDF更新中"
    
    ' パワークエリの接続を全て更新（PDF関連のクエリを一括更新）
    For Each conn In wb.Connections
        ' PDF関連のクエリを探して更新
        If InStr(conn.Name, "Table001") > 0 Or InStr(conn.Name, "Table002") > 0 Then
            Debug.Print "クエリ更新中：" & conn.Name
            conn.Refresh
        End If
    Next conn
    
    ' 更新完了まで少し待つ（大きなPDFの場合時間がかかる）
    Application.Wait (Now + TimeValue("0:00:05"))
    
    MsgBox "PDFデータの更新が完了したで〜♪", vbInformation, "更新完了"
    
    Exit Sub

UpdateError:
    errorMessage = "PDFの更新でエラーが起きたで〜" & vbCrLf & _
                   "エラー内容：" & Err.Description & vbCrLf & _
                   "PDFファイルが見つからんか、パワークエリの設定を確認してや"
    
    MsgBox errorMessage, vbCritical, "PDF更新エラー"
End Sub

Sub テーブル振り分け()
    ' -----------------------------------------------
    ' Table001とTable002を別々のワークシートに振り分ける
    ' パワークエリで取得したテーブルを整理するで〜
    ' -----------------------------------------------
    
    Dim wsTable1 As Worksheet
    Dim wsTable2 As Worksheet
    Dim wsSource As Worksheet    ' パワークエリの結果が入っているシート
    Dim lo1 As ListObject        ' Table001のリストオブジェクト
    Dim lo2 As ListObject        ' Table002のリストオブジェクト
    
    ' エラーハンドリング設定
    On Error GoTo SeparateError
    
    ' ソースシートを取得（パワークエリの結果が入っているシート）
    ' ※事前にパワークエリでPDFを読み込んでおく必要があるで
    Set wsSource = ThisWorkbook.Worksheets("PDFデータ")  ' シート名は環境に合わせて変更してや
    
    ' ワークシートの準備（なかったら作成）
    Set wsTable1 = 取得または作成ワークシート("Table001_Page1")
    Set wsTable2 = 取得または作成ワークシート("Table002_Page1")
    
    ' 既存データをクリア
    wsTable1.Cells.Clear
    wsTable2.Cells.Clear
    
    ' Table001 (Page 1)のデータをコピー
    If パワークエリテーブル存在確認(wsSource, "Table001") Then
        Set lo1 = wsSource.ListObjects("Table001")
        
        ' ヘッダー付きでコピー
        lo1.Range.Copy
        wsTable1.Range("A1").PasteSpecial Paste:=xlPasteValues
        
        ' 見た目も整える
        wsTable1.Range("A1").CurrentRegion.AutoFit
        wsTable1.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
        
        Debug.Print "Table001のデータを「Table001_Page1」シートにコピー完了"
    Else
        MsgBox "Table001が見つからへんわ〜" & vbCrLf & _
               "パワークエリでPDFを読み込んだか確認してや", vbExclamation, "テーブルなし"
    End If
    
    ' Table002 (Page 1)のデータをコピー
    If パワークエリテーブル存在確認(wsSource, "Table002") Then
        Set lo2 = wsSource.ListObjects("Table002")
        
        ' ヘッダー付きでコピー
        lo2.Range.Copy
        wsTable2.Range("A1").PasteSpecial Paste:=xlPasteValues
        
        ' 見た目も整える
        wsTable2.Range("A1").CurrentRegion.AutoFit
        wsTable2.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
        
        Debug.Print "Table002のデータを「Table002_Page1」シートにコピー完了"
    Else
        MsgBox "Table002が見つからへんわ〜" & vbCrLf & _
               "パワークエリでPDFを読み込んだか確認してや", vbExclamation, "テーブルなし"
    End If
    
    ' コピーモードを解除
    Application.CutCopyMode = False
    
    MsgBox "テーブルの振り分けが完了したで〜♪" & vbCrLf & _
           "「Table001_Page1」と「Table002_Page1」シートを確認してや", _
           vbInformation, "振り分け完了"
    
    Exit Sub

SeparateError:
    MsgBox "テーブル振り分けでエラーが起きたで〜" & vbCrLf & Err.Description, _
           vbCritical, "振り分けエラー"
End Sub

Function 取得または作成ワークシート(sheetName As String) As Worksheet
    ' -----------------------------------------------
    ' ワークシートを取得、なかったら新規作成する関数
    ' -----------------------------------------------
    
    Dim ws As Worksheet
    
    ' シートが存在するかチェック
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    ' シートがなかったら新規作成
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
        Debug.Print "新しいワークシート「" & sheetName & "」を作成したで〜"
    Else
        Debug.Print "既存のワークシート「" & sheetName & "」を使用するで〜"
    End If
    
    Set 取得または作成ワークシート = ws
End Function

Function パワークエリテーブル存在確認(ws As Worksheet, tableName As String) As Boolean
    ' -----------------------------------------------
    ' 指定されたワークシートに指定されたテーブルがあるかチェック
    ' -----------------------------------------------
    
    Dim lo As ListObject
    
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If Not lo Is Nothing Then
        パワークエリテーブル存在確認 = True
        Debug.Print "テーブル「" & tableName & "」が見つかったで〜"
    Else
        パワークエリテーブル存在確認 = False
        Debug.Print "テーブル「" & tableName & "」が見つからへんわ〜"
    End If
End Function

Sub パワークエリ接続確認()
    ' -----------------------------------------------
    ' 現在のパワークエリ接続を確認する（デバッグ用）
    ' -----------------------------------------------
    
    Dim conn As WorkbookConnection
    Dim queryCount As Integer
    
    queryCount = 0
    
    Debug.Print "=== パワークエリ接続一覧 ==="
    
    For Each conn In ThisWorkbook.Connections
        queryCount = queryCount + 1
        Debug.Print queryCount & ". " & conn.Name & " (" & conn.Type & ")"
    Next conn
    
    If queryCount = 0 Then
        Debug.Print "パワークエリ接続が見つからへんわ〜"
        MsgBox "パワークエリ接続がないで〜" & vbCrLf & _
               "まずはPDFファイルをパワークエリで読み込んでや", _
               vbExclamation, "接続なし"
    Else
        Debug.Print "合計 " & queryCount & " 個の接続が見つかったで〜"
    End If
End Sub