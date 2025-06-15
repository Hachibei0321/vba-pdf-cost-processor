' ===============================================
' 完全自動PDF読み込み処理（マクロ記録ベース改良版）
' 作成者：関西のおばちゃん
' 作成日：2025/06/15
' 概要：マクロ記録をベースに任意のPDFファイルを完全自動で読み込むで〜♪
' ===============================================

Option Explicit

Sub 完全自動PDF読み込み()
    ' -----------------------------------------------
    ' マクロ記録をベースにした完全自動PDF読み込み
    ' 手動設定なしで一気に処理するで〜♪
    ' -----------------------------------------------
    
    Dim pdfFilePath As String
    Dim errorMessage As String
    
    ' エラーハンドリング設定
    On Error GoTo AutoProcessError
    
    ' 画面更新を止めて処理を高速化
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' ステップ1：PDFファイルを選択
    MsgBox "完全自動でPDFを読み込むで〜♪" & vbCrLf & _
           "PDFファイルを選んでや〜", vbInformation, "完全自動PDF読み込み"
    
    pdfFilePath = PDFファイル選択ダイアログ()
    
    If pdfFilePath = "" Then
        MsgBox "PDFファイルが選ばれへんかったから、処理をやめるで〜", vbInformation, "処理中止"
        GoTo CleanUp
    End If
    
    ' ステップ2：既存の接続とシートをクリーンアップ
    Call 既存データクリーンアップ()
    
    ' ステップ3：Table001を完全自動で作成
    MsgBox "Table001を自動作成中やで〜", vbInformation, "処理中"
    Call Table001自動作成(pdfFilePath)
    
    ' ステップ4：Table002を完全自動で作成
    MsgBox "Table002を自動作成中やで〜", vbInformation, "処理中"
    Call Table002自動作成(pdfFilePath)
    
    ' 処理完了
    MsgBox "完全自動PDF読み込みが完了したで〜♪" & vbCrLf & _
           "「Table001_Page1」と「Table002_Page1」シートを確認してや", _
           vbInformation, "処理完了"

CleanUp:
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Exit Sub

AutoProcessError:
    errorMessage = "完全自動処理でエラーが起きたで〜" & vbCrLf & _
                   "エラー内容：" & Err.Description
    
    MsgBox errorMessage, vbCritical, "自動処理エラー"
    GoTo CleanUp
End Sub

Sub 既存データクリーンアップ()
    ' -----------------------------------------------
    ' 既存のクエリ・接続・ワークシートをクリーンアップ
    ' 新しいPDF用に準備するで〜
    ' -----------------------------------------------
    
    Dim ws As Worksheet
    Dim conn As WorkbookConnection
    Dim query As WorkbookQuery
    Dim i As Integer
    
    ' 既存のクエリを削除
    For i = ActiveWorkbook.Queries.Count To 1 Step -1
        Set query = ActiveWorkbook.Queries(i)
        If InStr(LCase(query.Name), "table001") > 0 Or _
           InStr(LCase(query.Name), "table002") > 0 Then
            Debug.Print "削除するクエリ：" & query.Name
            query.Delete
        End If
    Next i
    
    ' 既存の接続を削除
    For i = ThisWorkbook.Connections.Count To 1 Step -1
        Set conn = ThisWorkbook.Connections(i)
        If InStr(LCase(conn.Name), "table") > 0 Then
            Debug.Print "削除する接続：" & conn.Name
            conn.Delete
        End If
    Next i
    
    ' 既存のワークシートをクリア（削除はしない）
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Table001_Page1")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        Debug.Print "Table001_Page1シートをクリア"
    End If
    
    Set ws = ThisWorkbook.Worksheets("Table002_Page1")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        Debug.Print "Table002_Page1シートをクリア"
    End If
    On Error GoTo 0
    
    Debug.Print "既存データクリーンアップ完了"
End Sub

Sub Table001自動作成(pdfPath As String)
    ' -----------------------------------------------
    ' Table001を完全自動で作成（マクロ記録ベース）
    ' -----------------------------------------------
    
    Dim ws As Worksheet
    Dim queryName As String
    Dim formulaText As String
    
    ' 一意のクエリ名を生成
    queryName = "Table001_" & Format(Now, "yyyymmdd_hhnnss")
    
    ' パワークエリの数式を作成（パスを動的に設定）
    formulaText = "let" & Chr(13) & Chr(10) & _
                  "    ソース = Pdf.Tables(File.Contents(""" & pdfPath & """), [Implementation=""1.3""])," & Chr(13) & Chr(10) & _
                  "    Table001 = ソース{[Id=""Table001""]}[Data]," & Chr(13) & Chr(10) & _
                  "    昇格されたヘッダー数 = Table.PromoteHeaders(Table001, [PromoteAllScalars=true])," & Chr(13) & Chr(10) & _
                  "    変更された型 = Table.TransformColumnTypes(昇格されたヘッダー数,{{""Column1"", type text}, {"""", type text}, {""品名記号"", type text}, {""質量(kg)"", type number}, {""仕入単価(ｔ)"", Int64.Type}, {""?般市中?さ"", type text}})" & Chr(13) & Chr(10) & _
                  "in" & Chr(13) & Chr(10) & _
                  "    変更された型"
    
    ' クエリを作成
    ActiveWorkbook.Queries.Add Name:=queryName, Formula:=formulaText
    
    ' ワークシートを準備
    Set ws = 取得または作成ワークシート("Table001_Page1")
    ws.Cells.Clear
    
    ' テーブルとして読み込み
    With ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & queryName & """;Extended Properties=""""", _
        Destination:=ws.Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table001_Auto"
        .Refresh BackgroundQuery:=False
    End With
    
    Debug.Print "Table001自動作成完了：" & queryName
End Sub

Sub Table002自動作成(pdfPath As String)
    ' -----------------------------------------------
    ' Table002を完全自動で作成（マクロ記録ベース）
    ' -----------------------------------------------
    
    Dim ws As Worksheet
    Dim queryName As String
    Dim formulaText As String
    
    ' 一意のクエリ名を生成
    queryName = "Table002_" & Format(Now, "yyyymmdd_hhnnss")
    
    ' パワークエリの数式を作成（パスを動的に設定）
    formulaText = "let" & Chr(13) & Chr(10) & _
                  "    ソース = Pdf.Tables(File.Contents(""" & pdfPath & """), [Implementation=""1.3""])," & Chr(13) & Chr(10) & _
                  "    Table002 = ソース{[Id=""Table002""]}[Data]," & Chr(13) & Chr(10) & _
                  "    変更された型 = Table.TransformColumnTypes(Table002,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}, {""Column5"", type text}, {""Column6"", type text}, {""Column7"", type text}, {""Column8"", type text}, {""Column9"", type text}, {""Column10"", type text}, {""Column11"", type text}})" & Chr(13) & Chr(10) & _
                  "in" & Chr(13) & Chr(10) & _
                  "    変更された型"
    
    ' クエリを作成
    ActiveWorkbook.Queries.Add Name:=queryName, Formula:=formulaText
    
    ' ワークシートを準備
    Set ws = 取得または作成ワークシート("Table002_Page1")
    ws.Cells.Clear
    
    ' テーブルとして読み込み
    With ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & queryName & """;Extended Properties=""""", _
        Destination:=ws.Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table002_Auto"
        .Refresh BackgroundQuery:=False
    End With
    
    Debug.Print "Table002自動作成完了：" & queryName
End Sub

Function PDFファイル選択ダイアログ() As String
    ' -----------------------------------------------
    ' PDFファイル選択ダイアログ（再利用）
    ' -----------------------------------------------
    
    Dim fileDialog As FileDialog
    Dim result As String
    
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "完全自動処理するPDFファイルを選んでや〜"
        .Filters.Clear
        .Filters.Add "PDFファイル", "*.pdf"
        .FilterIndex = 1
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        
        If .Show = -1 Then
            result = .SelectedItems(1)
            Debug.Print "選択されたPDFファイル：" & result
        Else
            result = ""
        End If
    End With
    
    PDFファイル選択ダイアログ = result
    Set fileDialog = Nothing
End Function

Function 取得または作成ワークシート(sheetName As String) As Worksheet
    ' -----------------------------------------------
    ' ワークシートを取得、なかったら新規作成（再利用）
    ' -----------------------------------------------
    
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
        Debug.Print "新しいワークシート「" & sheetName & "」を作成"
    End If
    
    Set 取得または作成ワークシート = ws
End Function