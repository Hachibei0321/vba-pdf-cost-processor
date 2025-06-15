' ===============================================
' PDFデータプロセッサー：任意のPDF読み込み・ワークシート作成
' 作成者：関西のおばちゃん
' 作成日：2025/06/15
' 概要：任意のPDFファイルを選択してパワークエリで読み込むで〜♪
' 修正：End Sub抜け修正済み
' ===============================================

Option Explicit

Sub 任意PDF読み込み()
    ' -----------------------------------------------
    ' 任意のPDFファイルを選択してパワークエリで読み込む
    ' メイン処理やで〜♪
    ' -----------------------------------------------
    
    Dim pdfFilePath As String
    Dim errorMessage As String
    
    ' エラーハンドリング設定
    On Error GoTo PDFLoadError
    
    ' 画面更新を止めて処理を高速化
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' ステップ1：PDFファイルを選択
    MsgBox "処理したいPDFファイルを選んでや〜♪", vbInformation, "PDF選択"
    
    pdfFilePath = PDFファイル選択ダイアログ()
    
    If pdfFilePath = "" Then
        MsgBox "PDFファイルが選ばれへんかったから、処理をやめるで〜", vbInformation, "処理中止"
        GoTo CleanUp
    End If
    
    ' ステップ2：パワークエリでPDFを読み込み
    MsgBox "PDFを読み込んでるで〜" & vbCrLf & _
           "ファイル：" & Dir(pdfFilePath) & vbCrLf & _
           "少し時間がかかるかもしれんで〜", vbInformation, "PDF読み込み中"
    
    Call PDFをパワークエリで読み込み(pdfFilePath)
    
    ' ステップ3：テーブルをワークシートに振り分け
    MsgBox "テーブルをワークシートに振り分けてるで〜", vbInformation, "振り分け中"
    
    Call テーブル振り分け()
    
    ' 処理完了
    MsgBox "任意PDFの読み込みが完了したで〜♪" & vbCrLf & _
           "「Table001_Page1」と「Table002_Page1」シートを確認してや", _
           vbInformation, "処理完了"

CleanUp:
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Exit Sub

PDFLoadError:
    errorMessage = "PDF読み込みでエラーが起きたで〜" & vbCrLf & _
                   "エラー内容：" & Err.Description & vbCrLf & _
                   "PDFファイルが壊れてるか、読み込めない形式かもしれんで〜"
    
    MsgBox errorMessage, vbCritical, "PDF読み込みエラー"
    GoTo CleanUp
End Sub

Function PDFファイル選択ダイアログ() As String
    ' -----------------------------------------------
    ' PDFファイル選択ダイアログを表示
    ' 任意のPDFファイルを選んでもらうで〜
    ' -----------------------------------------------
    
    Dim fileDialog As FileDialog
    Dim result As String
    
    ' ファイル選択ダイアログを作成
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    ' ダイアログの設定
    With fileDialog
        .Title = "処理したいPDFファイルを選んでや〜"
        .Filters.Clear
        .Filters.Add "PDFファイル", "*.pdf"
        .FilterIndex = 1
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        
        ' ダイアログを表示
        If .Show = -1 Then
            result = .SelectedItems(1)
            Debug.Print "選択されたPDFファイル：" & result
        Else
            result = ""
            Debug.Print "PDFファイル選択がキャンセルされました"
        End If
    End With
    
    PDFファイル選択ダイアログ = result
    
    ' メモリ解放
    Set fileDialog = Nothing
End Function

Sub PDFをパワークエリで読み込み(pdfFilePath As String)
    ' -----------------------------------------------
    ' 指定されたPDFファイルをパワークエリで読み込む
    ' 動的に接続を作成するで〜
    ' -----------------------------------------------
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' エラーハンドリング
    On Error GoTo PowerQueryError
    
    Set wb = ThisWorkbook
    
    ' PDFデータ用のワークシートを準備
    Set ws = 取得または作成ワークシート("PDFデータ")
    ws.Cells.Clear  ' 既存データをクリア
    
    ' 既存のPDF接続があったら削除
    Call 既存PDF接続削除()
    
    ' パワークエリでPDF読み込み（手動設定方式）
    Call PDFファイル読み込み実行(pdfFilePath, ws)
    
    Debug.Print "PDFファイルの読み込み完了：" & pdfFilePath
    
    Exit Sub

PowerQueryError:
    MsgBox "パワークエリでのPDF読み込みでエラーが起きたで〜" & vbCrLf & _
           "エラー内容：" & Err.Description & vbCrLf & _
           "PDFファイルが読み込める形式か確認してや", vbCritical, "パワークエリエラー"
End Sub

Sub PDFファイル読み込み実行(pdfPath As String, targetWS As Worksheet)
    ' -----------------------------------------------
    ' PDFファイルを実際に読み込む処理
    ' ここはExcelのバージョンによって方法が変わるで〜
    ' -----------------------------------------------
    
    Dim response As VbMsgBoxResult
    
    response = MsgBox("PDFファイルが選択されました：" & vbCrLf & pdfPath & vbCrLf & vbCrLf & _
                      "次の手順で手動設定してくださいな：" & vbCrLf & _
                      "1. データタブ→「データの取得」→「ファイルから」→「PDF から」" & vbCrLf & _
                      "2. 選択されたPDFファイルを開く" & vbCrLf & _
                      "3. Table001とTable002を確認" & vbCrLf & _
                      "4. 「読み込み」ボタンを押す" & vbCrLf & vbCrLf & _
                      "設定が完了したら「OK」を押してや〜", _
                      vbOKCancel + vbInformation, "手動設定のお願い")
    
    If response = vbCancel Then
        MsgBox "処理をキャンセルしたで〜", vbInformation, "キャンセル"
        Exit Sub
    End If
    
    ' パワークエリの完了を待つ
    Application.Wait (Now + TimeValue("0:00:02"))
    
    Debug.Print "PDF読み込み処理完了（手動設定方式）"
End Sub

Sub 既存PDF接続削除()
    ' -----------------------------------------------
    ' 既存のPDF関連接続を削除する
    ' 新しいPDFファイル用に準備するで〜
    ' -----------------------------------------------
    
    Dim conn As WorkbookConnection
    Dim i As Integer
    
    ' 後ろから削除していく（インデックスがずれるのを防ぐため）
    For i = ThisWorkbook.Connections.Count To 1 Step -1
        Set conn = ThisWorkbook.Connections(i)
        
        ' PDF関連やTable関連の接続を削除
        If InStr(LCase(conn.Name), "table") > 0 Or _
           InStr(LCase(conn.Name), "pdf") > 0 Then
            
            Debug.Print "削除する接続：" & conn.Name
            conn.Delete
        End If
    Next i
    
    Debug.Print "既存PDF接続の削除完了"
End Sub

Sub テーブル振り分け()
    ' -----------------------------------------------
    ' Table001とTable002を別々のワークシートに振り分ける
    ' パワークエリで取得したテーブルを整理するで〜
    ' -----------------------------------------------
    
    Dim wsTable1 As Worksheet
    Dim wsTable2 As Worksheet
    Dim wsSource As Worksheet
    Dim lo1 As ListObject
    Dim lo2 As ListObject
    
    ' エラーハンドリング設定
    On Error GoTo SeparateError
    
    ' ソースシートを取得
    Set wsSource = ThisWorkbook.Worksheets("PDFデータ")
    
    ' ワークシートの準備
    Set wsTable1 = 取得または作成ワークシート("Table001_Page1")
    Set wsTable2 = 取得または作成ワークシート("Table002_Page1")
    
    ' 既存データをクリア
    wsTable1.Cells.Clear
    wsTable2.Cells.Clear
    
    ' Table001のデータをコピー
    If パワークエリテーブル存在確認(wsSource, "Table001") Then
        Set lo1 = wsSource.ListObjects("Table001")
        lo1.Range.Copy
        wsTable1.Range("A1").PasteSpecial Paste:=xlPasteValues
        wsTable1.Range("A1").CurrentRegion.AutoFit
        Debug.Print "Table001の振り分け完了"
    Else
        MsgBox "Table001が見つからへんわ〜" & vbCrLf & _
               "パワークエリでPDFを正しく読み込んだか確認してや", vbExclamation, "テーブルなし"
    End If
    
    ' Table002のデータをコピー
    If パワークエリテーブル存在確認(wsSource, "Table002") Then
        Set lo2 = wsSource.ListObjects("Table002")
        lo2.Range.Copy
        wsTable2.Range("A1").PasteSpecial Paste:=xlPasteValues
        wsTable2.Range("A1").CurrentRegion.AutoFit
        Debug.Print "Table002の振り分け完了"
    Else
        MsgBox "Table002が見つからへんわ〜" & vbCrLf & _
               "PDFにテーブルが2つあるか確認してや", vbExclamation, "テーブルなし"
    End If
    
    Application.CutCopyMode = False
    
    MsgBox "テーブルの振り分けが完了したで〜♪", vbInformation, "振り分け完了"
    
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

Function パワークエリテーブル存在確認(ws As Worksheet, tableName As String) As Boolean
    ' -----------------------------------------------
    ' 指定されたワークシートに指定されたテーブルがあるかチェック
    ' -----------------------------------------------
    
    Dim lo As ListObject
    
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    
    パワークエリテーブル存在確認 = (Not lo Is Nothing)
End Function