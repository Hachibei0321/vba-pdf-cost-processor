' ===============================================
' プロシージャ名：AddProductCode（商品コード追加）
' 作成者：関西のおばちゃん
' 作成日：2025/06/15
' 修正日：2025/06/15
' 概要：任意のExcelファイルから商品コードの照合・追加のみを行う処理
'       ※ワークシートのコピーは不要になったで〜♪
' ※変数名は英語、コメントは関西弁で初心者にもわかりやすく♪
' ===============================================

Option Explicit

Sub AddProductCode()
    ' -----------------------------------------------
    ' メイン処理：商品コード追加プロシージャ
    ' シートコピーなしで商品コード照合だけするで〜♪
    ' -----------------------------------------------
    
    ' 変数の宣言（どんな箱を用意するか決めるで〜）
    Dim selectedFilePath As String          ' 選択したファイルのパス
    Dim selectedWorkbook As Workbook        ' 選択したワークブック
    Dim currentWorkbook As Workbook         ' 今開いてるワークブック
    Dim errorMessage As String              ' エラーが起きた時のメッセージ
    
    ' エラーが起きた時の処理を設定（何かあった時の備えやね）
    On Error GoTo ErrorHandler
    
    ' 現在開いているワークブックを覚えておく
    Set currentWorkbook = ThisWorkbook
    
    ' 画面更新を止めて処理を早くする（ちらちらしないようにな）
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' ステップ1：Excelファイルを選択してもらう
    MsgBox "処理を開始するで〜！まずはExcelファイルを選んでや♪", vbInformation, "商品コード追加"
    
    selectedFilePath = ShowFileDialog()
    
    ' ファイルが選ばれなかった場合は処理終了
    If selectedFilePath = "" Then
        MsgBox "ファイルが選ばれへんかったから、処理をやめるで〜", vbInformation, "処理中止"
        GoTo CleanUp
    End If
    
    ' ステップ2：選択したファイルを開く
    MsgBox "ファイルを開くで〜", vbInformation, "処理中"
    Set selectedWorkbook = Workbooks.Open(selectedFilePath)
    
    ' ステップ3：商品コードの照合・追加処理（これがメインやで！）
    MsgBox "商品コードの照合を始めるで〜", vbInformation, "処理中"
    Call MatchAndAddProductCode(selectedWorkbook, currentWorkbook)
    
    ' ステップ4：選択したファイルを閉じる
    selectedWorkbook.Close SaveChanges:=False
    
    ' 処理完了のお知らせ
    MsgBox "お疲れさま〜！商品コード追加処理が完了したで♪", vbInformation, "処理完了"
    
CleanUp:
    ' 画面更新を元に戻す（お片付けやね）
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' メモリを解放する（使った箱をキレイにするねん）
    Set selectedWorkbook = Nothing
    Set currentWorkbook = Nothing
    
    Exit Sub
    
ErrorHandler:
    ' エラーが起きた時の処理
    errorMessage = "あらあら〜エラーが起きてしもたわ！" & vbCrLf & _
                   "エラー番号：" & Err.Number & vbCrLf & _
                   "エラー内容：" & Err.Description
    
    MsgBox errorMessage, vbCritical, "エラー発生"
    
    ' 開いたファイルがあれば閉じる
    If Not selectedWorkbook Is Nothing Then
        selectedWorkbook.Close SaveChanges:=False
    End If
    
    GoTo CleanUp
    
End Sub

' -----------------------------------------------
' ファイル選択ダイアログを表示する関数
' ユーザーにExcelファイルを選んでもらうねん♪
' -----------------------------------------------
Function ShowFileDialog() As String
    
    Dim fileDialog As FileDialog
    Dim result As String
    
    ' ファイル選択ダイアログを作成
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    ' ダイアログの設定（どんなファイルを選ぶか決めるで）
    With fileDialog
        .Title = "処理するExcelファイルを選んでや〜"                    ' タイトル
        .Filters.Clear                                                ' フィルタをクリア
        .Filters.Add "Excelマクロファイル", "*.xlsm"                    ' xlsmファイルのみ
        .FilterIndex = 1                                              ' 最初のフィルタを選択
        .AllowMultiSelect = False                                     ' 複数選択は無し
        .InitialFileName = Application.DefaultFilePath                ' 初期フォルダ
        
        ' ダイアログを表示して、OKボタンが押されたら
        If .Show = -1 Then
            result = .SelectedItems(1)  ' 選択されたファイルのパス
        Else
            result = ""                 ' キャンセルされた
        End If
    End With
    
    ' 結果を返す
    ShowFileDialog = result
    
    ' メモリ解放（使った箱をキレイにする）
    Set fileDialog = Nothing
    
End Function

' -----------------------------------------------
' 商品コード照合処理
' B列の値を比較して、一致したらA列に商品コードを追加するねん
' ※シートコピーなしで、直接参照するで〜♪
' -----------------------------------------------
Sub MatchAndAddProductCode(sourceWorkbook As Workbook, targetWorkbook As Workbook)
    
    Dim costListSheet As Worksheet      ' 原価リストのシート
    Dim rateListSheet As Worksheet      ' 料率リストのシート
    Dim costLastRow As Long             ' 原価リストの最終行
    Dim rateLastRow As Long             ' 料率リストの最終行
    Dim i As Long, j As Long            ' ループ用カウンタ
    Dim matchCount As Long              ' 照合できた件数
    
    ' エラーが起きた時の処理
    On Error GoTo MatchProcessError
    
    ' 料率リストが存在するかチェック（念のため確認するで）
    If Not DoesWorksheetExist(sourceWorkbook, "料率リスト") Then
        MsgBox "選択したファイルに「料率リスト」シートが見つからへんわ〜", vbExclamation, "シートなし"
        Exit Sub
    End If
    
    ' 原価リストが存在するかチェック
    If Not DoesWorksheetExist(targetWorkbook, "原価リスト") Then
        MsgBox "現在のファイルに「原価リスト」シートが見つからへんわ〜", vbExclamation, "シートなし"
        Exit Sub
    End If
    
    ' シートの取得（どのシートを使うか決めるで）
    Set costListSheet = targetWorkbook.Worksheets("原価リスト")
    Set rateListSheet = sourceWorkbook.Worksheets("料率リスト")
    
    ' 最終行を取得（B列の最後のデータがある行を探すで）
    costLastRow = costListSheet.Cells(costListSheet.Rows.Count, "B").End(xlUp).Row
    rateLastRow = rateListSheet.Cells(rateListSheet.Rows.Count, "B").End(xlUp).Row
    
    ' デバッグ情報を出力（どこまでデータがあるか確認）
    Debug.Print "原価リスト最終行：" & costLastRow
    Debug.Print "料率リスト最終行：" & rateLastRow
    
    matchCount = 0  ' 照合件数をゼロからスタート
    
    ' 原価リストの各行をチェック（2行目から開始、1行目はヘッダーやからな）
    For i = 2 To costLastRow
        
        ' 原価リストのB列の値を取得（前後の空白は削除）
        Dim costBValue As String
        costBValue = Trim(CStr(costListSheet.Cells(i, "B").Value))
        
        ' 空白セルなら処理終了（データが終わったってことやね）
        If costBValue = "" Then
            Exit For
        End If
        
        ' 料率リストで同じ値を探す（ここからが照合作業や）
        For j = 2 To rateLastRow
            
            Dim rateBValue As String
            rateBValue = Trim(CStr(rateListSheet.Cells(j, "B").Value))
            
            ' 空白セルなら次の原価リスト行へ
            If rateBValue = "" Then
                Exit For
            End If
            
            ' 値が一致した場合（やった！見つかった♪）
            If costBValue = rateBValue Then
                
                ' 料率リストのA列の値を取得
                Dim rateAValue As String
                rateAValue = Trim(CStr(rateListSheet.Cells(j, "A").Value))
                
                ' 原価リストのA列にコピー（ここが一番大事な処理や）
                costListSheet.Cells(i, "A").Value = rateAValue
                
                ' 照合件数をカウントアップ
                matchCount = matchCount + 1
                
                ' デバッグ情報を出力
                Debug.Print "照合完了：" & costBValue & " → " & rateAValue
                
                ' 一致したら次の原価リスト行へ（同じ値を何度も処理せんでええからな）
                Exit For
                
            End If
            
        Next j
        
    Next i
    
    ' 結果をお知らせ（何件処理できたか教えたるで）
    MsgBox matchCount & " 件の商品コードを追加したで〜♪", vbInformation, "照合結果"
    
    Exit Sub
    
MatchProcessError:
    ' 照合処理でエラーが起きた場合
    MsgBox "商品コード照合でエラーが起きたで〜" & vbCrLf & Err.Description, vbCritical, "照合エラー"
    
End Sub

' -----------------------------------------------
' ワークシートの存在チェック関数
' 指定されたシートがあるかどうか調べるねん
' -----------------------------------------------
Function DoesWorksheetExist(targetWorkbook As Workbook, sheetName As String) As Boolean
    
    Dim ws As Worksheet
    
    ' エラーが起きても処理を続ける（シートがなくてもエラーにならんようにな）
    On Error Resume Next
    
    Set ws = targetWorkbook.Worksheets(sheetName)
    
    ' シートが見つかったかどうかチェック
    If Not ws Is Nothing Then
        DoesWorksheetExist = True   ' シートがある
    Else
        DoesWorksheetExist = False  ' シートがない
    End If
    
    ' エラー処理を元に戻す
    On Error GoTo 0
    
    ' メモリ解放
    Set ws = Nothing
    
End Function