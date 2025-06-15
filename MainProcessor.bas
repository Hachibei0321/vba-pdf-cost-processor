' ===============================================
' メインプロセッサー：原価表PDF自動処理システム
' 作成者：関西のおばちゃん
' 作成日：2025/06/15
' 概要：PDFから原価表を取得して全自動で処理するシステムやで〜♪
' 
' 処理フロー：
' 1. PDFからパワークエリでテーブル1・2取得
' 2. 列削除・空白行削除
' 3. 商品名の全角半角修正・文字列変換
' 4. 商品グループの値を平均化して集約
' 5. 商品コード追加処理
' ===============================================

Option Explicit

Sub メイン原価表処理()
    ' -----------------------------------------------
    ' 原価表PDF自動処理のメインプロシージャ
    ' 全ての処理をここから実行するで〜♪
    ' -----------------------------------------------
    
    Dim startTime As Date
    Dim errorMessage As String
    
    ' 処理開始時刻を記録
    startTime = Now
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 画面更新を止めて処理を高速化
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' 処理開始のお知らせ
    MsgBox "原価表の自動処理を開始するで〜！" & vbCrLf & _
           "しばらくお待ちくださいな〜♪", vbInformation, "処理開始"
    
    ' ステップ1：PDFからデータ取得
    Call ステップ1_PDF取得()
    
    ' ステップ2：データクリーニング
    Call ステップ2_データクリーニング()
    
    ' ステップ3：文字列補正
    Call ステップ3_文字列補正()
    
    ' ステップ4：データ集約
    Call ステップ4_データ集約()
    
    ' ステップ5：商品コード追加
    Call ステップ5_商品コード追加()
    
    ' 処理完了
    Dim processingTime As Long
    processingTime = DateDiff("s", startTime, Now)
    
    MsgBox "お疲れさま〜！" & vbCrLf & _
           "原価表処理が完了したで♪" & vbCrLf & _
           "処理時間：" & processingTime & "秒", vbInformation, "処理完了"

CleanUp:
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub

ErrorHandler:
    errorMessage = "あらあら〜処理中にエラーが起きてしもたわ！" & vbCrLf & _
                   "エラー番号：" & Err.Number & vbCrLf & _
                   "エラー内容：" & Err.Description & vbCrLf & _
                   "処理時間：" & DateDiff("s", startTime, Now) & "秒"
    
    MsgBox errorMessage, vbCritical, "エラー発生"
    
    GoTo CleanUp
End Sub

Sub ステップ1_PDF取得()
    ' -----------------------------------------------
    ' ステップ1：PDFからパワークエリでデータ取得
    ' -----------------------------------------------
    
    MsgBox "ステップ1：PDFからデータを取得中やで〜", vbInformation, "処理中"
    
    ' パワークエリを更新（事前にPDF接続を設定しておく必要があるで）
    Call PDF更新()
    
    ' テーブルを別シートに振り分け
    Call テーブル振り分け()
    
End Sub

Sub ステップ2_データクリーニング()
    ' -----------------------------------------------
    ' ステップ2：列削除・空白行削除
    ' -----------------------------------------------
    
    MsgBox "ステップ2：データをキレイにしてるで〜", vbInformation, "処理中"
    
    ' 各シートのデータクリーニング
    Call テーブル1クリーニング()
    Call テーブル2クリーニング()
    
End Sub

Sub ステップ3_文字列補正()
    ' -----------------------------------------------
    ' ステップ3：商品名の全角半角修正・文字列変換
    ' -----------------------------------------------
    
    MsgBox "ステップ3：文字列を補正してるで〜", vbInformation, "処理中"
    
    ' 文字列補正処理
    Call 商品名文字列補正()
    
End Sub

Sub ステップ4_データ集約()
    ' -----------------------------------------------
    ' ステップ4：商品グループの値を平均化して集約
    ' -----------------------------------------------
    
    MsgBox "ステップ4：データを集約してるで〜", vbInformation, "処理中"
    
    ' データ集約処理
    Call 商品グループ集約()
    
End Sub

Sub ステップ5_商品コード追加()
    ' -----------------------------------------------
    ' ステップ5：商品コード追加処理
    ' -----------------------------------------------
    
    MsgBox "ステップ5：商品コードを追加してるで〜", vbInformation, "処理中"
    
    ' 商品コード追加処理（ProductCodeMatcher.basの関数を呼び出し）
    Call AddProductCode()
    
End Sub