# 📊 VBA PDF原価表自動処理システム

**関西のおばちゃん特製！PDFの原価表を全自動で処理するVBAプロジェクトやで〜♪**

## 🎯 このシステムでできること

PDFの原価表を読み込んで、以下の処理を全自動でやってくれるで〜：

1. **PDFからデータ取得** - パワークエリでテーブル1・2を自動取得
2. **データクリーニング** - 不要な列削除・空白行削除
3. **文字列補正** - 商品名の全角半角修正・定型文字列変換
4. **データ集約** - 商品グループの値を平均化して集約
5. **商品コード追加** - 別ファイルから商品コードを自動照合・追加

## 📁 ファイル構成

```
vba-pdf-cost-processor/
├── README.md              ← この説明ファイル
├── MainProcessor.bas       ← メイン処理（全部ここから実行）
├── PDFDataProcessor.bas    ← PDF処理・データ整理
├── ProductCodeMatcher.bas  ← 商品コード追加処理
└── Utils.bas              ← 共通関数集
```

## 🚀 使い方

### 1. ファイルの準備
- VBA対応のExcelファイル（.xlsm）を用意
- 各.basファイルをVBAエディタにインポート
- パワークエリでPDF接続を事前設定

### 2. メイン処理実行
```vba
Sub メイン原価表処理()
```
↑これを実行するだけで全自動処理開始！

### 3. 必要なシート
処理に必要なシート名：
- `PDFデータ` - パワークエリの結果が入るシート
- `原価リスト` - 最終的な原価リストが入るシート
- `料率リスト` - 商品コード照合用（別ファイル）

## ⚙️ システム要件

- **Excel 2016以降** (パワークエリ対応版)
- **VBA有効** (マクロ実行可能)
- **PDFファイル** (テキストベース推奨)

## 🔧 カスタマイズ

### 列の設定変更
`PDFDataProcessor.bas`で列番号や削除する列を変更できるで〜

### 文字列補正ルール追加
商品名の補正ルールは`商品名文字列補正()`で設定

### 集約条件変更
商品グループの集約条件は`商品グループ集約()`で調整

## 🐛 トラブルシューティング

### よくあるエラー

**「シートが見つからへん」エラー**
→ シート名を確認してや（全角・半角・スペースに注意）

**「パワークエリが更新されへん」エラー**
→ PDF接続設定を確認（データ→クエリと接続）

**「商品コードが追加されへん」エラー**
→ 料率リストファイルの形式を確認（A列:商品コード、B列:商品名）

## 📝 更新履歴

- **2025/06/15** - 初版リリース（関西のおばちゃん製）
- メイン処理システム構築
- 商品コード照合機能追加
- GitHub連携開始

## 💬 サポート

何か分からんことがあったら、Issueに関西弁で気軽に書いてや〜♪

**作成者：関西のおばちゃん**  
**連絡先：GitHubのIssueでお願いします〜**

---

*このプロジェクトは関西のおばちゃんの愛情と関西弁のコメントで作られてます〜♪*
