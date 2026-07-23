# Excel Japanese Katakana Converter with DeepL API

Excel VBA から DeepL API を呼び出し、日本語氏名の読みを全角カタカナで取得して任意のセルへ書き込むサンプルです。

特定のシート名やExcelレイアウトには依存せず、設定欄を変更するだけで対象シート・トリガーセル・氏名セル・出力セルを指定できます。

## 動作

基本フロー:

1. 設定したトリガーセルが変更される
2. 設定した氏名元セルを再計算
3. 氏名を DeepL API に送信
4. 全角カタカナの結果を設定した出力セルへ書き込む

## 設定できる項目

`src/modDeepL.bas` の先頭にある設定欄を変更します。

```vb
' 空欄ならすべてのワークシートが対象
Private Const TARGET_SHEET_NAME As String = ""

Private Const TRIGGER_CELL As String = "A2"
Private Const SOURCE_CELL As String = "B2"
Private Const OUTPUT_CELL As String = "C2"
```

### 対象シート

すべてのワークシートで動作させる場合:

```vb
Private Const TARGET_SHEET_NAME As String = ""
```

特定のシートだけで動作させる場合:

```vb
Private Const TARGET_SHEET_NAME As String = "Sheet1"
```

### セル位置

例として、

- `A2` を変更したときに処理を開始
- `B2` にある氏名を DeepL API へ送信
- `C2` にカタカナ結果を書き込む

という設定になっています。

別のレイアウトでは、次の3つだけ変更してください。

```vb
Private Const TRIGGER_CELL As String = "D5"
Private Const SOURCE_CELL As String = "F5"
Private Const OUTPUT_CELL As String = "G5"
```

氏名元セルが数式の場合も、処理前にそのセルを再計算します。

## ファイル

```text
src/
├─ ThisWorkbook.bas
└─ modDeepL.bas
```

- `ThisWorkbook.bas`  
  `Workbook_SheetChange` イベントを含みます。対象シートやセル位置そのものはここへ直書きせず、`modDeepL.bas` の設定を参照します。

- `modDeepL.bas`  
  対象シート・セル位置の設定、DeepL API 呼び出し、JSON処理、UTF-8変換、結果整形を含みます。

## 導入方法

### 1. ExcelでVBAエディタを開く

`Alt + F11` を押します。

### 2. ThisWorkbookへコードを入れる

`src/ThisWorkbook.bas` の `Option Explicit` 以降を、対象ブックの `ThisWorkbook` に貼り付けます。

> GitHub上の `.bas` 先頭にある `Attribute VB_Name = "ThisWorkbook"` は、手動貼り付け時には不要です。

### 3. 標準モジュールを追加する

VBAエディタで:

`挿入` → `標準モジュール`

モジュール名を `modDeepL` にして、`src/modDeepL.bas` の `Option Explicit` 以降を貼り付けます。

### 4. 対象シートとセル位置を設定する

`modDeepL.bas` の次の項目を、自分のExcelレイアウトに合わせて変更します。

```vb
Private Const TARGET_SHEET_NAME As String = ""
Private Const TRIGGER_CELL As String = "A2"
Private Const SOURCE_CELL As String = "B2"
Private Const OUTPUT_CELL As String = "C2"
```

### 5. DeepL APIキーを設定する

`modDeepL.bas` の次の部分を、自分のDeepL APIキーへ変更します。

```vb
Private Const DEEPL_AUTH_KEY As String = "ここにDeepLのAPIキー"
```

**実際のAPIキーをGitHubへコミットしないでください。**

### 6. DeepL APIエンドポイント

Free API:

```vb
Private Const DEEPL_ENDPOINT As String = "https://api-free.deepl.com/v2/translate"
```

DeepL API Proを使用する場合:

```vb
Private Const DEEPL_ENDPOINT As String = "https://api.deepl.com/v2/translate"
```

## 必要なもの

- Microsoft Excel for Windows
- VBA
- インターネット接続
- DeepL APIキー

VBAでは遅延バインディングを使用しているため、通常はVBEの「参照設定」で追加ライブラリを選択する必要はありません。

## セキュリティ

このリポジトリには実際のDeepL APIキーを含めないでください。

誤ってAPIキーをGitHubへ公開した場合は、Git履歴から削除するだけではなく、そのキーをDeepL側で無効化・再発行してください。

## カスタマイズ例

たとえば `Input` シートの `E3` を変更したとき、`H3` の氏名を読み取り、`I3` にカタカナを入れたい場合:

```vb
Private Const TARGET_SHEET_NAME As String = "Input"
Private Const TRIGGER_CELL As String = "E3"
Private Const SOURCE_CELL As String = "H3"
Private Const OUTPUT_CELL As String = "I3"
```

このため、特定の業務用シート名や固定セル配置に依存せず、通常のExcelブックへ流用できます。
