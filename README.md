# Excel Japanese Katakana Converter with DeepL API

Excel VBA から DeepL API を呼び出し、氏名の読みを全角カタカナで取得してセルへ書き込むサンプルです。

## 動作

対象シート:

- `雇入れ`
- `離職`

処理フロー:

1. `AZ9` が変更される
2. 数式セル `K14` を再計算
3. `K14` の氏名を DeepL API に送信
4. 結果を `K12` に書き込む

## ファイル

```text
src/
├─ ThisWorkbook.bas
└─ modDeepL.bas
```

- `ThisWorkbook.bas`  
  `Workbook_SheetChange` イベントを含みます。

- `modDeepL.bas`  
  DeepL API 呼び出し、JSON処理、UTF-8変換、結果整形を含みます。

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

### 4. DeepL APIキーを設定する

`modDeepL.bas` の次の部分を、自分のDeepL APIキーへ変更します。

```vb
Private Const DEEPL_AUTH_KEY As String = "ここにDeepLのAPIキー"
```

**実際のAPIキーをGitHubへコミットしないでください。**

### 5. DeepL APIエンドポイント

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

## 注意

このコードは特定のExcelレイアウト向けです。

- トリガー: `AZ9`
- 氏名: `K14`
- 出力: `K12`
- 対象シート: `雇入れ`, `離職`

別のブックで使う場合はセル番地やシート名を変更してください。
