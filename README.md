# google-drive-auto-conversion

Google ドライブにアップロードしたPDF・画像・Officeファイルを、自動でGoogleスプレッドシート / ドキュメントに変換し、取引台帳に記録するツール。

## 初期設定（3ステップ）

### 1. コードをデプロイ

```bash
clasp push
```

### 2. GASエディタで `setup()` を実行

Apps Script エディタを開き、関数選択で `setup` を選んで実行ボタンを押す。
初回は権限承認のダイアログが出るので承認する。

これだけで以下がすべて完了する:

- マイドライブに「Googleドライブ自動変換」フォルダ階層を作成
  - `UPLOAD` / `処理済み` / `出力` / `取引台帳`
- 各IDをScriptPropertiesに自動保存（**`config.gs` の編集は不要**）
- 5分間隔の自動実行トリガーを登録

何度実行しても重複は作られない（冪等）。

### 3. 完了

マイドライブの「Googleドライブ自動変換 / UPLOAD」フォルダにファイルを入れるだけで、5分以内に自動変換されます。

## 使い方

- **自動処理**: UPLOAD フォルダにファイルを入れれば、トリガーが自動で変換・台帳記録・処理済み移動を行います。
- **手動処理**: 特定のファイルだけ即座に処理したい場合は、GASエディタで `processManual('ファイルID')` を実行。
- **動作確認**: `diagnose()` を実行すると、設定状況とUPLOADフォルダ内のファイル一覧が確認できます。

## トラブルシューティング

- **「初期セットアップ未完了」と出る** → `setup()` を実行してください。
- **トリガーを止めたい** → `removeTrigger()` を実行。
- **フォルダIDをリセットしたい** → Apps Script エディタ「プロジェクトの設定 → スクリプトプロパティ」で該当キーを削除し、`setup()` を再実行。

## 設定のカスタマイズ

[config.gs](config.gs) でカスタマイズ可能な項目:

- `trigger.intervalMinutes`: トリガー実行間隔（デフォルト5分）
- `processing.maxFilesPerExecution`: 1回の実行で処理する最大ファイル数（デフォルト5）
- `notification.recipientEmail`: 通知メール宛先（空ならログイン中ユーザー）
- `ocr.language`: OCR言語（デフォルト `ja`）

フォルダIDと台帳IDは ScriptProperties から動的に読み込まれるため、`config.gs` に書く必要はありません。
