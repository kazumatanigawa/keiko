# GAS Backend

`gas/Code.gs` は、`index.html` から呼ばれる Web App 用の Apps Script 実装です。

## 想定

- スプレッドシートは既存運用を継続する
- 既存データは削除しない
- 初回実行時にバックアップ、ID採番、`teamId`/`userId` 補完、`部員一覧` 再構築を行う

## デプロイ手順

1. Apps Script プロジェクトを開く
2. `Code.gs` をこの内容で置き換える
3. `appsscript.json` を反映する
4. 必要ならスクリプトプロパティ `SPREADSHEET_ID` を設定する
5. Web App を再デプロイする

`SPREADSHEET_ID` を設定しない場合は、コンテナバインドされたスプレッドシートを使います。
