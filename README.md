# excel4eco
エコファーマ申請書類作成支援Excelシート

このリポジトリでは Excel のマクロコードをテキスト管理します。 
`make source` で `EXCEL_FILE` からマクロを抽出し、`lib` ディレクトリに保存します。
`make install` で `lib` のコードを `EXCEL_FILE` に組み込みます。

## 使い方
1. `EXCEL_FILE` 変数に対象の `.xlsm` ファイルを指定します。
2. `make source` を実行してマクロを抽出します。
3. `lib` 内の `.bas` `.cls` `.frm` を編集します。
4. `make install` でマクロを再度ファイルへ組み込みます。

現在のバージョンは `VERSION` ファイルを参照してください。
