# excel4eco
エコファーマ申請書類作成支援Excelシート

このリポジトリでは Excel のマクロコードをテキスト管理します。
`.xlsm` ワークブックはこのリポジトリには含まれていないため、`EXCEL_FILE` で使用するファイルを指定してください。
スクリプトは Excel がインストールされた Windows 環境での実行を前提としています。
VBA プロジェクトを抽出するには、Excel のオプションで
「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください。
この設定が無効だと VBS スクリプトからプロジェクトにアクセスできず、
`VBProject is inaccessible` というエラーが表示されます。
VBS 実行時に不明なエラーが発生する場合は、同等の機能を提供する
PowerShell スクリプト `bin/extract_macros.ps1` と `bin/install_macros.ps1`
を利用できます。
Makefile ではこれらのスクリプトを `bin/powershell.sh` 経由で呼び出します。
また、パスワードで保護されたプロジェクトは解除しないと抽出できません。
`make source` で `EXCEL_FILE` からマクロを抽出し、`lib` ディレクトリに保存します。
モジュール名に `\\/:*?"<>|` が含まれる場合、ファイル名として使用できないため
抽出時にこれらの文字はアンダースコアに置き換えられます。
VBA プロジェクトがパスワードで保護されていると
`VBA project is protected` というメッセージが表示され、抽出は行えません。
`make install` で `lib` のコードを `EXCEL_FILE` に組み込みます。

## 使い方
1. `EXCEL_FILE` 変数に対象の `.xlsm` ファイルを指定します。指定しない場合は
   Makefile と同じディレクトリにある `workbook.xlsm` を使用します。
   指定したファイルが存在しない場合、各VBSスクリプトはエラーを表示して終了します。
   パスに空白が含まれる場合は、変数値をダブルクオートで囲んでください。
2. `make source` を実行してマクロを抽出します。内部では PowerShell 用ラッパー
   `bin/powershell.sh` が呼び出されます。直接実行する場合は
   `bin/powershell.sh bin/extract_macros.ps1 "$(EXCEL_FILE)" "$(MACRO_DIR)"`
   を利用してください。
3. `make edit` で `lib` 内の `.bas` `.cls` `.frm` を開きます。`$EDITOR` が未設定の場合は `emacs -nw` が使用されます。
4. `make install` でマクロを再度ファイルへ組み込みます。同様に
   `bin/powershell.sh bin/install_macros.ps1 "$(EXCEL_FILE)" "$(MACRO_DIR)"`
   と実行できます。

### MSYS2 シェルでの利用
MSYS2 環境では Windows 側の `powershell.exe` を直接実行できるよう
`bin/powershell.sh` を用意しています。パス変換も行うため、
`make source` や `make install` をそのまま実行できます。
ターミナルの挙動に問題がある場合は `winpty` を併用してください。

現在のバージョンは `VERSION` ファイルを参照してください。
