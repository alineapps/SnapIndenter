# VBA Formatter

Excel / Access / Word などの Visual Basic Editor (VBE) から直接使える、軽量な VBA コード整形アドインです。
詳しい仕様・インストール方法は [docs/仕様書.md](docs/仕様書.md) を参照してください。
SmartIndenter 相当のインデント整形に機能を絞っているので、起動はほぼ一瞬です。

## インストール

1. Excel / Access などの Office アプリをすべて終了します。
2. `install.bat` をダブルクリックします。
   - ユーザー単位で登録するため、管理者権限は不要です。
   - 32bit 版／64bit 版 Office のどちらにも対応しています。
   - 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」の設定は不要です。
3. Office を起動して `Alt+F11` で VBE を開くと、メニューバーに **整形(J)** が追加されています。

アンインストールするときは `uninstall.bat` を実行します。

> 動作要件: Windows 10 / 11（.NET Framework 4.8 は標準で入っています）、Office 2010 以降

## 使い方

| 操作 | メニュー | 既定のショートカット |
|---|---|---|
| カーソル位置のプロシージャを整形（複数行を選択しているときは選択範囲） | 整形 → プロシージャ/選択範囲を整形 | `Ctrl+Shift+F` |
| 表示中のモジュール全体を整形 | 整形 → モジュールを整形 | （未割り当て） |
| アクティブなプロジェクトの全モジュールを整形 | 整形 → プロジェクト全体を整形 | |
| 直前の整形を元に戻す | 整形 → 直前の整形を元に戻す | |
| 設定（プレビュー付き） | 整形 → 設定 | |

コード ウィンドウの右クリックメニューからも実行できます。

※ VBE の「元に戻す」(Ctrl+Z) ではアドインによる変更を戻せません。整形メニューの「直前の整形を元に戻す」を使ってください。

## 整形内容

- ブロック構造に合わせたインデント
  `Sub/Function/Property`、`If/ElseIf/Else`、`Select Case/Case`、`For/For Each/Next`（`Next i, j` も対応）、
  `Do/Loop`、`While/Wend`、`With`、`Type`、`Enum`、`#If/#Else/#End If`
- 1 行 If（`If x Then y`）はブロックとして扱わない
- 行ラベル・行番号は 1 桁目に置く
- 行継続（` _`）の 2 行目以降は、1 行目からの相対位置を保ったまま移動する
- 空白だけの行と行末の空白を削除する
- 文字列やコメントの中のキーワードは無視する
- 行数は変わりません（ブレークポイントやブックマークの位置はずれません）

空白の入れ方やキーワードの大文字・小文字は、VBE が自動で直すため触りません。

## 設定

設定ダイアログで変更できます。設定は `%APPDATA%\VbaFormatter\settings.ini` に保存されます。

- インデント幅 / タブ使用
- プロシージャ本体・Case・Type/Enum メンバー・`#If` の中をインデントするか
- コメント行をインデントするか
- 継続行の相対位置を保持するか、継続行に追加するインデントの段数
- ショートカットキー（例: `Ctrl+Shift+F`、`Ctrl+Alt+I`、`F12`）

## 開発

```
src/VbaFormatter.Core    整形エンジン（Office に依存しない。netstandard2.0）
src/VbaFormatter.AddIn   VBE 用 COM アドイン（net48。Core のソースも取り込んで 1 つの DLL にする）
tests/VbaFormatter.Tests 整形エンジンの単体テスト（xUnit）
installer/               インストール・アンインストール用スクリプト
```

```powershell
.\build.ps1      # テスト → Release ビルド → dist\VbaFormatter-x.y.z.zip を作成
```

開発中の DLL をインストールするときは `installer\install.ps1 -Source src\VbaFormatter.AddIn\bin\Release\net48` を実行します。
ログは `%LOCALAPPDATA%\VbaFormatter\VbaFormatter.log` に出力されます。
