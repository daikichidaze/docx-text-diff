# Windows 11 Word backend implementation plan

作成日: 2026-06-04

## 前提

この計画は、Windows 11 環境で Microsoft Word がインストール済みであることを前提にする。

目的は、2つの `.docx` ファイルを Microsoft Word の比較機能で比較し、差分をWordの変更履歴として含む1つの `.docx` ファイルを生成するCLIスクリプトを開発すること。

## 採用方針

Microsoft Word の COM Automation から `Application.CompareDocuments` を呼び出す方式を採用する。

この方式では、差分検出と変更履歴マークアップ生成をWord本体に委ねる。スクリプト側は、入力ファイルの検証、Word起動、比較オプション指定、出力保存、後片付けを担当する。

## 技術スタック

- OS: Windows 11
- Runtime: Python 3.11 以上
- Word Automation: `pywin32`
- CLI: Python標準の `argparse`
- パス処理: `pathlib`
- ログ: Python標準の `logging`
- テスト: `pytest`

## CLI仕様

コマンド名案:

```powershell
python -m docx_diff_track original.docx revised.docx output.docx
```

オプション案:

```powershell
python -m docx_diff_track original.docx revised.docx output.docx `
  --granularity word `
  --author "Comparison" `
  --compare-formatting `
  --compare-case `
  --compare-whitespace `
  --compare-tables `
  --compare-headers `
  --compare-footnotes `
  --compare-textboxes `
  --compare-fields `
  --compare-comments `
  --compare-moves `
  --visible false
```

### 引数

| 引数 | 必須 | 説明 |
| --- | --- | --- |
| `original` | yes | 比較元DOCX |
| `revised` | yes | 比較先DOCX |
| `output` | yes | 変更履歴付きDOCXの出力先 |

### オプション

| オプション | 既定値 | 説明 |
| --- | --- | --- |
| `--granularity word|char` | `word` | Wordの比較粒度 |
| `--author TEXT` | `DocxDiff` | 変更履歴の作成者名 |
| `--compare-formatting / --no-compare-formatting` | `false` | 書式差分を比較する |
| `--compare-case / --no-compare-case` | `true` | 大文字小文字の差分を比較する |
| `--compare-whitespace / --no-compare-whitespace` | `true` | 空白差分を比較する |
| `--compare-tables / --no-compare-tables` | `true` | 表を比較する |
| `--compare-headers / --no-compare-headers` | `true` | ヘッダー/フッターを比較する |
| `--compare-footnotes / --no-compare-footnotes` | `true` | 脚注/文末脚注を比較する |
| `--compare-textboxes / --no-compare-textboxes` | `true` | テキストボックスを比較する |
| `--compare-fields / --no-compare-fields` | `true` | フィールドを比較する |
| `--compare-comments / --no-compare-comments` | `true` | コメントを比較する |
| `--compare-moves / --no-compare-moves` | `true` | 移動を検出する |
| `--ignore-all-comparison-warnings` | `false` | Wordの比較警告を無視する |
| `--visible true|false` | `false` | Word UIを表示する |
| `--overwrite` | `false` | 出力ファイルが存在する場合に上書きする |

## ディレクトリ構成案

```text
docx_diff/
  pyproject.toml
  README.md
  src/
    docx_diff_track/
      __init__.py
      __main__.py
      cli.py
      word_compare.py
      errors.py
  tests/
    test_cli_args.py
    test_path_validation.py
    test_word_compare_smoke.py
  docs/
    docx_tracked_changes_research.md
    windows_word_backend_implementation_plan.md
```

## 実装設計

### `cli.py`

責務:

- CLI引数の定義
- 入力ファイル存在チェック
- 拡張子チェック
- 出力ファイル上書き可否チェック
- ログレベル設定
- `word_compare.compare_documents()` の呼び出し
- 例外をユーザー向けメッセージに変換

### `word_compare.py`

責務:

- COM初期化
- Word Application 起動
- Word定数の定義
- 文書を読み取り専用で開く
- `CompareDocuments` 呼び出し
- 比較結果を `.docx` として保存
- すべてのDocumentを閉じる
- Word Applicationを終了する

### `errors.py`

責務:

- ユーザー向け例外クラスの定義
- Word未インストール、COM起動失敗、入力不正、出力失敗などを分類する

## Word COM処理フロー

1. `pythoncom.CoInitialize()` を呼ぶ。
2. `win32com.client.DispatchEx("Word.Application")` で独立したWordインスタンスを起動する。
3. `Application.Visible` をCLI指定に合わせる。
4. `Application.DisplayAlerts = 0` を設定する。
5. 元文書を `Documents.Open(..., ReadOnly=True, AddToRecentFiles=False)` で開く。
6. 改訂文書を同様に開く。
7. `Application.CompareDocuments(original_doc, revised_doc, ...)` を呼ぶ。
8. 戻り値の比較結果Documentを `SaveAs2(output_path, FileFormat=wdFormatXMLDocument)` で保存する。
9. 比較結果Document、元文書、改訂文書を保存せずに閉じる。
10. Word Applicationを終了する。
11. `pythoncom.CoUninitialize()` を呼ぶ。

## Word定数

pywin32の定数取得は環境差が出ることがあるため、必要最小限の定数は自前で定義する。

```python
WD_COMPARE_TARGET_NEW = 2
WD_GRANULARITY_CHAR_LEVEL = 0
WD_GRANULARITY_WORD_LEVEL = 1
WD_FORMAT_XML_DOCUMENT = 12
WD_DO_NOT_SAVE_CHANGES = 0
```

## `CompareDocuments` 呼び出し案

```python
comparison = word.CompareDocuments(
    OriginalDocument=original_doc,
    RevisedDocument=revised_doc,
    Destination=WD_COMPARE_TARGET_NEW,
    Granularity=granularity,
    CompareFormatting=compare_formatting,
    CompareCaseChanges=compare_case,
    CompareWhitespace=compare_whitespace,
    CompareTables=compare_tables,
    CompareHeaders=compare_headers,
    CompareFootnotes=compare_footnotes,
    CompareTextboxes=compare_textboxes,
    CompareFields=compare_fields,
    CompareComments=compare_comments,
    CompareMoves=compare_moves,
    RevisedAuthor=author,
    IgnoreAllComparisonWarnings=ignore_all_comparison_warnings,
)
```

## 既定値の考え方

初期既定値は「テキスト差分を見落とさない」ことを優先する。

- 粒度は `word`
  - 英語やスペース区切り文書で読みやすい。
  - 日本語文書で差分が粗い場合は `char` を指定できるようにする。
- 書式比較は `false`
  - 今回の主目的はテキスト差分の分析。
  - 書式差分を含めると変更履歴が増え、読みにくくなることがある。
- 表、ヘッダー、脚注、テキストボックス等は `true`
  - Word backendではWord本体に任せられるため、対象から外す必要は薄い。

## エラーハンドリング

想定するエラー:

- PythonがWindows以外で実行された
- `pywin32` が未インストール
- Microsoft Wordが未インストール
- COM起動に失敗した
- 入力ファイルが存在しない
- 入力ファイルが `.docx` ではない
- 入力ファイルがWordで開けない
- 出力先が既に存在し、`--overwrite` が指定されていない
- 出力先ディレクトリが存在しない
- Word比較時に警告またはエラーが発生した
- Wordプロセス終了時に例外が発生した

設計方針:

- 入力検証で検出できるものはWord起動前に止める。
- Word起動後は `try/finally` で必ず文書を閉じる。
- 例外発生時も `word.Quit()` を試みる。
- `--visible true` の場合でも、スクリプト終了時は明示的に閉じる。

## テスト計画

### Windows実環境で行うテスト

- 最小DOCX 2ファイルを比較し、出力DOCXが生成される。
- 出力DOCXをWordで開くと変更履歴が表示される。
- 単語追加、単語削除、置換が変更履歴になる。
- `--granularity char` で日本語の1文字差分が確認できる。
- `--no-compare-formatting` で書式のみの差分が無視される。
- `--overwrite` なしで既存出力を上書きしない。
- Word起動中に例外が起きてもWordプロセスが残りにくい。

### 自動テスト

COM依存部分は通常CIで安定しにくいため、層を分ける。

- `test_cli_args.py`: CLI引数のパースをテストする。
- `test_path_validation.py`: 入出力パス検証をテストする。
- `test_word_compare_smoke.py`: Windows + Word がある場合のみ実行するsmoke testにする。

`pytest` の marker 例:

```python
@pytest.mark.windows_word
def test_compare_documents_smoke(tmp_path):
    ...
```

## 実装ステップ

1. `pyproject.toml` を作成し、Pythonパッケージの骨格を用意する。
2. `src/docx_diff_track/cli.py` にCLI引数とパス検証を実装する。
3. `src/docx_diff_track/word_compare.py` にWord COM比較処理を実装する。
4. `src/docx_diff_track/__main__.py` からCLIを起動できるようにする。
5. CLI・パス検証の単体テストを追加する。
6. 手動確認用のサンプルDOCX作成手順をREADMEに追加する。
7. Windows 11 + Word環境でsmoke testを行う。
8. 出力DOCXをWordで開き、変更履歴表示を確認する。

## 運用上の注意

MicrosoftはOfficeアプリケーションのサーバーサイド無人自動化を一般に推奨していない。今回の方式は、個人PCや操作端末上でのバッチ処理・ローカルCLIとして使う前提にするのが安全。

大量ファイルを処理する場合は、次の対策が必要。

- 1件ごとにWordインスタンスを閉じるか、一定件数ごとに再起動する。
- タイムアウトを設ける。
- 入力ファイルをローカル一時ディレクトリへコピーしてから処理する。
- 処理失敗時に対象ファイル名とWordエラーをログに残す。
- 並列実行は避ける。Word COMは複数プロセス並列処理に向かない。

## 今回の実装で作らないもの

- Wordなし環境向けの純Python OOXML差分生成
- Web API化
- 複数ファイル一括処理
- GUI
- Wordの比較結果を独自に解析するレポート機能
- 既存変更履歴を事前accept/rejectする機能

## 参考資料

- Microsoft Learn: Application.CompareDocuments method (Word)  
  https://learn.microsoft.com/en-us/office/vba/api/word.application.comparedocuments
- Microsoft Learn: Document.SaveAs2 method (Word)  
  https://learn.microsoft.com/en-us/office/vba/api/word.saveas2
- Microsoft Learn: Documents.Open method (Word)  
  https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
- Microsoft Learn: Application.Quit method (Word)  
  https://learn.microsoft.com/en-us/office/vba/api/word.application.quit
- pywin32 project  
  https://github.com/mhammond/pywin32
