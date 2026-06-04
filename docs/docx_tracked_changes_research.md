# DOCX text diff to Word tracked-changes DOCX: research and design

作成日: 2026-06-04

## 目的

2つの `.docx` ファイルを入力として、本文テキスト差分を分析し、Microsoft Word の「変更履歴」として表示できる1つの `.docx` ファイルを生成するスクリプトを新規開発する。

## 結論

技術的には可能。

ただし「Wordの変更履歴として表示できるDOCXを作る」には2つの意味がある。

1. Word自身の比較エンジンを呼び出して、Wordが生成する変更履歴DOCXを保存する。
2. スクリプト側でDOCX内部のOOXMLを編集し、`w:ins` / `w:del` などの変更履歴マークアップを自前で生成する。

実用性・互換性・開発コストの観点では、まずは **Word COM Automation を使う方式** が最有力。Windows + Microsoft Word が使える環境なら、Wordの `Application.CompareDocuments` は「2文書を比較し、差分を変更履歴として含むDocumentを返す」と公式に説明されている。

Linux/macOSやMicrosoft Wordなしで動かす必要がある場合は、純Python/OOXML方式が候補になる。ただし、最初の実装範囲は「本文段落のテキスト差分」に限定するのが現実的。表、ヘッダー/フッター、脚注、文末脚注、テキストボックス、フィールド、コメント、移動、書式変更までWord同等に扱うのは別プロジェクト級になる。

## 調査結果

### DOCX変更履歴の構造

`.docx` はZIPパッケージで、本文は主に `word/document.xml` の WordprocessingML として格納される。基本構造は、文書 `document`、本文 `body`、段落 `p`、ラン `r`、テキスト `t` で構成される。

変更履歴はWordprocessingML上の専用要素で表現される。代表例は次の通り。

- 挿入: `w:ins`
- 削除: `w:del`
- 削除文字列: `w:delText`
- 段落プロパティ変更: `w:pPrChange`
- ランプロパティ変更: `w:rPrChange`
- 移動元/移動先: `w:moveFrom`, `w:moveTo`
- 変更履歴記録設定: `w:trackRevisions`

Microsoft Learn の Open XML SDK ドキュメントでは、`w:trackRevisions` が有効な場合、挿入テキストは `w:ins`、太字などの書式変更は `w:rPrChange` として出力される例が示されている。

### Word本体による比較

Microsoft Word の `Application.CompareDocuments` は、元文書と改訂文書を比較し、差分が変更履歴としてマークされたDocumentを返す。比較粒度は単語単位または文字単位を選べ、書式、大文字小文字、空白、表、ヘッダー/フッター、脚注、テキストボックス、フィールド、コメント、移動などの比較オプションも持つ。

この方式はWord自身の結果を得られるため、Wordで開いたときの互換性が最も高い。一方で、Windows + Microsoft Word が必要で、サーバーサイドの無人実行にはライセンス・安定性・プロセス管理の注意が必要。

### python-docxの制約

`python-docx` は `.docx` の作成・編集には便利だが、変更履歴の読み書きAPIとしては不十分。公式ドキュメント上も通常の文書操作が中心で、変更履歴をWord同等に扱うAPIは提供されていない。既存の解説や実装例でも、`python-docx` は `w:ins` / `w:del` 内の段落やランを通常APIで扱いにくい点が指摘されている。

ただし、`python-docx` が使えないという意味ではない。ZIP/XMLとして直接 `word/document.xml` を操作する、または変更履歴対応をうたうライブラリを併用すれば、限定的な変更履歴DOCXは生成できる。

### docx-revisions等のライブラリ

`docx-revisions` は `python-docx` を拡張し、OOXMLの挿入・削除マークアップを扱うライブラリとして公開されている。ドキュメントでは、変更履歴の読み取り、accept/reject、検索置換時の変更履歴付与、挿入・削除の追加APIが紹介されている。

ただし、今回必要なのは「2つの文書を比較して、整合する位置に変更履歴を構成する」ことであり、単なる検索置換や単一段落編集より難しい。ライブラリを使う場合も、差分抽出、段落対応付け、ラン/書式保持、XML再構成の設計が必要。

## 実装方式の比較

| 方式 | 概要 | 長所 | 短所 | 推奨度 |
| --- | --- | --- | --- | --- |
| A. Word COM Automation | Microsoft Wordの比較機能を自動実行 | Word互換性が最良。表、脚注、コメント等もWordの比較機能に任せられる | Windows + Word必須。サーバー運用に注意 | 高 |
| B. LibreOffice headless | LibreOfficeのCLI/UNOで比較・変換を試みる | Wordなしで動く可能性 | Wordの変更履歴DOCXとしての比較品質・API安定性を要検証 | 中低 |
| C. 純Python OOXML生成 | `.docx` をZIP/XMLとして編集し `w:ins` / `w:del` を生成 | OS非依存。配布しやすい。ロジックを制御できる | Word同等の比較は困難。初期実装範囲を絞る必要 | 中 |
| D. docx-revisions等を活用 | 変更履歴XML生成をライブラリに委譲 | 自前XML操作を一部減らせる | 文書比較全体は別途実装。成熟度の検証が必要 | 中 |

## 推奨アーキテクチャ

最初の実用版は、環境別に2つのバックエンドを持つCLIとして設計する。

1. `word` backend: Windows + Microsoft Word がある環境向け。Wordの `CompareDocuments` を呼び出す。
2. `ooxml-text` backend: Wordなし環境向け。本文テキストに限定し、OOXMLに変更履歴を直接挿入する。

CLI例:

```bash
docx-diff-track original.docx revised.docx output.docx \
  --backend word \
  --granularity word \
  --author "Comparison" \
  --compare-formatting false
```

純PythonバックエンドのCLI例:

```bash
docx-diff-track original.docx revised.docx output.docx \
  --backend ooxml-text \
  --granularity char \
  --author "Comparison"
```

## Word backend 設計

### 処理フロー

1. 入力パスを絶対パスに解決する。
2. Word COMを起動する。
3. 元文書と改訂文書を読み取り専用で開く。
4. `Application.CompareDocuments` を呼び出す。
5. 返された比較結果Documentを `output.docx` として保存する。
6. 開いた文書を閉じ、Wordプロセスを終了する。

### Python実装候補

- Windows: `pywin32`
- macOS: AppleScript/JXAでWord操作は可能だが、COMほど安定した自動化APIではないため別検証
- Linux: Word backend不可

### 設定項目

- `--granularity word|char`
- `--compare-formatting`
- `--compare-case`
- `--compare-whitespace`
- `--compare-tables`
- `--compare-headers`
- `--compare-footnotes`
- `--compare-textboxes`
- `--compare-fields`
- `--compare-comments`
- `--compare-moves`
- `--author`
- `--ignore-warnings`

## 純Python OOXML text backend 設計

### 対象範囲

初期版は以下に限定する。

- `word/document.xml` の本文段落
- 通常段落内のテキストラン
- 文字単位または単語単位の挿入・削除
- 改訂後文書をベースにし、削除箇所を `w:del`、挿入箇所を `w:ins` として埋め込む

初期版では対象外にする。

- 書式変更の検出
- 表の構造変更
- ヘッダー/フッター
- 脚注/文末脚注
- コメント
- フィールドコード
- テキストボックス
- 図形内テキスト
- 移動検出
- 既存変更履歴の完全保持/統合

### 処理フロー

1. 入力 `.docx` をZIPとして読み込む。
2. `word/document.xml` をXMLとして解析する。
3. 元文書と改訂文書から、段落単位のテキスト列を抽出する。
4. 段落対応付けを行う。
   - 基本は段落列に対するdiff。
   - 類似度が高い段落は同一段落の変更として扱う。
   - 類似度が低い段落は段落削除 + 段落挿入として扱う。
5. 対応した段落内でトークンdiffを行う。
   - `word` 粒度: 単語、句読点、空白をトークン化。
   - `char` 粒度: Unicodeコードポイント単位。ただし結合文字やサロゲートには注意。
6. 改訂後文書側の段落XMLをベースに、差分に応じてランを再構成する。
7. 挿入範囲を `w:ins w:id w:author w:date` で包む。
8. 削除範囲を `w:del w:id w:author w:date` で包み、削除文字列は `w:delText` に変換する。
9. `word/settings.xml` に `w:trackRevisions` を設定する。
10. ZIP内の該当XMLを差し替えて `output.docx` を生成する。

### 差分アルゴリズム

Python標準の `difflib.SequenceMatcher` から開始する。初期版には十分だが、長文契約書や条項移動が多い文書ではWordの比較結果とずれる可能性がある。

将来的に改善する場合:

- patience diff
- histogram diff
- 段落類似度によるマッチング
- 日本語向け形態素解析またはUnicode-aware tokenization

### XML生成上の注意

- `w:ins` / `w:del` には一意の `w:id`、`w:author`、`w:date` を付与する。
- 削除テキストは通常の `w:t` ではなく `w:delText` を使う。
- 空白保持が必要なテキストには `xml:space="preserve"` を付ける。
- 元のラン書式 `w:rPr` を可能な限りコピーする。
- 段落をまたぐ削除・挿入は難しいため、初期版では段落単位で分割して処理する。
- 既存の変更履歴が入力に含まれる場合は、事前にaccept済みの文書を入力として扱う前提にするのが安全。

## テスト設計

### 単体テスト

- トークン化
- 段落抽出
- 段落対応付け
- diff opcodeから変更履歴XMLへの変換
- `xml:space="preserve"` の付与
- revision id の採番

### 結合テスト

以下の入力ペアから `output.docx` を生成し、ZIP/XML上で `w:ins` / `w:del` の存在を検証する。

- 1文中の単語追加
- 1文中の単語削除
- 置換
- 段落追加
- 段落削除
- 日本語文の文字差分
- 空白・改行差分

### 手動確認

Microsoft Wordで `output.docx` を開き、変更履歴として表示されることを確認する。純Python backendではWord以外のビューアとの互換性も確認する。

## 推奨する開発ステップ

1. CLI骨格を作る。
2. Word backendを実装する。
3. Word backendのサンプル比較をWindows環境で確認する。
4. 純Python backendの最小版を作る。
5. 最小サンプルDOCXで `w:ins` / `w:del` をWord表示確認する。
6. 表・ヘッダー等の対応範囲を必要に応じて広げる。

## 未確定事項・確認したい点

実装前に確認したい。

1. 実行環境はWindows + Microsoft Wordを利用できるか。
2. 必須要件は「Wordで変更履歴として見えること」か、「Wordの比較機能と同等の精度」まで必要か。
3. 比較対象は本文のみでよいか。表、ヘッダー/フッター、脚注、コメント、テキストボックスも必要か。
4. 入力DOCXに既存の変更履歴が含まれる可能性があるか。
5. 差分粒度は単語単位、文字単位、どちらを優先するか。日本語文書では文字単位または形態素ベースの検討が必要。

## 参照資料

- Microsoft Learn: Application.CompareDocuments method (Word)  
  https://learn.microsoft.com/en-us/office/vba/api/word.application.comparedocuments
- Microsoft Support: Compare document differences using the legal blackline option  
  https://support.microsoft.com/en-us/office/compare-document-differences-using-the-legal-blackline-option-dbfc7351-4022-43a2-a0c4-54d1898702a0
- Microsoft Learn: How to accept all revisions in a word processing document  
  https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document
- Microsoft Learn: TrackRevisions Class, Open XML SDK  
  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.trackrevisions
- python-docx documentation: Working with Documents  
  https://python-docx.readthedocs.io/en/latest/user/documents.html
- docx-revisions documentation  
  https://balalofernandez.github.io/docx-revisions/
