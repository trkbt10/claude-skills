# DOCX リファレンス

## コマンド詳細

### info / list / show

```bash
scripts/oxen-cli docx info report.docx           # 段落数、テーブル数、セクション数
scripts/oxen-cli docx list report.docx           # セクション一覧（段落数サマリー）
scripts/oxen-cli docx show report.docx 1         # セクション内容
scripts/oxen-cli docx show report.docx 1 -o json # 構造化データ
```

### extract

```bash
scripts/oxen-cli docx extract report.docx                  # 全テキスト
scripts/oxen-cli docx extract report.docx --sections 1,3-5 # 指定セクション
scripts/oxen-cli docx extract report.docx -o json           # JSON形式
```

### styles

```bash
scripts/oxen-cli docx styles report.docx                    # 全スタイル
scripts/oxen-cli docx styles report.docx --type paragraph   # 段落スタイルのみ
scripts/oxen-cli docx styles report.docx --type character   # 文字スタイルのみ
scripts/oxen-cli docx styles report.docx --type table       # テーブルスタイルのみ
scripts/oxen-cli docx styles report.docx --all              # 非表示スタイル含む
```

### 構造系コマンド

```bash
scripts/oxen-cli docx numbering report.docx        # 番号定義（リスト/箇条書き）
scripts/oxen-cli docx headers-footers report.docx   # ヘッダー/フッター
scripts/oxen-cli docx tables report.docx            # テーブル構造
scripts/oxen-cli docx comments report.docx          # コメント
scripts/oxen-cli docx images report.docx            # 埋め込み画像情報
scripts/oxen-cli docx toc report.docx               # 目次（アウトラインレベルベース）
```

### preview

```bash
scripts/oxen-cli docx preview report.docx             # 全セクション
scripts/oxen-cli docx preview report.docx 1            # 特定セクション
scripts/oxen-cli docx preview report.docx --width 120  # 幅指定
```

---

## Build仕様

```jsonc
{
  "template": "./template.docx",
  "output": "./output.docx"
}
```

テンプレートをコピーして出力。コンテンツの編集は検査コマンドで構造を確認しながら手動で行う。

---

## 典型ワークフロー

```bash
# 文書構造を把握
scripts/oxen-cli docx info report.docx
scripts/oxen-cli docx list report.docx
scripts/oxen-cli docx toc report.docx

# スタイルと書式を調べる
scripts/oxen-cli docx styles report.docx --type paragraph
scripts/oxen-cli docx numbering report.docx

# コンテンツを見る
scripts/oxen-cli docx show report.docx 1 -o json
scripts/oxen-cli docx tables report.docx
scripts/oxen-cli docx preview report.docx
```
