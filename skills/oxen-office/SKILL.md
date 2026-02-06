---
name: oxen-office
description: |
  Inspect, build, and verify Office documents (PPTX, DOCX, XLSX) using the oxen CLI.
  Use this skill when: (1) Inspecting existing Office files, (2) Building PPTX/DOCX/XLSX from JSON spec,
  (3) Previewing slide/document/sheet content in terminal, (4) Extracting text or data,
  (5) Verifying build results, (6) User asks about ".pptx", ".docx", ".xlsx" inspection or generation.
---

# Oxen Office

`scripts/oxen-cli` CLIでOfficeドキュメントを操作する。Bunランタイムで動作。

## ファイルがある → 何ができる？

### PPTXファイルがある

| やりたいこと | コマンド |
|-------------|---------|
| スライド数・サイズ確認 | `scripts/oxen-cli pptx info <file>` |
| スライド一覧 | `scripts/oxen-cli pptx list <file>` |
| スライド内容を見る | `scripts/oxen-cli pptx show <file> <n>` |
| テキスト抽出 | `scripts/oxen-cli pptx extract <file>` |
| テーマ色・フォント確認 | `scripts/oxen-cli pptx theme <file>` |
| ASCIIプレビュー | `scripts/oxen-cli pptx preview <file> [n] --border` |
| テンプレートから構築 | `scripts/oxen-cli pptx build <spec.json>` |
| ビルド結果を検証 | `scripts/oxen-cli pptx verify <path>` |

→ 詳細: `references/pptx.md`

### DOCXファイルがある

| やりたいこと | コマンド |
|-------------|---------|
| 文書メタデータ | `scripts/oxen-cli docx info <file>` |
| セクション一覧 | `scripts/oxen-cli docx list <file>` |
| セクション内容 | `scripts/oxen-cli docx show <file> <n>` |
| テキスト抽出 | `scripts/oxen-cli docx extract <file>` |
| スタイル確認 | `scripts/oxen-cli docx styles <file>` |
| 目次 | `scripts/oxen-cli docx toc <file>` |
| テーブル / 画像 / コメント | `scripts/oxen-cli docx tables\|images\|comments <file>` |
| ASCIIプレビュー | `scripts/oxen-cli docx preview <file> [n]` |

→ 詳細: `references/docx.md`

### XLSXファイルがある

| やりたいこと | コマンド |
|-------------|---------|
| ブック情報 | `scripts/oxen-cli xlsx info <file>` |
| シート一覧 | `scripts/oxen-cli xlsx list <file>` |
| シート内容 | `scripts/oxen-cli xlsx show <file> <sheet> --range A1:E10` |
| データ抽出（CSV/JSON） | `scripts/oxen-cli xlsx extract <file> --format json` |
| 数式 / 名前付き範囲 | `scripts/oxen-cli xlsx formulas\|names <file>` |
| 条件付き書式 / 入力規則 | `scripts/oxen-cli xlsx conditional\|validation <file>` |
| ASCIIプレビュー | `scripts/oxen-cli xlsx preview <file> [sheet]` |

→ 詳細: `references/xlsx.md`

## 共通パターン

```bash
# JSON出力（パース可能な構造化データ）
scripts/oxen-cli pptx show deck.pptx 1 -o json

# 特定範囲の抽出
scripts/oxen-cli pptx extract deck.pptx --slides 1,3-5
scripts/oxen-cli docx extract doc.docx --sections 1,3-5
```

## ビルドしたい → どうする？

PPTXのみJSON specベースの本格ビルドに対応（DOCX/XLSXはテンプレートコピー）。

```bash
# 1. テンプレートを調べる
scripts/oxen-cli pptx show template.pptx 1 -o json
scripts/oxen-cli pptx theme template.pptx

# 2. spec.jsonを作成（→ references/pptx.md のBuild仕様を参照）

# 3. ビルド → 確認
scripts/oxen-cli pptx build spec.json
scripts/oxen-cli pptx preview output.pptx --border
```
