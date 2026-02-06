---
name: slide-studio
description: |
  Create and modify PowerPoint presentations (PPTX) through direct XML manipulation.
  Use this skill when: (1) Creating presentations from scratch, (2) Adding/editing/deleting slides,
  (3) Inserting images or custom shapes, (4) Applying POTX templates for consistent branding,
  (5) Editing existing PPTX files, (6) Working with slide layouts and placeholders,
  (7) User asks to make a "presentation", "slides", "deck", or mentions ".pptx".
  Supports all 11 standard Office slide layouts.
---

# Slide Studio

PowerPointプレゼンテーションをXML直接編集で作成。

## Step 1: 目的を確認

**必ず最初にユーザーに確認する:**

> 何をしたいですか？
>
> 1. **新規作成**: 新しいプレゼンテーションを作成
> 2. **テーマ適用**: 既存のPPTXに新しいテーマ（POTX）を適用してデザインを変更
> 3. **編集**: 既存のPPTXの内容を編集（テーマ変更なし）

### テンプレートについて

> プレゼンテーションのテンプレート（POTX）はお持ちですか？
>
> **持っている場合**: テンプレートのテーマ、色、レイアウトを適用します
> **持っていない場合**: クリエイティブなテンプレートを探すお手伝いをしましょうか？

### クリエイティブなテンプレートが欲しい場合

以下のサイトで無料のPOTX/PPTXテンプレートを入手可能:

| サイト | 特徴 |
|--------|------|
| [Slidesgo](https://slidesgo.com/) | モダン、多彩なスタイル、Google Slides/PPTX両対応 |
| [SlidesCarnival](https://www.slidescarnival.com/) | ビジネス向け、シンプル、完全無料 |
| [ALLPPT](https://www.free-powerpoint-templates-design.com/) | 大量のテンプレート、業界別カテゴリ |
| [Canva](https://www.canva.com/templates/presentations/) | デザイン性高い、PPTXエクスポート可能 |
| [Microsoft 365](https://create.microsoft.com/en-us/templates/presentations) | 公式、Office最適化 |

**ユーザーへの提案例:**
- 「どんな雰囲気のプレゼンを作りたいですか？（ビジネス/クリエイティブ/ミニマル/カラフル）」
- 「Slidesgoで『[キーワード] presentation template』で検索すると良いテンプレートが見つかります」
- 「テンプレートをダウンロードしたら、POTXまたはPPTXファイルを共有してください」

---

## Step 2: ワークフロー選択

| シナリオ | ワークフロー |
|----------|--------------|
| テンプレート**あり**（新規作成） | [テンプレート適用](#テンプレート適用ワークフロー) |
| テンプレート**なし**（新規作成） | [ベースから作成](#ベースから作成ワークフロー) |
| 既存PPTXにテーマ適用 | [既存ファイルへのテーマ適用](#既存ファイルへのテーマ適用ワークフロー) |
| 既存ファイル編集（テーマ変更なし） | [編集ワークフロー](#編集ワークフロー) |

---

## テンプレート適用ワークフロー

```bash
# 1. ベース展開
node scripts/unpack.js assets/base.pptx ./work

# 2. テンプレート適用
node scripts/apply-potx.js ./work user-template.potx

# 3. レイアウト確認
node scripts/list-layouts.js ./work

# 4. スライド構築
node scripts/add-slide.js ./work --layout 2
node scripts/edit-text.js ./work --slide 2 --placeholder title --text "タイトル"

# 5. パック
node scripts/pack.js ./work output.pptx
```

---

## ベースから作成ワークフロー

```bash
# 1. ベース展開
node scripts/unpack.js assets/base.pptx ./work

# 2. スライド構築
node scripts/add-slide.js ./work --layout 1  # タイトルスライド
node scripts/add-slide.js ./work --layout 2  # コンテンツスライド

# 3. テキスト編集
node scripts/edit-text.js ./work --slide 1 --placeholder ctrTitle --text "タイトル"

# 4. パック
node scripts/pack.js ./work output.pptx
```

---

## 既存ファイルへのテーマ適用ワークフロー

既存のPPTXファイルに新しいテーマ（POTX）を適用して、デザインを一新する。
スライドの内容は保持しつつ、色、フォント、背景、レイアウトスタイルが変更される。

```bash
# 1. 既存PPTXを展開
node scripts/unpack.js existing-presentation.pptx ./work

# 2. 現在の構造を確認
node scripts/list-slides.js ./work
node scripts/list-layouts.js ./work

# 3. テーマ（POTX）を適用
node scripts/apply-potx.js ./work new-theme.potx

# 4. 新しいレイアウトを確認
node scripts/list-layouts.js ./work

# 5. （オプション）内容の微調整
# テーマ変更後、レイアウトの調整が必要な場合
node scripts/edit-text.js ./work --slide 1 --placeholder title --text "調整後タイトル"

# 6. パック
node scripts/pack.js ./work themed-output.pptx
```

### 注意点

- **レイアウトマッピング**: テンプレートに同じ番号のレイアウトがあれば自動的にマッピングされる。存在しない場合はレイアウト1（タイトルスライド）にフォールバック
- **内容の保持**: テキストや画像などのコンテンツは保持されるが、位置やスタイルはテーマに依存
- **確認推奨**: 適用後は必ずPowerPoint/Google Slidesで開いて表示を確認する

### PPTXをテンプレートとして使用

POTXファイルがなくても、PPTXファイルからテーマを抽出できる:

```bash
# PPTXファイルでも同じコマンドで適用可能
node scripts/apply-potx.js ./work design-source.pptx
```

---

## 編集ワークフロー

```bash
# 1. 展開
node scripts/unpack.js existing.pptx ./work

# 2. 構造確認
node scripts/list-slides.js ./work

# 3. 編集
node scripts/edit-text.js ./work --slide 1 --placeholder title --text "新タイトル"

# 4. パック
node scripts/pack.js ./work updated.pptx
```

---

## スクリプト一覧

| スクリプト | 用途 | 使用例 |
|------------|------|--------|
| `unpack.js` | PPTX展開 | `node scripts/unpack.js input.pptx ./work` |
| `pack.js` | PPTX作成 | `node scripts/pack.js ./work output.pptx` |
| `apply-potx.js` | テンプレート適用 | `node scripts/apply-potx.js ./work template.potx` |
| `add-slide.js` | スライド追加 | `node scripts/add-slide.js ./work --layout 2` |
| `delete-slide.js` | スライド削除 | `node scripts/delete-slide.js ./work --slide 3` |
| `clone-slide.js` | スライド複製 | `node scripts/clone-slide.js ./work --source 1` |
| `edit-text.js` | テキスト編集 | `node scripts/edit-text.js ./work --slide 1 --placeholder title --text "..."` |
| `add-image.js` | 画像追加 | `node scripts/add-image.js ./work --slide 1 --image logo.png` |
| `list-layouts.js` | レイアウト一覧 | `node scripts/list-layouts.js ./work` |
| `list-slides.js` | スライド一覧 | `node scripts/list-slides.js ./work` |
| `visualize-slide.js` | ASCII構造表示 | `node scripts/visualize-slide.js ./work --slide 1` |
| `add-shape.js` | シェイプ自由配置 | `node scripts/add-shape.js ./work --slide 1 --type textbox ...` |

---

## レイアウト

| # | 名前 | プレースホルダー |
|---|------|------------------|
| 1 | Title Slide | ctrTitle, subTitle |
| 2 | Title and Content | title, body |
| 3 | Section Header | title, body |
| 4 | Two Content | title, body×2 |
| 5 | Comparison | title, body×4 |
| 6 | Title Only | title |
| 7 | Blank | (なし) |
| 8 | Content with Caption | title, body×2 |
| 9 | Picture with Caption | title, body, pic |
| 10 | Title and Vertical Text | title, body |
| 11 | Vertical Title and Text | title, body |

---

## カスタムレイアウト（自由配置）

標準レイアウトでは表現できない場合、`add-shape.js`でシェイプを自由配置。

### 構造確認（ASCII可視化）

```bash
# スライドのシェイプ配置をASCIIアートで表示
node scripts/visualize-slide.js ./work --slide 1
```

出力例:
```
+----------------------------------------------------------------------+
|                                                                      |
|          +--------------------------+                                |
|          |            [1]           |                                |
|          +--------------------------+                                |
|                                                                      |
+----------------------------------------------------------------------+

Shapes:
[1] TextBox 4
    Position: (2.00", 3.00")  Size: 5.00" × 1.00"
    Text: "Custom Text..."
```

### シェイプ追加

```bash
# テキストボックス追加（テーマ色対応）
node scripts/add-shape.js ./work --slide 1 --type textbox \
    --text "見出し" --x 2 --y 1 --width 5 --height 1 \
    --color accent1 --font-size 24 --bold

# 矩形追加
node scripts/add-shape.js ./work --slide 1 --type rect \
    --x 1 --y 5 --width 10 --height 0.5 --fill accent2
```

### テーマ色

テンプレート適用後、以下のテーマ色が使用可能:

| 色名 | 用途 |
|------|------|
| `dk1`, `dk2` | ダーク（テキスト向け） |
| `lt1`, `lt2` | ライト（背景向け） |
| `accent1`〜`accent6` | アクセントカラー |
| `hlink`, `folHlink` | ハイパーリンク |

RGB直接指定も可能: `--color FF0000`

### add-shape.js オプション

| オプション | 説明 |
|------------|------|
| `--type` | `textbox` または `rect` |
| `--x`, `--y` | 位置（インチ） |
| `--width`, `--height` | サイズ（インチ） |
| `--text` | テキスト内容 |
| `--color` | テキスト色（テーマ色 or RGB） |
| `--fill` | 塗りつぶし色 |
| `--font-size` | フォントサイズ（pt） |
| `--bold`, `--italic` | 太字/斜体 |
| `--align` | `left`, `center`, `right` |

### 手動編集が必要な場合

複雑なシェイプや高度な書式は `./work/ppt/slides/slideN.xml` を直接編集。
詳細は `references/slide-patterns.md` を参照。

---

## 確認方法

### テキスト抽出
```bash
# markitdownでテキスト内容を確認（インストール済みの場合）
python -m markitdown output.pptx
```

### 構造確認
```bash
# スライド一覧
node scripts/list-slides.js ./work

# XMLを直接読む
# slideN.xmlの<p:sp>要素でシェイプ構造を確認
```

### ユーザー確認
生成後、ユーザーにPowerPoint/Google Slidesで開いて確認してもらう。
