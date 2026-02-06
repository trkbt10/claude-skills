# PPTX リファレンス

## コマンド詳細

### info / list / show

```bash
scripts/oxen-cli pptx info deck.pptx           # スライド数、サイズ(px/EMU)
scripts/oxen-cli pptx list deck.pptx           # 各スライドのタイトルとシェイプ数
scripts/oxen-cli pptx show deck.pptx 1         # シェイプ一覧（ID, 名前, プレースホルダー, テキスト）
scripts/oxen-cli pptx show deck.pptx 1 -o json # 構造化データ（bounds, paragraphs, runs含む）
```

### extract

```bash
scripts/oxen-cli pptx extract deck.pptx              # 全スライドのテキスト
scripts/oxen-cli pptx extract deck.pptx --slides 1,3-5  # 指定スライドのみ
scripts/oxen-cli pptx extract deck.pptx -o json       # JSON形式
```

### theme

```bash
scripts/oxen-cli pptx theme deck.pptx
# → Font Scheme (Major/Minor), Color Scheme (dk1,lt1,accent1-6,...), Format Scheme
```

### preview

```bash
scripts/oxen-cli pptx preview deck.pptx              # 全スライド
scripts/oxen-cli pptx preview deck.pptx 1 --border   # 特定スライド、枠線付き
scripts/oxen-cli pptx preview deck.pptx --width 120   # 幅指定
```

### verify

```bash
scripts/oxen-cli pptx verify test-case.json   # 単一テスト
scripts/oxen-cli pptx verify tests/           # ディレクトリ内全テスト
scripts/oxen-cli pptx verify tests/ --tag smoke  # タグでフィルタ
```

---

## Build仕様

### JSON Spec トップレベル

```jsonc
{
  "template": "./template.pptx",   // 必須: ベースPPTX
  "output": "./output.pptx",       // 必須: 出力先
  "theme": { ... },                // テーマ編集
  "addSlides": [ ... ],            // スライド追加
  "duplicateSlides": [ ... ],      // スライド複製
  "reorderSlides": [ ... ],        // 並べ替え
  "removeSlides": [ ... ],         // 削除
  "slides": [ ... ]                // スライド別編集
}
```

### スライド編集（slides配列の要素）

```jsonc
{
  "slideNumber": 1,            // 対象スライド番号（1始まり）
  "background": { ... },
  "addShapes": [ ... ],
  "addImages": [ ... ],
  "addConnectors": [ ... ],
  "addGroups": [ ... ],
  "addTables": [ ... ],
  "addCharts": [ ... ],
  "updateCharts": [ ... ],
  "updateTables": [ ... ],
  "addAnimations": [ ... ],
  "addComments": [ ... ],
  "speakerNotes": "ノート文字列",
  "transition": { ... }
}
```

---

### addShapes

```jsonc
{
  "type": "rectangle",          // シェイプタイプ（後述）
  "x": 100, "y": 100,           // 位置（px）
  "width": 400, "height": 200,  // サイズ（px）
  "rotation": 0,
  "flipH": false, "flipV": false,
  "fill": "#4472C4",            // 塗り（後述）
  "line": { "color": "#000", "width": 1 },
  "text": "テキスト",            // 単純テキスト or リッチテキスト（後述）
  "textBody": { "anchor": "ctr", "wrapping": "square" },
  "placeholder": { "type": "title", "idx": 0 },
  "effects": { ... },
  "shape3d": { ... }
}
```

#### シェイプタイプ

| カテゴリ | タイプ |
|---------|--------|
| 基本図形 | rectangle, ellipse, roundRect, triangle, diamond, pentagon, hexagon, octagon |
| 矢印 | rightArrow, leftArrow, upArrow, downArrow, chevron, bentArrow, uturnArrow, curvedRightArrow |
| 星 | star4, star5, star6, star8, star10, star12, star16, star24, star32 |
| 吹き出し | wedgeRectCallout, wedgeEllipseCallout, cloudCallout, borderCallout1-3 |
| フローチャート | flowChartProcess, flowChartDecision, flowChartTerminator, flowChartDocument, flowChartData, flowChartConnector |
| 装飾 | ribbon, ribbon2, wave, doubleWave, heart, cloud, frame, teardrop, funnel |
| 角丸バリエーション | round1Rect, round2SameRect, snip1Rect, snip2SameRect, snipRoundRect |

#### 塗り（fill）

```jsonc
// HEX直接指定
"fill": "#4472C4"

// ソリッド
"fill": { "type": "solid", "color": "#FF0000" }

// テーマ色
"fill": { "type": "theme", "theme": "accent1", "lumMod": 75, "lumOff": 25, "tint": 50, "shade": 50 }

// グラデーション
"fill": {
  "type": "gradient", "gradientType": "linear", "angle": 90,
  "stops": [
    { "position": 0, "color": "#FF0000" },
    { "position": 100, "color": "#0000FF" }
  ]
}

// パターン
"fill": { "type": "pattern", "preset": "dkDnDiag", "fgColor": "#000", "bgColor": "#FFF" }
```

テーマ色名: `dk1`, `dk2`, `lt1`, `lt2`, `accent1`〜`accent6`, `hlink`, `folHlink`

#### テキスト

```jsonc
// 単純テキスト
"text": "Hello"

// リッチテキスト（段落の配列）
"text": [
  {
    "runs": [
      { "text": "太字", "bold": true, "fontSize": 24, "color": "#FF0000" },
      { "text": " 通常", "fontFamily": "Arial" }
    ],
    "alignment": "center",
    "bullet": { "type": "char", "char": "•" },
    "lineSpacing": { "type": "percent", "value": 150 },
    "spaceBefore": 6, "spaceAfter": 6,
    "level": 0
  }
]
```

**Runプロパティ**: `text`, `bold`, `italic`, `underline`(single/double/heavy/dotted/dashed/wavy), `strikethrough`(single/double), `fontSize`, `fontFamily`, `color`(HEX), `caps`(all/small), `letterSpacing`, `verticalPosition`(normal/superscript/subscript), `hyperlink`({url, tooltip})

**段落プロパティ**: `alignment`(left/center/right/justify), `level`, `bullet`, `lineSpacing`, `spaceBefore`, `spaceAfter`, `indent`, `marginLeft`

**textBody**: `anchor`(t/ctr/b), `verticalType`, `wrapping`(square/none), `anchorCenter`, `insetLeft/Top/Right/Bottom`

---

### addImages

```jsonc
{
  "path": "./images/logo.png",   // 画像パス（specからの相対）
  "x": 100, "y": 100,
  "width": 200, "height": 150,
  "rotation": 0,
  "effects": { ... }             // blipエフェクト（省略可）
}
```

対応形式: PNG, JPG/JPEG, GIF, SVG

### addConnectors

```jsonc
{
  "preset": "straightConnector1",
  "x": 100, "y": 100, "width": 300, "height": 0,
  "startShapeId": "2", "startSiteIndex": 1,
  "endShapeId": "3", "endSiteIndex": 3,
  "line": { "color": "#000", "width": 2 }
}
```

---

## 典型ワークフロー

```bash
# 調査
scripts/oxen-cli pptx info deck.pptx
scripts/oxen-cli pptx list deck.pptx
scripts/oxen-cli pptx show deck.pptx 1 -o json
scripts/oxen-cli pptx theme deck.pptx
scripts/oxen-cli pptx preview deck.pptx --border

# ビルド
scripts/oxen-cli pptx build spec.json
scripts/oxen-cli pptx preview output.pptx --border
scripts/oxen-cli pptx extract output.pptx
```
