# XLSX リファレンス

## コマンド詳細

### info / list / show

```bash
scripts/oxen-cli xlsx info book.xlsx                       # シート数、シート名、行数
scripts/oxen-cli xlsx list book.xlsx                       # シート一覧（行数サマリー）
scripts/oxen-cli xlsx show book.xlsx Sheet1                # シート全体
scripts/oxen-cli xlsx show book.xlsx Sheet1 --range A1:E10 # セル範囲指定
scripts/oxen-cli xlsx show book.xlsx Sheet1 -o json        # 構造化データ
```

### extract

```bash
scripts/oxen-cli xlsx extract book.xlsx                         # 最初のシートをCSV出力
scripts/oxen-cli xlsx extract book.xlsx --sheet Sheet2           # シート指定
scripts/oxen-cli xlsx extract book.xlsx --format json            # JSON出力
scripts/oxen-cli xlsx extract book.xlsx --sheet Sheet1 -o json   # 組み合わせ
```

### データ分析系

```bash
scripts/oxen-cli xlsx formulas book.xlsx              # 数式一覧
scripts/oxen-cli xlsx names book.xlsx                 # 名前付き範囲（定義名）
scripts/oxen-cli xlsx strings book.xlsx               # 共有文字列（リッチテキスト書式含む）
```

### 構造・書式系

```bash
scripts/oxen-cli xlsx tables book.xlsx                # テーブル定義（ListObjects）
scripts/oxen-cli xlsx comments book.xlsx              # セルコメント
scripts/oxen-cli xlsx hyperlinks book.xlsx            # ハイパーリンク
scripts/oxen-cli xlsx autofilter book.xlsx            # オートフィルター設定
scripts/oxen-cli xlsx validation book.xlsx            # データ入力規則
scripts/oxen-cli xlsx conditional book.xlsx           # 条件付き書式ルール
scripts/oxen-cli xlsx styles book.xlsx                # フォント、塗り、罫線、表示形式
```

### preview

```bash
scripts/oxen-cli xlsx preview book.xlsx                          # 全シート
scripts/oxen-cli xlsx preview book.xlsx Sheet1                   # 特定シート
scripts/oxen-cli xlsx preview book.xlsx Sheet1 --range A1:E10    # 範囲指定
scripts/oxen-cli xlsx preview book.xlsx --width 120              # 幅指定
```

---

## Build仕様

```jsonc
{
  "template": "./template.xlsx",
  "output": "./output.xlsx"
}
```

テンプレートをコピーして出力。コンテンツの編集は検査コマンドで構造を確認しながら手動で行う。

---

## 典型ワークフロー

```bash
# シート構造を把握
scripts/oxen-cli xlsx info book.xlsx
scripts/oxen-cli xlsx list book.xlsx

# データを見る
scripts/oxen-cli xlsx show book.xlsx Sheet1 --range A1:E10
scripts/oxen-cli xlsx preview book.xlsx Sheet1
scripts/oxen-cli xlsx extract book.xlsx --format json

# 数式・ルールを調べる
scripts/oxen-cli xlsx formulas book.xlsx
scripts/oxen-cli xlsx validation book.xlsx
scripts/oxen-cli xlsx conditional book.xlsx

# スタイルを確認
scripts/oxen-cli xlsx styles book.xlsx
```
