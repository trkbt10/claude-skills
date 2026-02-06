# Claude Skills

A collection of Office document manipulation plugins for Claude Code.

## Install

```shell
/plugin marketplace add trkbt10/claude-skills
```

## Plugins

| Plugin | Description |
|--------|-------------|
| [slide-studio](./skills/slide-studio/) | PowerPoint (PPTX) creation and editing via direct XML manipulation |
| [oxen-office](./skills/oxen-office/) | Office document (PPTX/DOCX/XLSX) inspection, building, and verification |

### slide-studio

```shell
/plugin install slide-studio@claude-skills
```

PPTX を XML 直接編集で作成・編集。POTX テンプレート適用、全 11 種の標準スライドレイアウト対応。

### oxen-office

```shell
/plugin install oxen-office@claude-skills
```

oxen CLI で Office ドキュメントを操作。ASCII プレビュー、JSON スペックによるビルド、テキスト抽出に対応。
