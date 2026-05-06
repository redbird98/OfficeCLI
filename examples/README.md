# OfficeCLI Examples

Comprehensive examples demonstrating OfficeCLI capabilities for Word, Excel, and PowerPoint automation.

## 📂 Directory Structure

```
examples/
├── README.md                          # This file
├── word/                              # 📄 Word examples
│   ├── formulas.sh / formulas.docx
│   ├── tables.sh / tables.docx
│   ├── textbox.sh
│   └── numbering-showcase.sh / numbering-showcase.docx
├── excel/                             # 📊 Excel examples
│   ├── charts.sh / charts.xlsx        # Master chart showcase
│   ├── charts-demo.sh / charts-demo.xlsx
│   ├── charts-<type>.py / .xlsx       # Per-type chart scripts
│   │   (basic, advanced, extended, area, bar, boxwhisker,
│   │    bubble, column, combo, histogram, line, pie, radar,
│   │    scatter, stock, waterfall)
│   └── pivot-tables.py / pivot-tables.xlsx
└── ppt/                               # 🎨 PowerPoint examples
    ├── presentation.{md,sh,pptx}
    ├── animations.{md,sh,pptx}
    ├── video.{md,py,pptx}
    └── 3d-model.{md,sh,pptx}
```

Each example follows the same trio: `<name>.md` (walkthrough), `<name>.sh`/`.py` (build script), `<name>.<ext>` (pre-generated output).

---

## 🚀 Quick Start

### By Document Type

**Word (.docx):**
```bash
cd word
bash formulas.sh             # LaTeX math formulas
bash tables.sh               # Styled tables
bash textbox.sh              # Formatted text boxes
bash numbering-showcase.sh   # List/numbering styles
```

**Excel (.xlsx):**
```bash
cd excel
bash charts.sh               # Master chart showcase
bash charts-demo.sh          # 14+ chart types
python charts-line.py        # Single-type example (any charts-<type>.py)
python pivot-tables.py       # Pivot tables
```

**PowerPoint (.pptx):**
```bash
cd ppt
bash presentation.sh         # Morph transitions / full deck
bash animations.sh           # Animation effects
python video.py              # Video embedding
bash 3d-model.sh             # 3D model embedding
```

---

## 📚 Documentation by Type

### 📄 [Word Examples →](word/)
- Mathematical formulas (LaTeX)
- Complex tables
- Text boxes and styling
- Numbering / list showcases

### 📊 [Excel Examples →](excel/)
- Master and per-type chart scripts (line, bar, pie, scatter, stock, waterfall, …)
- Pivot tables
- Number formatting and styling

### 🎨 [PowerPoint Examples →](ppt/)
- Slide / shape construction
- Morph transitions and animations
- Video and 3D model embedding

---

## 🔧 Common Patterns

### Create and Populate

```bash
#!/bin/bash
set -e

FILE="document.docx"
officecli create "$FILE"
officecli add "$FILE" /body --type paragraph --prop text="Hello World"
officecli validate "$FILE"
```

### Batch Operations

```bash
cat << 'EOF' > commands.json
[
  {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Para 1"}},
  {"command":"set","path":"/body/p[1]","props":{"bold":"true","size":"24"}}
]
EOF
officecli batch document.docx < commands.json
```

### Resident Mode (3+ operations)

```bash
officecli open document.docx
officecli add document.docx /body --type paragraph --prop text="Fast operation"
officecli set document.docx /body/p[1] --prop bold=true
officecli close document.docx
```

### Query and Modify

```bash
# Find all Heading1 paragraphs
officecli query report.docx "paragraph[style=Heading1]" --json

# Change their color
officecli set report.docx /body/p[1] --prop color=FF0000
```

---

## 📊 Quick Reference

### Document Types

| Format | Extension | Create | View | Modify |
|--------|-----------|--------|------|--------|
| Word | .docx | ✓ | ✓ | ✓ |
| Excel | .xlsx | ✓ | ✓ | ✓ |
| PowerPoint | .pptx | ✓ | ✓ | ✓ |

### Common Commands

| Command | Purpose | Example |
|---------|---------|---------|
| `create` | Create blank document | `officecli create file.docx` |
| `view` | View content | `officecli view file.docx text` |
| `get` | Get element | `officecli get file.docx /body/p[1]` |
| `set` | Modify element | `officecli set file.docx /body/p[1] --prop bold=true` |
| `add` | Add element | `officecli add file.docx /body --type paragraph` |
| `remove` | Remove element | `officecli remove file.docx /body/p[5]` |
| `query` | CSS-like query | `officecli query file.docx "paragraph[style=Normal]"` |
| `batch` | Multiple operations | `officecli batch file.docx < commands.json` |
| `validate` | Check schema | `officecli validate file.docx` |

### View Modes

| Mode | Description | Usage |
|------|-------------|-------|
| `text` | Plain text | `officecli view file.docx text` |
| `annotated` | Text with formatting | `officecli view file.docx annotated` |
| `outline` | Structure | `officecli view file.docx outline` |
| `stats` | Statistics | `officecli view file.docx stats` |
| `issues` | Problems | `officecli view file.docx issues` |
| `html` | HTML preview | `officecli view file.docx html` |
| `svg` | SVG preview | `officecli view file.docx svg` |
| `forms` | Form fields | `officecli view file.docx forms` |

---

## 💡 Tips

1. **Explore before modifying:**
   ```bash
   officecli view document.docx outline
   officecli get document.docx /body --depth 2
   ```

2. **Use `--json` for automation:**
   ```bash
   officecli query data.xlsx "cell[formula~=SUM]" --json | jq
   ```

3. **Check help for properties** (schema reference is under the `help` verb):
   ```bash
   officecli help docx set paragraph
   officecli help xlsx set cell
   officecli help pptx set shape
   ```

4. **Validate after changes:**
   ```bash
   officecli validate document.docx
   ```

5. **Use resident mode for performance** (3+ operations on same file):
   ```bash
   officecli open file.pptx
   # ... multiple commands ...
   officecli close file.pptx
   ```

---

## 🤝 Contributing Examples

1. **Create script** with clear comments
2. **Test and verify** output
3. **Add to appropriate directory** (word/excel/ppt)
4. **Update directory README**
5. **Submit PR**

**Example format:**
```bash
#!/bin/bash
# Brief description of what this demonstrates
# Key techniques: list them here

set -e

FILE="output.docx"
officecli create "$FILE"
# ... your commands ...
officecli validate "$FILE"
echo "Created: $FILE"
```

---

## 📖 More Resources

- **[SKILL.md](../SKILL.md)** - Complete command reference for AI agents
- **[README.md](../README.md)** - Project overview and installation

---

## 🆘 Getting Help

**Top-level help:**
```bash
officecli --help                       # CLI usage
officecli help                         # Schema reference entry point
officecli help docx                    # All docx elements
officecli help docx set                # Elements that support `set` for docx
officecli help docx set paragraph      # Settable properties on paragraph
officecli help docx paragraph --json   # Raw schema JSON
officecli help all                     # Flat dump of every (format, element, property)
```

Format aliases: `word→docx`, `excel→xlsx`, `ppt`/`powerpoint→pptx`.
Verbs: `add`, `set`, `get`, `query`, `remove`.

---

**Happy automating! 🚀**

For questions or issues, visit [GitHub Issues](https://github.com/iOfficeAI/OfficeCLI/issues).
