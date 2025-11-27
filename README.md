# term-xlsx

Terminal-based XLSX editor built with Rust.

A TUI (Text User Interface) application for viewing and editing Excel files directly in your terminal.

## Features

- View and edit Excel (.xlsx) files in terminal
- WASD + Arrow keys navigation (FPS-style)
- Cell selection with Shift+Arrow keys
- Copy/Paste support
- Cell color marking (saves to Excel styles)
- Dynamic column width adjustment
- Multi-sheet support
- Excel-compatible shortcuts

## Installation

```bash
cargo build --release
```

Binary will be at `target/release/term-xlsx.exe` (Windows) or `target/release/term-xlsx` (Linux/macOS).

## Usage

```bash
term-xlsx <filename.xlsx>
```

If the file doesn't exist, a new spreadsheet will be created.

## Key Bindings

### Navigation

| Key | Action |
|-----|--------|
| W / Up | Move up |
| A / Left | Move left |
| S / Down | Move down |
| D / Right | Move right |
| Enter | Move down |
| Tab | Move right |
| Shift+Tab | Move left |
| Shift+Enter | Move up |
| Home | Jump to column A |
| End | Jump to last used column |
| Ctrl+Home | Jump to A1 |
| Ctrl+End | Jump to last used cell |
| PageUp | Previous sheet |
| PageDown | Next sheet |
| F4 | Open sheet selector |

### Selection

| Key | Action |
|-----|--------|
| Shift+W/A/S/D | Extend selection |
| Shift+Arrow keys | Extend selection |
| Esc | Clear selection |

### Editing

| Key | Action |
|-----|--------|
| F2 | Enter edit mode |
| Enter (in edit mode) | Save and move down |
| Tab (in edit mode) | Save and move right |
| Esc (in edit mode) | Cancel editing |

### Clipboard

| Key | Action |
|-----|--------|
| C / F5 | Copy selection |
| V / F6 | Paste |

### Column Width

| Key | Action |
|-----|--------|
| E | Expand column width |
| R | Reduce column width |

### Cell Marking (Colors)

| Key | Style |
|-----|-------|
| 1 | Clear (reset) |
| 2 | Yellow background |
| 3 | Red text |
| 4 | Green text |
| 5 | Blue background |
| 6 | Magenta text |

Colors are saved to Excel file styles.

### File Operations

| Key | Action |
|-----|--------|
| Ctrl+S | Save file |
| Ctrl+W | Quit |

## Limits

- Columns: A to IV (256 columns, like classic Excel)
- Rows: 1 to 65536 (like classic Excel)

## Dependencies

- [ratatui](https://github.com/ratatui-org/ratatui) - TUI framework
- [crossterm](https://github.com/crossterm-rs/crossterm) - Terminal manipulation
- [umya-spreadsheet](https://github.com/MathNya/umya-spreadsheet) - XLSX read/write
- [clap](https://github.com/clap-rs/clap) - CLI argument parsing
- [tui-textarea](https://github.com/rhysd/tui-textarea) - Text editing widget

## License

MIT
