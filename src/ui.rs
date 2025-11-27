use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    widgets::{Block, Borders, Cell, Paragraph, Row, Table},
    Frame,
};
use crate::app::{App, CellMark, Mode};

pub fn draw(f: &mut Frame, app: &mut App) {
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(3), // Header/Tabs
            Constraint::Min(0),    // Grid
            Constraint::Length(3), // Status/Input
        ])
        .split(f.area());

    draw_header(f, app, chunks[0]);
    draw_grid(f, app, chunks[1]);
    draw_status(f, app, chunks[2]);
}

fn draw_header(f: &mut Frame, app: &App, area: Rect) {
    let sheet_names: Vec<String> = app.spreadsheet.get_sheet_collection()
        .iter()
        .map(|s| s.get_name().to_string())
        .collect();

    let tabs = sheet_names.join(" | ");
    let title = format!("File: {:?} | Sheets: {}", app.path, tabs);
    
    let block = Block::default().borders(Borders::ALL).title(title);
    f.render_widget(block, area);
}

fn draw_grid(f: &mut Frame, app: &mut App, area: Rect) {
    let block = Block::default().borders(Borders::ALL).title("Grid");
    let inner = block.inner(area);

    // Calculate how many rows/cols we can fit
    let row_num_width: u16 = 6; // Width for row numbers
    let available_width = inner.width.saturating_sub(row_num_width);
    let available_height = inner.height.saturating_sub(1); // -1 for header row

    // Calculate visible columns based on their widths
    let start_col = app.scroll.1 + 1;
    let mut num_cols = 0u32;
    let mut used_width: u16 = 0;
    loop {
        let col_idx = start_col + num_cols;
        let col_width = app.get_column_width(col_idx) + 1; // +1 for spacing
        if used_width + col_width > available_width {
            break;
        }
        used_width += col_width;
        num_cols += 1;
        if num_cols >= 50 {
            break; // Safety limit
        }
    }
    num_cols = num_cols.max(1);

    let num_rows = (available_height as u32).max(1);

    // Update viewport size for scroll calculations
    app.viewport_size = (num_rows as u16, num_cols as u16);

    let start_row = app.scroll.0 + 1;

    let mut rows = Vec::new();

    // Header row (Column letters)
    let mut header_cells = vec![Cell::from("     ")];
    for c in 0..num_cols {
        let col_idx = start_col + c;
        let col_letter = number_to_column(col_idx);
        header_cells.push(Cell::from(col_letter).style(Style::default().add_modifier(Modifier::BOLD)));
    }
    rows.push(Row::new(header_cells));

    for r in 0..num_rows {
        let row_idx = start_row + r;
        let mut row_cells = vec![Cell::from(format!("{:>5}", row_idx)).style(Style::default().add_modifier(Modifier::BOLD))];

        for c in 0..num_cols {
            let col_idx = start_col + c;
            let value = app.get_cell_display(col_idx, row_idx);

            let is_cursor = row_idx == app.cursor.0 && col_idx == app.cursor.1;
            let is_selected = app.selection.contains(row_idx, col_idx);
            let mark = app.get_cell_mark(row_idx, col_idx);

            // Build style: cursor > selection > mark > default
            let style = if is_cursor {
                Style::default().bg(Color::Blue).fg(Color::White)
            } else if is_selected {
                Style::default().bg(Color::DarkGray).fg(Color::White)
            } else {
                match mark {
                    CellMark::None => Style::default(),
                    CellMark::YellowBg => Style::default().bg(Color::Yellow).fg(Color::Black),
                    CellMark::RedText => Style::default().fg(Color::Red),
                    CellMark::GreenText => Style::default().fg(Color::Green),
                    CellMark::BlueBg => Style::default().bg(Color::LightBlue).fg(Color::Black),
                    CellMark::MagentaText => Style::default().fg(Color::Magenta),
                }
            };

            row_cells.push(Cell::from(value).style(style));
        }
        rows.push(Row::new(row_cells));
    }

    // Build dynamic column widths
    let mut widths = vec![Constraint::Length(row_num_width)];
    for c in 0..num_cols {
        let col_idx = start_col + c;
        let width = app.get_column_width(col_idx);
        widths.push(Constraint::Length(width));
    }

    let table = Table::new(rows, widths)
        .block(block)
        .column_spacing(1);

    f.render_widget(table, area);
}

fn draw_status(f: &mut Frame, app: &mut App, area: Rect) {
    match app.mode {
        Mode::View => {
            // Show status message if present, otherwise show help
            let text = if let Some(ref msg) = app.status_message {
                msg.clone()
            } else {
                // Build cell reference
                let cell_ref = format!("{}{}",
                    number_to_column(app.cursor.1),
                    app.cursor.0
                );

                // Show selection info if multi-cell
                let sel_info = if !app.selection.is_single() {
                    let (r1, c1, r2, c2) = app.selection.bounds();
                    format!(" [{}{}:{}{}]",
                        number_to_column(c1), r1,
                        number_to_column(c2), r2
                    )
                } else {
                    String::new()
                };

                format!("{}{} | ^W:Quit ^S:Save | WASD:Move | C/V:Copy/Paste | F2:Edit | 1-6:Mark",
                    cell_ref, sel_info)
            };

            let p = Paragraph::new(text).block(Block::default().borders(Borders::ALL));
            f.render_widget(p, area);
        }
        Mode::Edit => {
            app.textarea.set_block(Block::default().borders(Borders::ALL).title("Editing (Enter:Save+Down, Tab:Save+Right, Esc:Cancel)"));
            f.render_widget(&app.textarea, area);
        }
    }
}

fn number_to_column(n: u32) -> String {
    let mut n = n;
    let mut result = String::new();
    while n > 0 {
        n -= 1;
        let remainder = n % 26;
        let char_code = (remainder as u8) + b'A';
        result.insert(0, char_code as char);
        n /= 26;
    }
    result
}
