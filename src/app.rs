use anyhow::Result;
use crossterm::event::{KeyCode, KeyEvent, KeyEventKind, KeyModifiers};
use std::collections::HashMap;
use std::path::PathBuf;
use tui_textarea::TextArea;
use umya_spreadsheet::{Color, NumberingFormat, PatternValues, Spreadsheet, helper::number_format::to_formatted_string};

pub const DEFAULT_COLUMN_WIDTH: u16 = 10;
pub const COLUMN_WIDTH_STEP: u16 = 2;
pub const MAX_COLUMN_WIDTH: u16 = 50;

// Old Excel limits (XLS format)
pub const MAX_COLUMNS: u32 = 256;   // A to IV
pub const MAX_ROWS: u32 = 65536;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Mode {
    View,
    Edit,
    SheetSelect,
}

/// Selection range: (start_row, start_col, end_row, end_col) all 1-based
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Selection {
    pub start: (u32, u32),  // (row, col)
    pub end: (u32, u32),    // (row, col)
}

impl Selection {
    pub fn single(row: u32, col: u32) -> Self {
        Self {
            start: (row, col),
            end: (row, col),
        }
    }

    /// Get normalized bounds (top-left to bottom-right)
    pub fn bounds(&self) -> (u32, u32, u32, u32) {
        let min_row = self.start.0.min(self.end.0);
        let max_row = self.start.0.max(self.end.0);
        let min_col = self.start.1.min(self.end.1);
        let max_col = self.start.1.max(self.end.1);
        (min_row, min_col, max_row, max_col)
    }

    pub fn contains(&self, row: u32, col: u32) -> bool {
        let (min_row, min_col, max_row, max_col) = self.bounds();
        row >= min_row && row <= max_row && col >= min_col && col <= max_col
    }

    pub fn is_single(&self) -> bool {
        self.start == self.end
    }
}

/// Clipboard for copy/paste
#[derive(Debug, Clone, Default)]
pub struct Clipboard {
    /// 2D array of cell values: clipboard[row][col]
    pub data: Vec<Vec<String>>,
}

/// Cell marking style
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CellMark {
    #[default]
    None,        // 1: Clear/reset
    YellowBg,    // 2: Yellow background - attention/review
    RedText,     // 3: Red text - error/warning
    GreenText,   // 4: Green text - OK/done
    BlueBg,      // 5: Blue background - category A
    MagentaText, // 6: Magenta text - category B
}

pub struct App<'a> {
    pub path: PathBuf,
    pub spreadsheet: Spreadsheet,
    pub current_sheet_index: usize,
    pub cursor: (u32, u32), // (row, col) 1-based
    pub selection: Selection,
    pub mode: Mode,
    pub scroll: (u32, u32), // (row_offset, col_offset) 0-based
    pub textarea: TextArea<'a>,
    pub should_quit: bool,
    pub column_widths: HashMap<u32, u16>, // col index -> width
    pub clipboard: Clipboard,
    pub status_message: Option<String>,
    pub viewport_size: (u16, u16), // (rows, cols) visible in grid
    pub cell_marks: HashMap<(usize, u32, u32), CellMark>, // (sheet_index, row, col) -> mark
    pub sheet_select_index: usize, // cursor position in sheet select mode
}

impl<'a> App<'a> {
    pub fn new(path: PathBuf) -> Result<Self> {
        let spreadsheet = if path.exists() {
            umya_spreadsheet::reader::xlsx::read(&path).map_err(|e| anyhow::anyhow!("Failed to read file: {}", e))?
        } else {
            let mut book = umya_spreadsheet::new_file();
            let _ = book.new_sheet("Sheet1");
            book
        };

        // Load existing cell marks from spreadsheet styles
        let cell_marks = Self::load_cell_marks_from_spreadsheet(&spreadsheet);

        Ok(Self {
            path,
            spreadsheet,
            current_sheet_index: 0,
            cursor: (1, 1),
            selection: Selection::single(1, 1),
            mode: Mode::View,
            scroll: (0, 0),
            textarea: TextArea::default(),
            should_quit: false,
            column_widths: HashMap::new(),
            clipboard: Clipboard::default(),
            status_message: None,
            viewport_size: (20, 10), // Default, will be updated by UI
            cell_marks,
            sheet_select_index: 0,
        })
    }

    fn load_cell_marks_from_spreadsheet(spreadsheet: &Spreadsheet) -> HashMap<(usize, u32, u32), CellMark> {
        let mut marks = HashMap::new();

        for (sheet_idx, sheet) in spreadsheet.get_sheet_collection().iter().enumerate() {
            for cell in sheet.get_cell_collection() {
                let coord = cell.get_coordinate();
                let col = *coord.get_col_num();
                let row_num = *coord.get_row_num();

                let style = cell.get_style();

                // Check background color
                if let Some(fill) = style.get_fill() {
                    if let Some(pattern_fill) = fill.get_pattern_fill() {
                        if let Some(fg_color) = pattern_fill.get_foreground_color() {
                            let argb = fg_color.get_argb();
                            if !argb.is_empty() {
                                let mark = Self::argb_to_bg_mark(argb);
                                if mark != CellMark::None {
                                    marks.insert((sheet_idx, row_num, col), mark);
                                    continue;
                                }
                            }
                        }
                    }
                }

                // Check font color
                if let Some(font) = style.get_font() {
                    let argb = font.get_color().get_argb();
                    if !argb.is_empty() {
                        let mark = Self::argb_to_font_mark(argb);
                        if mark != CellMark::None {
                            marks.insert((sheet_idx, row_num, col), mark);
                        }
                    }
                }
            }
        }

        marks
    }

    fn argb_to_bg_mark(argb: &str) -> CellMark {
        let argb_upper = argb.to_uppercase();
        match argb_upper.as_str() {
            // Yellow variations (including our adjusted FFFFEF00)
            "FFFFFF00" | "FFFF00" | "FFFFEF00" => CellMark::YellowBg,
            // Blue variations (including our adjusted FF0000FE)
            "FF0000FF" | "0000FF" | "FF0000FE" | "FF00BFFF" | "00BFFF" => CellMark::BlueBg,
            _ => CellMark::None,
        }
    }

    fn argb_to_font_mark(argb: &str) -> CellMark {
        let argb_upper = argb.to_uppercase();
        match argb_upper.as_str() {
            // Red variations (including our adjusted FFFF0001)
            "FFFF0000" | "FF0000" | "FFFF0001" => CellMark::RedText,
            // Green variations (including our adjusted FF008001)
            "FF008000" | "008000" | "FF008001" | "FF00FF00" | "00FF00" => CellMark::GreenText,
            // Magenta variations (including our adjusted FFFF00FE)
            "FFFF00FF" | "FF00FF" | "FFFF00FE" => CellMark::MagentaText,
            _ => CellMark::None,
        }
    }

    #[allow(dead_code)]
    pub fn on_tick(&mut self) {}

    pub fn on_key(&mut self, key: KeyEvent) {
        if key.kind != KeyEventKind::Press {
            return;
        }

        // Clear status message on any key press
        self.status_message = None;

        match self.mode {
            Mode::View => {
                let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
                let shift = key.modifiers.contains(KeyModifiers::SHIFT);

                match key.code {
                    KeyCode::Char('w') if ctrl => self.should_quit = true,
                    KeyCode::Char('s') if ctrl => {
                        match self.save_file() {
                            Ok(_) => self.status_message = Some(format!("Saved: {:?}", self.path)),
                            Err(e) => self.status_message = Some(format!("Error: {}", e)),
                        }
                    }
                    // Copy: C or F5
                    KeyCode::Char('c') if !ctrl => self.copy_selection(),
                    KeyCode::F(5) => self.copy_selection(),
                    // Paste: V or F6
                    KeyCode::Char('v') if !ctrl => self.paste_clipboard(),
                    KeyCode::F(6) => self.paste_clipboard(),
                    // Column width: E to expand, R to shrink
                    KeyCode::Char('e') if !ctrl && !shift => self.widen_column(),
                    KeyCode::Char('r') if !ctrl && !shift => self.shrink_column(),
                    KeyCode::F(2) => self.enter_edit_mode(),
                    // WASD movement (FPS style) + Shift for selection
                    KeyCode::Char('w') if !ctrl && shift => self.move_cursor(0, -1, true),
                    KeyCode::Char('s') if !ctrl && shift => self.move_cursor(0, 1, true),
                    KeyCode::Char('a') if !ctrl && shift => self.move_cursor(-1, 0, true),
                    KeyCode::Char('d') if !ctrl && shift => self.move_cursor(1, 0, true),
                    KeyCode::Char('w') if !ctrl => self.move_cursor(0, -1, false),
                    KeyCode::Char('a') if !ctrl => self.move_cursor(-1, 0, false),
                    KeyCode::Char('s') if !ctrl => self.move_cursor(0, 1, false),
                    KeyCode::Char('d') if !ctrl => self.move_cursor(1, 0, false),
                    // Arrow keys + Shift for selection
                    KeyCode::Enter if shift => self.move_cursor(0, -1, false),
                    KeyCode::Enter => self.move_cursor(0, 1, false),
                    KeyCode::Left if shift => self.move_cursor(-1, 0, true),
                    KeyCode::Right if shift => self.move_cursor(1, 0, true),
                    KeyCode::Up if shift => self.move_cursor(0, -1, true),
                    KeyCode::Down if shift => self.move_cursor(0, 1, true),
                    KeyCode::Left => self.move_cursor(-1, 0, false),
                    KeyCode::Right => self.move_cursor(1, 0, false),
                    KeyCode::Up => self.move_cursor(0, -1, false),
                    KeyCode::Down => self.move_cursor(0, 1, false),
                    KeyCode::Tab if shift => self.move_cursor(-1, 0, false),
                    KeyCode::Tab => self.move_cursor(1, 0, false),
                    KeyCode::BackTab => self.move_cursor(-1, 0, false),
                    KeyCode::PageUp => self.prev_sheet(),
                    KeyCode::PageDown => self.next_sheet(),
                    KeyCode::Home if ctrl => self.jump_to_start(),
                    KeyCode::End if ctrl => self.jump_to_end(),
                    KeyCode::Home => self.jump_to_row_start(),
                    KeyCode::End => self.jump_to_row_end(),
                    KeyCode::Esc => self.clear_selection(),
                    // Cell marking with number keys (1=clear, 2-6=colors)
                    KeyCode::Char('1') => self.set_mark_for_selection(CellMark::None),
                    KeyCode::Char('2') => self.set_mark_for_selection(CellMark::YellowBg),
                    KeyCode::Char('3') => self.set_mark_for_selection(CellMark::RedText),
                    KeyCode::Char('4') => self.set_mark_for_selection(CellMark::GreenText),
                    KeyCode::Char('5') => self.set_mark_for_selection(CellMark::BlueBg),
                    KeyCode::Char('6') => self.set_mark_for_selection(CellMark::MagentaText),
                    // F4: Enter sheet selection mode
                    KeyCode::F(4) => self.enter_sheet_select_mode(),
                    _ => {}
                }
            }
            Mode::SheetSelect => {
                match key.code {
                    KeyCode::Esc => self.mode = Mode::View,
                    KeyCode::Enter => self.confirm_sheet_selection(),
                    // Navigation: W/S or Up/Down
                    KeyCode::Char('w') | KeyCode::Up => self.sheet_select_move(-1),
                    KeyCode::Char('s') | KeyCode::Down => self.sheet_select_move(1),
                    _ => {}
                }
            }
            Mode::Edit => match key.code {
                KeyCode::Esc => self.mode = Mode::View,
                KeyCode::Enter => {
                    self.save_cell_value();
                    self.mode = Mode::View;
                    self.move_cursor(0, 1, false);
                }
                KeyCode::Tab => {
                    self.save_cell_value();
                    self.mode = Mode::View;
                    self.move_cursor(1, 0, false);
                }
                _ => {
                    self.textarea.input(key);
                }
            },
        }
    }

    fn move_cursor(&mut self, dx: i32, dy: i32, extend_selection: bool) {
        let (row, col) = self.cursor;
        let new_row = (row as i32 + dy).clamp(1, MAX_ROWS as i32) as u32;
        let new_col = (col as i32 + dx).clamp(1, MAX_COLUMNS as i32) as u32;
        self.cursor = (new_row, new_col);

        if extend_selection {
            // Extend selection from anchor
            self.selection.end = self.cursor;
        } else {
            // Reset selection to single cell
            self.selection = Selection::single(new_row, new_col);
        }

        self.adjust_scroll();
    }

    fn adjust_scroll(&mut self) {
        let (row, col) = self.cursor;
        let (view_rows, view_cols) = self.viewport_size;

        // Adjust vertical scroll
        if row <= self.scroll.0 {
            self.scroll.0 = row.saturating_sub(1);
        } else if row > self.scroll.0 + view_rows as u32 {
            self.scroll.0 = row - view_rows as u32;
        }

        // Adjust horizontal scroll
        if col <= self.scroll.1 {
            self.scroll.1 = col.saturating_sub(1);
        } else if col > self.scroll.1 + view_cols as u32 {
            self.scroll.1 = col - view_cols as u32;
        }
    }

    fn jump_to_start(&mut self) {
        self.cursor = (1, 1);
        self.selection = Selection::single(1, 1);
        self.scroll = (0, 0);
    }

    fn jump_to_end(&mut self) {
        // Find the last used cell
        if let Some(sheet) = self.spreadsheet.get_sheet(&self.current_sheet_index) {
            let max_row = sheet.get_highest_row();
            let max_col = sheet.get_highest_column();
            let row = max_row.max(1);
            let col = max_col.max(1);
            self.cursor = (row, col);
            self.selection = Selection::single(row, col);
            self.adjust_scroll();
        }
    }

    fn jump_to_row_start(&mut self) {
        self.cursor.1 = 1;
        self.selection = Selection::single(self.cursor.0, 1);
        self.adjust_scroll();
    }

    fn jump_to_row_end(&mut self) {
        // Find last used column in current row
        if let Some(sheet) = self.spreadsheet.get_sheet(&self.current_sheet_index) {
            let max_col = sheet.get_highest_column();
            let col = max_col.max(1);
            self.cursor.1 = col;
            self.selection = Selection::single(self.cursor.0, col);
            self.adjust_scroll();
        }
    }

    fn clear_selection(&mut self) {
        self.selection = Selection::single(self.cursor.0, self.cursor.1);
    }

    fn next_sheet(&mut self) {
        let count = self.spreadsheet.get_sheet_count();
        if count > 0 {
            self.current_sheet_index = (self.current_sheet_index + 1) % count;
        }
    }

    fn prev_sheet(&mut self) {
        let count = self.spreadsheet.get_sheet_count();
        if count > 0 {
            if self.current_sheet_index == 0 {
                self.current_sheet_index = count - 1;
            } else {
                self.current_sheet_index -= 1;
            }
        }
    }

    fn enter_edit_mode(&mut self) {
        // Check if this is a formula cell (read-only)
        if self.is_formula_cell(self.cursor.1, self.cursor.0) {
            // Show the formula in status message instead of editing
            if let Some(sheet) = self.spreadsheet.get_sheet(&self.current_sheet_index) {
                let formula = sheet.get_cell_value((self.cursor.1, self.cursor.0)).get_formula();
                self.status_message = Some(format!("Formula (read-only): ={}", formula));
            }
            return;
        }

        self.mode = Mode::Edit;
        if let Some(sheet) = self.spreadsheet.get_sheet(&self.current_sheet_index) {
            let value = sheet.get_cell_value((self.cursor.1, self.cursor.0));
            self.textarea = TextArea::from(vec![value.get_value().to_string()]);
        }
    }

    fn save_cell_value(&mut self) {
        if let Some(sheet) = self.spreadsheet.get_sheet_mut(&self.current_sheet_index) {
             let content = self.textarea.lines().join("\n");
             sheet.get_cell_mut((self.cursor.1, self.cursor.0)).set_value(content);
        }
    }

    fn save_file(&self) -> Result<()> {
        umya_spreadsheet::writer::xlsx::write(&self.spreadsheet, &self.path)
            .map_err(|e| anyhow::anyhow!("Failed to save file: {}", e))
    }

    fn widen_column(&mut self) {
        let col = self.cursor.1;
        let current = self.column_widths.get(&col).copied().unwrap_or(DEFAULT_COLUMN_WIDTH);
        let new_width = (current + COLUMN_WIDTH_STEP).min(MAX_COLUMN_WIDTH);
        self.column_widths.insert(col, new_width);
        if new_width >= MAX_COLUMN_WIDTH {
            self.status_message = Some(format!("Column width at maximum ({})", MAX_COLUMN_WIDTH));
        }
    }

    fn shrink_column(&mut self) {
        let col = self.cursor.1;
        let current = self.column_widths.get(&col).copied().unwrap_or(DEFAULT_COLUMN_WIDTH);
        let min_width: u16 = 3;
        let new_width = current.saturating_sub(COLUMN_WIDTH_STEP).max(min_width);
        if new_width <= min_width {
            self.column_widths.remove(&col);
            self.status_message = Some(format!("Column width at minimum ({})", min_width));
        } else {
            self.column_widths.insert(col, new_width);
        }
    }

    pub fn get_column_width(&self, col: u32) -> u16 {
        self.column_widths.get(&col).copied().unwrap_or(DEFAULT_COLUMN_WIDTH)
    }

    fn copy_selection(&mut self) {
        let (min_row, min_col, max_row, max_col) = self.selection.bounds();

        let mut data = Vec::new();
        if let Some(sheet) = self.spreadsheet.get_sheet(&self.current_sheet_index) {
            for r in min_row..=max_row {
                let mut row_data = Vec::new();
                for c in min_col..=max_col {
                    let value = sheet.get_cell_value((c, r)).get_value().to_string();
                    row_data.push(value);
                }
                data.push(row_data);
            }
        }

        let cells = (max_row - min_row + 1) * (max_col - min_col + 1);
        self.clipboard = Clipboard { data };
        self.status_message = Some(format!("Copied {} cell(s)", cells));
    }

    fn paste_clipboard(&mut self) {
        if self.clipboard.data.is_empty() {
            self.status_message = Some("Clipboard is empty".to_string());
            return;
        }

        let (start_row, start_col) = self.cursor;

        if let Some(sheet) = self.spreadsheet.get_sheet_mut(&self.current_sheet_index) {
            for (dr, row_data) in self.clipboard.data.iter().enumerate() {
                for (dc, value) in row_data.iter().enumerate() {
                    let target_row = start_row + dr as u32;
                    let target_col = start_col + dc as u32;

                    if target_row <= MAX_ROWS && target_col <= MAX_COLUMNS {
                        sheet.get_cell_mut((target_col, target_row)).set_value(value);
                    }
                }
            }
        }

        let rows = self.clipboard.data.len();
        let cols = self.clipboard.data.first().map(|r| r.len()).unwrap_or(0);
        self.status_message = Some(format!("Pasted {}x{} cells", rows, cols));
    }

    /// Check if a cell contains a formula
    pub fn is_formula_cell(&self, col: u32, row: u32) -> bool {
        if let Some(sheet) = self.spreadsheet.get_sheet(&self.current_sheet_index) {
            let cell_value = sheet.get_cell_value((col, row));
            !cell_value.get_formula().is_empty()
        } else {
            false
        }
    }

    /// Check if a format code represents a date/time format
    fn is_date_format(format_code: &str) -> bool {
        // Skip general and text formats
        if format_code == NumberingFormat::FORMAT_GENERAL || format_code == NumberingFormat::FORMAT_TEXT {
            return false;
        }
        // Check for date/time indicators in format code
        let lower = format_code.to_lowercase();
        lower.contains('y') || lower.contains('m') || lower.contains('d')
            || lower.contains('h') || lower.contains("am") || lower.contains("pm")
    }

    /// Convert date format to ISO format (yyyy-mm-dd or yyyy-mm-dd hh:mm:ss)
    fn normalize_date_format(format_code: &str) -> &'static str {
        let lower = format_code.to_lowercase();
        // Check if it includes time
        if lower.contains('h') || lower.contains("am") || lower.contains("pm") {
            // Date + Time
            if lower.contains('y') || lower.contains('d') {
                "yyyy-mm-dd hh:mm:ss"
            } else {
                // Time only
                "hh:mm:ss"
            }
        } else {
            // Date only
            "yyyy-mm-dd"
        }
    }

    /// Get cell value, truncated to fit column width with ellipsis
    pub fn get_cell_display(&self, col: u32, row: u32) -> String {
        if let Some(sheet) = self.spreadsheet.get_sheet(&self.current_sheet_index) {
            let cell_value = sheet.get_cell_value((col, row));
            let formula = cell_value.get_formula();
            let width = self.get_column_width(col) as usize;

            // For formula cells, show the cached result
            let display_value = if !formula.is_empty() {
                // Show calculated result (cached by Excel)
                let result = cell_value.get_value().to_string();
                if result.is_empty() {
                    "=...".to_string()  // Formula with no cached result
                } else {
                    result
                }
            } else {
                // Check for number format (date/time formatting)
                let raw_value = cell_value.get_value().to_string();

                // Try to get cell style and format
                if let Some(cell) = sheet.get_cell((col, row)) {
                    if let Some(num_fmt) = cell.get_style().get_number_format() {
                        let format_code = num_fmt.get_format_code();
                        if Self::is_date_format(format_code) {
                            // Apply ISO date formatting (yyyy-mm-dd)
                            let iso_format = Self::normalize_date_format(format_code);
                            to_formatted_string(&raw_value, iso_format)
                        } else if format_code != NumberingFormat::FORMAT_GENERAL {
                            // Apply other number formatting
                            to_formatted_string(&raw_value, format_code)
                        } else {
                            raw_value
                        }
                    } else {
                        raw_value
                    }
                } else {
                    raw_value
                }
            };

            if display_value.len() > width {
                // Truncate with ellipsis
                let truncated: String = display_value.chars().take(width.saturating_sub(1)).collect();
                format!("{}~", truncated)
            } else {
                display_value
            }
        } else {
            String::new()
        }
    }

    fn set_mark_for_selection(&mut self, mark: CellMark) {
        let (min_row, min_col, max_row, max_col) = self.selection.bounds();
        let sheet_idx = self.current_sheet_index;

        let mut count = 0;
        if let Some(sheet) = self.spreadsheet.get_sheet_mut(&sheet_idx) {
            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    let key = (sheet_idx, r, c);
                    if mark == CellMark::None {
                        self.cell_marks.remove(&key);
                    } else {
                        self.cell_marks.insert(key, mark);
                    }

                    // Apply style to Excel cell
                    let cell = sheet.get_cell_mut((c, r));
                    let style = cell.get_style_mut();

                    // Use slightly adjusted colors to avoid indexed color mapping bug
                    // umya-spreadsheet converts exact palette matches to indexed colors incorrectly
                    match mark {
                        CellMark::None => {
                            // Clear styles - reset to default
                            style.get_font_mut().set_color(Color::default().set_argb("FF000001").clone());
                            style.get_fill_mut().get_pattern_fill_mut().set_pattern_type(PatternValues::None);
                        }
                        CellMark::YellowBg => {
                            // Yellow slightly adjusted to avoid indexed: FFFFEF00
                            let bg = Color::default().set_argb("FFFFEF00").clone();
                            style.get_fill_mut().get_pattern_fill_mut()
                                .set_foreground_color(bg)
                                .set_pattern_type(PatternValues::Solid);
                            style.get_font_mut().set_color(Color::default().set_argb("FF000001").clone());
                        }
                        CellMark::RedText => {
                            // Red slightly adjusted: FFFF0001
                            style.get_font_mut().set_color(Color::default().set_argb("FFFF0001").clone());
                        }
                        CellMark::GreenText => {
                            // Dark green slightly adjusted: FF008001
                            style.get_font_mut().set_color(Color::default().set_argb("FF008001").clone());
                        }
                        CellMark::BlueBg => {
                            // Blue slightly adjusted: FF0000FE
                            let bg = Color::default().set_argb("FF0000FE").clone();
                            style.get_fill_mut().get_pattern_fill_mut()
                                .set_foreground_color(bg)
                                .set_pattern_type(PatternValues::Solid);
                            style.get_font_mut().set_color(Color::default().set_argb("FFFFFFFE").clone());
                        }
                        CellMark::MagentaText => {
                            // Magenta slightly adjusted: FFFF00FE
                            style.get_font_mut().set_color(Color::default().set_argb("FFFF00FE").clone());
                        }
                    }

                    count += 1;
                }
            }
        }

        let mark_name = match mark {
            CellMark::None => "cleared",
            CellMark::YellowBg => "yellow bg",
            CellMark::RedText => "red text",
            CellMark::GreenText => "green text",
            CellMark::BlueBg => "blue bg",
            CellMark::MagentaText => "magenta text",
        };
        self.status_message = Some(format!("Marked {} cell(s): {}", count, mark_name));
    }

    pub fn get_cell_mark(&self, row: u32, col: u32) -> CellMark {
        let key = (self.current_sheet_index, row, col);
        self.cell_marks.get(&key).copied().unwrap_or(CellMark::None)
    }

    fn enter_sheet_select_mode(&mut self) {
        self.sheet_select_index = self.current_sheet_index;
        self.mode = Mode::SheetSelect;
    }

    fn sheet_select_move(&mut self, delta: i32) {
        let count = self.spreadsheet.get_sheet_count();
        if count == 0 {
            return;
        }
        let new_index = (self.sheet_select_index as i32 + delta).rem_euclid(count as i32) as usize;
        self.sheet_select_index = new_index;
    }

    fn confirm_sheet_selection(&mut self) {
        self.current_sheet_index = self.sheet_select_index;
        self.mode = Mode::View;
    }

    pub fn get_sheet_names(&self) -> Vec<String> {
        self.spreadsheet.get_sheet_collection()
            .iter()
            .map(|s| s.get_name().to_string())
            .collect()
    }
}
