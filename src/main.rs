use crossterm::{
    event::{
        self, DisableMouseCapture, EnableMouseCapture, Event as CEvent, KeyCode, KeyModifiers,
    },
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ironcalc::{
    base::{
        expressions::types::Area, expressions::utils::number_to_column, UserModel,
        COLUMN_WIDTH_FACTOR,
    },
    export::save_to_xlsx,
    import::load_from_xlsx,
};
use ratatui::{
    backend::CrosstermBackend,
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style, Stylize},
    text::{Line, Span},
    widgets::{Block, BorderType, Borders, Cell, Clear, Paragraph, Row, Table},
    Terminal,
};
use std::time::{Duration, Instant};
use std::{io, sync::mpsc};
use std::{str::FromStr, thread};
use tui_input::{backend::crossterm::EventHandler, Input};

use std::env;

enum Event<I> {
    Input(I),
    Tick,
}

#[derive(PartialEq)]
enum CursorMode {
    Navigate,
    Input,
    Popup,
    ColorPicker,
    Help,
}

const HELP: [(&str, &str); 15] = [
    ("↑↓←→", "navigate cells"),
    ("PgUp/PgDn", "jump 10 rows"),
    ("e", "edit the current cell"),
    ("u / r", "undo/redo"),
    ("b", "background color"),
    ("c", "text color"),
    ("B / I / U", "bold/italic/underline"),
    ("S", "strikethrough"),
    ("< / >", "narrower/wider column"),
    ("- / =", "shorter/taller row"),
    ("f", "freeze/unfreeze panes"),
    ("s / a", "next/previous sheet"),
    ("+", "add a sheet"),
    ("?", "this help"),
    ("q", "quit (asks to save)"),
];

#[derive(PartialEq)]
enum ColorTarget {
    Background,
    Text,
}

// IronCalc column widths and row heights are in pixels: one terminal cell is
// COLUMN_WIDTH_FACTOR pixels wide and one terminal line is ROW_PX_PER_LINE
// pixels (the default row height) tall.
const ROW_PX_PER_LINE: f64 = 25.0;

// (name, hex). `None` clears the color back to the default.
const PALETTE: [(&str, Option<&str>); 12] = [
    ("Default", None),
    ("Black", Some("#000000")),
    ("White", Some("#FFFFFF")),
    ("Gray", Some("#B7B7B7")),
    ("Red", Some("#E06666")),
    ("Orange", Some("#F6B26B")),
    ("Yellow", Some("#FFD966")),
    ("Green", Some("#93C47D")),
    ("Teal", Some("#76A5AF")),
    ("Blue", Some("#6FA8DC")),
    ("Purple", Some("#8E7CC3")),
    ("Pink", Some("#C27BA0")),
];

fn main() -> Result<(), Box<dyn std::error::Error>> {
    enable_raw_mode()?;

    let args: Vec<String> = env::args().collect();
    let mut file_name = "model.xlsx";
    // UserModel wraps the engine model and keeps an undo/redo history.
    let mut model = if args.len() > 1 {
        file_name = &args[1];
        UserModel::from_model(load_from_xlsx(file_name, "en", "UTC", "en").unwrap())
    } else {
        UserModel::new_empty(file_name, "en", "UTC", "en").unwrap()
    };
    let mut selected_sheet = 0;
    let mut selected_row_index = 1;
    let mut selected_column_index = 1;
    let mut minimum_row_index = 1;
    let mut minimum_column_index = 1;
    let sheet_list_width = 20;
    let mut cursor_mode = CursorMode::Navigate;
    let mut input_formula = Input::default();

    let mut input_file_name: Input = file_name.into();

    let mut popup_open = false;
    let mut color_target = ColorTarget::Background;
    let mut color_picker_index = 0;

    let (tx, rx) = mpsc::channel();
    let tick_rate = Duration::from_millis(200);
    thread::spawn(move || {
        let mut last_tick = Instant::now();
        loop {
            let timeout = tick_rate
                .checked_sub(last_tick.elapsed())
                .unwrap_or_else(|| Duration::from_secs(0));

            if event::poll(timeout).expect("poll works") {
                if let CEvent::Key(key) = event::read().expect("can read events") {
                    tx.send(Event::Input(key)).expect("can send events");
                }
            }

            if last_tick.elapsed() >= tick_rate && tx.send(Event::Tick).is_ok() {
                last_tick = Instant::now();
            }
        }
    });

    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;

    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;
    terminal.clear()?;

    let header_style = Style::default().fg(Color::Black).bg(Color::Gray);
    let frozen_header_style = Style::default().fg(Color::White).bg(Color::DarkGray);
    let selected_header_style = Style::default()
        .fg(Color::Black)
        .bg(Color::Cyan)
        .add_modifier(Modifier::BOLD);

    let selected_cell_style = Style::default()
        .fg(Color::Black)
        .bg(Color::LightCyan)
        .add_modifier(Modifier::BOLD);

    let background_style = Style::default().bg(Color::Black);
    let selected_sheet_style = Style::default().bg(Color::White).fg(Color::LightMagenta);
    let non_selected_sheet_style = Style::default().fg(Color::White);
    let mut sheet_names = model.get_model().workbook.get_worksheet_names();
    loop {
        terminal.draw(|rect| {
            let size = rect.area();

            // Everything above a one-line status bar at the bottom
            let outer_chunks = Layout::default()
                .direction(Direction::Vertical)
                .constraints([Constraint::Min(3), Constraint::Length(1)].as_ref())
                .split(size);

            let global_chunks = Layout::default()
                .direction(Direction::Horizontal)
                .constraints([Constraint::Length(sheet_list_width), Constraint::Min(3)].as_ref())
                .split(outer_chunks[0]);

            // Sheet list to the left
            let sheets = Block::default()
                .borders(Borders::ALL)
                .style(Style::default().fg(Color::White))
                .title("Sheets")
                .border_type(BorderType::Plain)
                .style(background_style);
            let mut rows = vec![];
            (0..sheet_names.len()).for_each(|sheet_index| {
                let sheet_name = &sheet_names[sheet_index];
                let style = if sheet_index == selected_sheet {
                    selected_sheet_style
                } else {
                    non_selected_sheet_style
                };
                rows.push(Row::new(vec![Cell::from(sheet_name.clone()).style(style)]));
            });
            let widths = &[Constraint::Length(100)];
            let sheet_list = Table::new(rows, widths).block(sheets).column_spacing(0);

            rect.render_widget(sheet_list, global_chunks[0]);

            // The spreadsheet is the formula bar at the top and the sheet data
            let spreadsheet_chunks = Layout::default()
                .direction(Direction::Vertical)
                .margin(0)
                .constraints([Constraint::Length(1), Constraint::Min(2)].as_ref())
                .split(global_chunks[1]);

            let spreadsheet_width = size.width - sheet_list_width;
            // The formula bar and the status bar take one line each
            let spreadsheet_height = size.height.saturating_sub(2);
            let row_count = spreadsheet_height.saturating_sub(1);

            let status_bar = Paragraph::new(Line::from(vec![
                Span::styled(
                    " ?",
                    Style::default()
                        .fg(Color::LightGreen)
                        .add_modifier(Modifier::BOLD),
                ),
                Span::raw(" for help"),
            ]))
            .style(Style::default().bg(Color::DarkGray).fg(Color::White));
            rect.render_widget(status_bar, outer_chunks[1]);

            let first_row_width: u16 = 3;
            let available_width = spreadsheet_width.saturating_sub(first_row_width);
            let column_char_width = |column: i32| -> u16 {
                let width_px = model
                    .get_column_width(selected_sheet as u32, column)
                    .unwrap_or(10.0 * COLUMN_WIDTH_FACTOR);
                (width_px / COLUMN_WIDTH_FACTOR).round().max(1.0) as u16
            };
            let row_line_height = |row: u16| -> u16 {
                let height_px = model
                    .get_row_height(selected_sheet as u32, row as i32)
                    .unwrap_or(ROW_PX_PER_LINE);
                (height_px / ROW_PX_PER_LINE).round().max(1.0) as u16
            };
            let mut rows = vec![];
            // The first row in the column headers
            let mut row = Vec::new();
            // The first cell in that row is the top left square of the spreadsheet
            row.push(Cell::from(""));

            // Frozen rows and columns are always visible at the top/left of
            // the grid; only the region after them scrolls.
            let frozen_rows = model
                .get_frozen_rows_count(selected_sheet as u32)
                .unwrap_or(0)
                .max(0) as u16;
            let frozen_columns = model
                .get_frozen_columns_count(selected_sheet as u32)
                .unwrap_or(0)
                .max(0);
            let frozen_width: u16 = (1..=frozen_columns).map(column_char_width).sum();
            let frozen_height: u16 = (1..=frozen_rows).map(row_line_height).sum();
            let scroll_width = available_width.saturating_sub(frozen_width);
            let scroll_height = row_count.saturating_sub(frozen_height);

            // The scrolled region starts after the frozen panes.
            minimum_column_index = minimum_column_index.max(frozen_columns + 1);
            minimum_row_index = minimum_row_index.max(frozen_rows + 1);

            // We want to make sure the selected cell is fully visible. A
            // selected cell inside the frozen panes is always visible.
            if selected_column_index > frozen_columns {
                if selected_column_index < minimum_column_index {
                    minimum_column_index = selected_column_index;
                }
                while minimum_column_index < selected_column_index {
                    let width: u16 = (minimum_column_index..=selected_column_index)
                        .map(column_char_width)
                        .sum();
                    if width <= scroll_width {
                        break;
                    }
                    minimum_column_index += 1;
                }
            }
            if selected_row_index > frozen_rows {
                if selected_row_index < minimum_row_index {
                    minimum_row_index = selected_row_index;
                }
                while minimum_row_index < selected_row_index {
                    let height: u16 = (minimum_row_index..=selected_row_index)
                        .map(row_line_height)
                        .sum();
                    if height <= scroll_height {
                        break;
                    }
                    minimum_row_index += 1;
                }
            }

            // Visible columns and rows with their sizes in terminal cells:
            // first the frozen ones, then the scrolled region. The last column
            // is capped to the remaining space: the widths must add up to
            // exactly the available width or the Table widget will flex-shrink
            // every column.
            let mut visible_columns: Vec<(i32, u16)> = Vec::new();
            let mut used_width = 0;
            let mut column_index = 1;
            while used_width < available_width && column_index <= frozen_columns {
                let width = column_char_width(column_index).min(available_width - used_width);
                visible_columns.push((column_index, width));
                used_width += width;
                column_index += 1;
            }
            let mut column_index = minimum_column_index;
            while used_width < available_width {
                let width = column_char_width(column_index).min(available_width - used_width);
                visible_columns.push((column_index, width));
                used_width += width;
                column_index += 1;
            }
            let mut visible_rows: Vec<(u16, u16)> = Vec::new();
            let mut used_height = 0;
            let mut row_index = 1;
            while used_height < row_count && row_index <= frozen_rows {
                let height = row_line_height(row_index);
                visible_rows.push((row_index, height));
                used_height += height;
                row_index += 1;
            }
            let mut row_index = minimum_row_index;
            while used_height < row_count {
                let height = row_line_height(row_index);
                visible_rows.push((row_index, height));
                used_height += height;
                row_index += 1;
            }

            for (column_index, width) in &visible_columns {
                let column_str = number_to_column(*column_index);
                let style = if *column_index == selected_column_index {
                    selected_header_style
                } else if *column_index <= frozen_columns {
                    frozen_header_style
                } else {
                    header_style
                };
                row.push(
                    Cell::from(format!(
                        "{:^width$}",
                        column_str.unwrap(),
                        width = *width as usize
                    ))
                    .style(style),
                );
            }
            rows.push(Row::new(row));
            for (row_index, row_height) in &visible_rows {
                let row_index = *row_index;
                let mut row = Vec::new();
                let style = if row_index == selected_row_index {
                    selected_header_style
                } else if row_index <= frozen_rows {
                    frozen_header_style
                } else {
                    header_style
                };
                row.push(Cell::from(format!("{}", row_index)).style(style));
                for (column_index, _) in &visible_columns {
                    let column_index = *column_index;
                    let value = model
                        .get_formatted_cell_value(
                            selected_sheet as u32,
                            row_index as i32,
                            column_index,
                        )
                        .unwrap();
                    let cell_style = model
                        .get_cell_style(selected_sheet as u32, row_index as i32, column_index)
                        .unwrap();
                    let mut style = if selected_row_index == row_index
                        && selected_column_index == column_index
                    {
                        selected_cell_style
                    } else {
                        let theme = &model.get_model().workbook.theme;
                        let bg_rgb = cell_style.fill.color.to_rgb(theme);
                        let bg_color = if bg_rgb.is_empty() {
                            Color::White
                        } else {
                            Color::from_str(&bg_rgb).unwrap_or(Color::White)
                        };
                        let fg_rgb = cell_style.font.color.to_rgb(theme);
                        let fg_color = if fg_rgb.is_empty() {
                            Color::Black
                        } else {
                            Color::from_str(&fg_rgb).unwrap_or(Color::Black)
                        };
                        Style::default().fg(fg_color).bg(bg_color)
                    };
                    if cell_style.font.b {
                        style = style.add_modifier(Modifier::BOLD);
                    }
                    if cell_style.font.i {
                        style = style.add_modifier(Modifier::ITALIC);
                    }
                    if cell_style.font.u {
                        style = style.add_modifier(Modifier::UNDERLINED);
                    }
                    if cell_style.font.strike {
                        style = style.add_modifier(Modifier::CROSSED_OUT);
                    }
                    row.push(Cell::from(value.to_string()).style(style));
                }
                rows.push(Row::new(row).height(*row_height));
            }
            let mut widths = Vec::new();
            widths.push(Constraint::Length(first_row_width));
            for (_, width) in &visible_columns {
                widths.push(Constraint::Length(*width));
            }
            let spreadsheet = Table::new(rows, widths)
                .block(Block::default().style(Style::default().bg(Color::Black)))
                .column_spacing(0);

            let text = if cursor_mode != CursorMode::Input {
                model
                    .get_cell_content(
                        selected_sheet as u32,
                        selected_row_index as i32,
                        selected_column_index,
                    )
                    .unwrap_or_default()
            } else {
                input_formula.value().to_string()
            };
            let cell_address_text = format!(
                "{}{}: ",
                number_to_column(selected_column_index).unwrap(),
                selected_row_index,
            );
            let formula_bar_text = format!("{}{}", cell_address_text, text,);
            let formula_bar = Paragraph::new(vec![Line::from(vec![Span::raw(formula_bar_text)])]);
            rect.render_widget(formula_bar.block(Block::default()), spreadsheet_chunks[0]);
            rect.render_widget(spreadsheet, spreadsheet_chunks[1]);
            if cursor_mode == CursorMode::Input {
                let area = spreadsheet_chunks[0];
                rect.set_cursor_position((
                    area.x
                        + (input_formula.visual_cursor() as u16)
                        + cell_address_text.len() as u16,
                    area.y,
                ))
            }

            if popup_open {
                let area = centered_rect(60, 20, size);
                rect.render_widget(Clear, area);
                let input_text = input_file_name.value();
                let text = vec![
                    Line::from(vec![input_text.fg(Color::Yellow)]),
                    "".into(),
                    Line::from(vec![
                        "Esc".green(),
                        " to abort. ".into(),
                        "Ctrl+Q".green(),
                        " to quit without saving. ".into(),
                        "Enter".green(),
                        " to save and quit".into(),
                    ]),
                ];
                rect.render_widget(
                    Paragraph::new(text).block(Block::bordered().title("Save as")),
                    area,
                );
                rect.set_cursor_position((
                    // Put cursor past the end of the input text
                    area.x + (input_file_name.visual_cursor() as u16) + 1,
                    // Move one line own, from the border to the input line
                    area.y + 1,
                ))
            }

            if cursor_mode == CursorMode::Help {
                let width = 40.min(size.width);
                let height = (HELP.len() as u16 + 3).min(size.height);
                let area = Rect {
                    x: size.width.saturating_sub(width) / 2,
                    y: size.height.saturating_sub(height) / 2,
                    width,
                    height,
                };
                rect.render_widget(Clear, area);
                let mut lines = Vec::new();
                for (key, description) in HELP {
                    lines.push(Line::from(vec![
                        Span::styled(
                            format!(" {:>9}", key),
                            Style::default()
                                .fg(Color::Green)
                                .add_modifier(Modifier::BOLD),
                        ),
                        Span::raw("  "),
                        Span::raw(description),
                    ]));
                }
                lines.push(Line::from(vec![Span::styled(
                    " press any key to close",
                    Style::default().fg(Color::DarkGray),
                )]));
                rect.render_widget(
                    Paragraph::new(lines).block(Block::bordered().title("Keys")),
                    area,
                );
            }

            if cursor_mode == CursorMode::ColorPicker {
                let title = match color_target {
                    ColorTarget::Background => "Background color",
                    ColorTarget::Text => "Text color",
                };
                let width = 24.min(size.width);
                let height = (PALETTE.len() as u16 + 3).min(size.height);
                let area = Rect {
                    x: size.width.saturating_sub(width) / 2,
                    y: size.height.saturating_sub(height) / 2,
                    width,
                    height,
                };
                rect.render_widget(Clear, area);
                let mut lines = Vec::new();
                for (index, (name, hex)) in PALETTE.iter().enumerate() {
                    let swatch = match hex {
                        Some(hex) => Span::styled(
                            "██ ",
                            Style::default().fg(Color::from_str(hex).unwrap_or(Color::White)),
                        ),
                        None => Span::raw("·· "),
                    };
                    let (marker, name_style) = if index == color_picker_index {
                        ("› ", Style::default().add_modifier(Modifier::BOLD))
                    } else {
                        ("  ", Style::default())
                    };
                    lines.push(Line::from(vec![
                        Span::raw(marker),
                        swatch,
                        Span::styled(*name, name_style),
                    ]));
                }
                lines.push(Line::from(vec![
                    " Enter".green(),
                    " set ".into(),
                    "Esc".green(),
                    " cancel".into(),
                ]));
                rect.render_widget(
                    Paragraph::new(lines).block(Block::bordered().title(title)),
                    area,
                );
            }
        })?;

        match cursor_mode {
            CursorMode::Popup => {
                match rx.recv()? {
                    Event::Input(event) => match event.code {
                        // Ctrl+Q quits without saving (End kept as a fallback)
                        KeyCode::Char('q') | KeyCode::Char('Q')
                            if event.modifiers.contains(KeyModifiers::CONTROL) =>
                        {
                            terminal.clear()?;
                            // restore terminal
                            disable_raw_mode()?;
                            execute!(
                                terminal.backend_mut(),
                                LeaveAlternateScreen,
                                DisableMouseCapture
                            )?;
                            terminal.show_cursor()?;
                            break;
                        }
                        KeyCode::End => {
                            terminal.clear()?;
                            // restore terminal
                            disable_raw_mode()?;
                            execute!(
                                terminal.backend_mut(),
                                LeaveAlternateScreen,
                                DisableMouseCapture
                            )?;
                            terminal.show_cursor()?;
                            break;
                        }
                        KeyCode::Enter => {
                            terminal.clear()?;
                            // restore terminal
                            disable_raw_mode()?;
                            execute!(
                                terminal.backend_mut(),
                                LeaveAlternateScreen,
                                DisableMouseCapture
                            )?;
                            terminal.show_cursor()?;
                            let _ = save_to_xlsx(model.get_model(), input_file_name.value());
                            break;
                        }
                        KeyCode::Esc => {
                            popup_open = false;
                            cursor_mode = CursorMode::Navigate;
                        }
                        _ => {
                            input_file_name.handle_event(&CEvent::Key(event));
                        }
                    },
                    Event::Tick => {}
                }
            }
            CursorMode::Navigate => {
                match rx.recv()? {
                    Event::Input(event) => match event.code {
                        KeyCode::Char('q') => {
                            popup_open = true;
                            cursor_mode = CursorMode::Popup;
                        }
                        KeyCode::Down => {
                            selected_row_index += 1;
                        }
                        KeyCode::Up => {
                            if selected_row_index > 1 {
                                selected_row_index -= 1;
                            }
                        }
                        KeyCode::Right => {
                            selected_column_index += 1;
                        }
                        KeyCode::Left => {
                            if selected_column_index > 1 {
                                selected_column_index -= 1;
                            }
                        }
                        KeyCode::PageDown => {
                            selected_row_index += 10;
                        }
                        KeyCode::PageUp => {
                            if selected_row_index > 10 {
                                selected_row_index -= 10;
                            } else {
                                selected_row_index = 1;
                            }
                        }
                        KeyCode::Char('s') => {
                            selected_sheet += 1;
                            if selected_sheet >= sheet_names.len() {
                                selected_sheet = 0;
                            }
                        }
                        KeyCode::Char('a') => {
                            selected_sheet = selected_sheet.saturating_sub(1);
                        }
                        KeyCode::Char('e') => {
                            cursor_mode = CursorMode::Input;
                            let input_str = model
                                .get_cell_content(
                                    selected_sheet as u32,
                                    selected_row_index as i32,
                                    selected_column_index,
                                )
                                .unwrap_or_default();
                            input_formula = input_formula.with_value(input_str);
                        }
                        KeyCode::Char('+') => {
                            let _ = model.new_sheet();
                            sheet_names = model.get_model().workbook.get_worksheet_names();
                        }
                        KeyCode::Char('u') => {
                            let _ = model.undo();
                        }
                        KeyCode::Char('r') => {
                            let _ = model.redo();
                        }
                        KeyCode::Char('>') | KeyCode::Char('.') => {
                            let sheet = selected_sheet as u32;
                            let column = selected_column_index;
                            if let Ok(width) = model.get_column_width(sheet, column) {
                                let _ = model.set_columns_width(
                                    sheet,
                                    column,
                                    column,
                                    width + COLUMN_WIDTH_FACTOR,
                                );
                            }
                        }
                        KeyCode::Char('<') | KeyCode::Char(',') => {
                            let sheet = selected_sheet as u32;
                            let column = selected_column_index;
                            if let Ok(width) = model.get_column_width(sheet, column) {
                                let width = (width - COLUMN_WIDTH_FACTOR).max(COLUMN_WIDTH_FACTOR);
                                let _ = model.set_columns_width(sheet, column, column, width);
                            }
                        }
                        KeyCode::Char('=') => {
                            let sheet = selected_sheet as u32;
                            let row = selected_row_index as i32;
                            if let Ok(height) = model.get_row_height(sheet, row) {
                                let _ = model.set_rows_height(
                                    sheet,
                                    row,
                                    row,
                                    height + ROW_PX_PER_LINE,
                                );
                            }
                        }
                        KeyCode::Char('-') => {
                            let sheet = selected_sheet as u32;
                            let row = selected_row_index as i32;
                            if let Ok(height) = model.get_row_height(sheet, row) {
                                let height = (height - ROW_PX_PER_LINE).max(ROW_PX_PER_LINE);
                                let _ = model.set_rows_height(sheet, row, row, height);
                            }
                        }
                        KeyCode::Char(c @ ('B' | 'I' | 'U' | 'S')) => {
                            let sheet = selected_sheet as u32;
                            let row = selected_row_index as i32;
                            let column = selected_column_index;
                            if let Ok(style) = model.get_cell_style(sheet, row, column) {
                                let (path, on) = match c {
                                    'B' => ("font.b", style.font.b),
                                    'I' => ("font.i", style.font.i),
                                    'U' => ("font.u", style.font.u),
                                    _ => ("font.strike", style.font.strike),
                                };
                                let range = Area {
                                    sheet,
                                    row,
                                    column,
                                    width: 1,
                                    height: 1,
                                };
                                let value = if on { "false" } else { "true" };
                                let _ = model.update_range_style(&range, path, value);
                            }
                        }
                        KeyCode::Char('?') => {
                            cursor_mode = CursorMode::Help;
                        }
                        KeyCode::Char('f') => {
                            let sheet = selected_sheet as u32;
                            let new_rows = (selected_row_index - 1) as i32;
                            let new_columns = selected_column_index - 1;
                            let rows = model.get_frozen_rows_count(sheet).unwrap_or(0);
                            let columns = model.get_frozen_columns_count(sheet).unwrap_or(0);
                            // Freezing at the same spot (or at A1) unfreezes.
                            if new_rows == rows && new_columns == columns {
                                let _ = model.set_frozen_rows_count(sheet, 0);
                                let _ = model.set_frozen_columns_count(sheet, 0);
                            } else {
                                let _ = model.set_frozen_rows_count(sheet, new_rows);
                                let _ = model.set_frozen_columns_count(sheet, new_columns);
                            }
                        }
                        KeyCode::Char('b') => {
                            color_target = ColorTarget::Background;
                            color_picker_index = 0;
                            cursor_mode = CursorMode::ColorPicker;
                        }
                        KeyCode::Char('c') => {
                            color_target = ColorTarget::Text;
                            color_picker_index = 0;
                            cursor_mode = CursorMode::ColorPicker;
                        }
                        _ => {
                            // println!("{:?}", event);
                        }
                    },
                    Event::Tick => {}
                }
            }
            CursorMode::Help => match rx.recv()? {
                Event::Input(_) => {
                    cursor_mode = CursorMode::Navigate;
                }
                Event::Tick => {}
            },
            CursorMode::ColorPicker => match rx.recv()? {
                Event::Input(event) => match event.code {
                    KeyCode::Esc => {
                        cursor_mode = CursorMode::Navigate;
                    }
                    KeyCode::Up => {
                        color_picker_index =
                            (color_picker_index + PALETTE.len() - 1) % PALETTE.len();
                    }
                    KeyCode::Down => {
                        color_picker_index = (color_picker_index + 1) % PALETTE.len();
                    }
                    KeyCode::Enter => {
                        let (_, hex) = PALETTE[color_picker_index];
                        let range = Area {
                            sheet: selected_sheet as u32,
                            row: selected_row_index as i32,
                            column: selected_column_index,
                            width: 1,
                            height: 1,
                        };
                        let path = match color_target {
                            ColorTarget::Background => "fill.color",
                            ColorTarget::Text => "font.color",
                        };
                        // An empty value clears the color back to the default.
                        let _ = model.update_range_style(&range, path, hex.unwrap_or(""));
                        cursor_mode = CursorMode::Navigate;
                    }
                    _ => {}
                },
                Event::Tick => {}
            },
            CursorMode::Input => match rx.recv()? {
                Event::Input(event) => match event.code {
                    // KeyCode::Char(c) => {
                    //     input_str.push(c);
                    // }
                    // KeyCode::Backspace => {
                    //     input_str.pop();
                    // }
                    KeyCode::Enter => {
                        cursor_mode = CursorMode::Navigate;
                        let sheet = selected_sheet as u32;
                        let row = selected_row_index as i32;
                        let column = selected_column_index;
                        let _ = model.set_user_input(sheet, row, column, input_formula.value());
                    }
                    _ => {
                        input_formula.handle_event(&CEvent::Key(event));
                    }
                },
                Event::Tick => {}
            },
        }
    }

    Ok(())
}

/// helper function to create a centered rect using up certain percentage of the available rect `r`
fn centered_rect(percent_x: u16, percent_y: u16, r: Rect) -> Rect {
    let popup_layout = Layout::vertical([
        Constraint::Percentage((100 - percent_y) / 2),
        Constraint::Percentage(percent_y),
        Constraint::Percentage((100 - percent_y) / 2),
    ])
    .split(r);

    Layout::horizontal([
        Constraint::Percentage((100 - percent_x) / 2),
        Constraint::Percentage(percent_x),
        Constraint::Percentage((100 - percent_x) / 2),
    ])
    .split(popup_layout[1])[1]
}
