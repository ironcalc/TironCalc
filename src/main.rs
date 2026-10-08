use crossterm::{
    event::{
        self, DisableMouseCapture, EnableMouseCapture, Event as CEvent, KeyCode, KeyEventKind,
        KeyModifiers,
    },
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ironcalc::{
    base::{
        expressions::types::Area,
        expressions::utils::number_to_column,
        types::{HorizontalAlignment, MergedCell},
        UserModel, COLUMN_WIDTH_FACTOR,
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
use std::io::Write;
use std::sync::atomic::{AtomicBool, Ordering};
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

const HELP: [(&str, &str); 17] = [
    ("↑↓←→", "navigate cells"),
    ("Shift+↑↓←→", "extend the selection"),
    ("PgUp/PgDn", "jump 10 rows"),
    ("e", "edit the current cell"),
    ("u / r", "undo/redo"),
    ("b", "background color"),
    ("c", "text color"),
    ("B / I / U", "bold/italic/underline"),
    ("S", "strikethrough"),
    ("< / >", "narrower/wider column"),
    ("- / =", "shorter/taller row"),
    ("m / M", "merge (& center)/unmerge"),
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

// Sheet size limits (IronCalc keeps its own copies crate-private)
const LAST_ROW: i32 = 1_048_576;
const LAST_COLUMN: i32 = 16_384;

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
        // Large workbooks take a while to load: show a spinner meanwhile.
        // It waits one frame before drawing so small files load silently.
        let loading = AtomicBool::new(true);
        let workbook = thread::scope(|scope| {
            scope.spawn(|| {
                let start = Instant::now();
                for frame in "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏".chars().cycle() {
                    thread::sleep(Duration::from_millis(100));
                    if !loading.load(Ordering::Relaxed) {
                        break;
                    }
                    // Spinner in IronCalc orange (#F2994A)
                    print!(
                        "\r\x1b[38;2;242;153;74m{frame}\x1b[0m Loading {file_name}… {:.1}s",
                        start.elapsed().as_secs_f64()
                    );
                    let _ = io::stdout().flush();
                }
                // Clear the spinner line
                print!("\r\x1b[2K");
                let _ = io::stdout().flush();
            });
            let workbook = load_from_xlsx(file_name, "en", "UTC", "en");
            loading.store(false, Ordering::Relaxed);
            workbook
        });
        UserModel::from_model(workbook.unwrap())
    } else {
        UserModel::new_empty(file_name, "en", "UTC", "en").unwrap()
    };
    let mut selected_sheet = 0;
    let mut selected_row_index = 1;
    let mut selected_column_index = 1;
    // Shift+arrows extend the selection from the selected cell to this one
    let mut end_row = selected_row_index;
    let mut end_column = selected_column_index;
    // Whole row/column selection; both together select every cell.
    let mut whole_row = false;
    let mut whole_column = false;
    let mut minimum_row_index = 1;
    let mut minimum_column_index = 1;
    let sheet_list_width = 20;
    let mut cursor_mode = CursorMode::Navigate;
    let mut input_formula = Input::default();

    let mut input_file_name: Input = file_name.into();

    let mut popup_open = false;
    let mut color_target = ColorTarget::Background;
    let mut color_picker_index = 0;
    // Shown in the status bar until the next key press (e.g. why a merge failed)
    let mut status_message: Option<String> = None;

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
                    if key.kind != KeyEventKind::Release {
                        tx.send(Event::Input(key)).expect("can send events");
                    }
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

    // IronCalc brand orange
    let ironcalc_orange = Color::Rgb(0xF2, 0x99, 0x4A);
    let header_style = Style::default()
        .fg(Color::Black)
        .bg(Color::Rgb(0xC8, 0xC8, 0xC8));
    let frozen_header_style = Style::default()
        .fg(Color::White)
        .bg(Color::Rgb(0x6E, 0x6E, 0x6E));
    let selected_header_style = Style::default()
        .fg(Color::Black)
        .bg(Color::Rgb(0xA8, 0xA8, 0xA8))
        .add_modifier(Modifier::BOLD);
    let frozen_line_color = Color::Rgb(0x50, 0x50, 0x50);
    let frozen_line_style = Style::default().fg(frozen_line_color).bg(Color::White);

    let selected_cell_style = Style::default()
        .fg(Color::Black)
        .bg(ironcalc_orange)
        .add_modifier(Modifier::BOLD);
    // Light tint of the orange for the rest of a selected row/column
    let selection_bg = Color::Rgb(0xFC, 0xE0, 0xC8);

    let background_style = Style::default().bg(Color::Black);
    let selected_sheet_style = Style::default().bg(Color::White).fg(Color::LightMagenta);
    let non_selected_sheet_style = Style::default().fg(Color::White);
    let mut sheet_names = model.get_model().workbook.get_worksheet_names();
    loop {
        // Merged cells of the selected sheet; they drive both the rendering
        // and the navigation.
        let merged_cells = model
            .get_merged_cells(selected_sheet as u32)
            .unwrap_or_default();
        // A covered cell can never be selected: snap to the anchor of its
        // merged cell. This covers sheet changes and undo/redo too.
        if let Some(m) = merged_cell_containing(
            &merged_cells,
            selected_row_index as i32,
            selected_column_index,
        ) {
            if (m.row, m.column) != (selected_row_index as i32, selected_column_index) {
                selected_row_index = m.row as u16;
                selected_column_index = m.column;
                end_row = selected_row_index;
                end_column = selected_column_index;
            }
        }
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

            let key_style = Style::default()
                .fg(ironcalc_orange)
                .add_modifier(Modifier::BOLD);
            let mut footer = vec![Span::styled(" ?", key_style), Span::raw(" for help")];
            if let Some(message) = &status_message {
                footer.push(Span::styled(
                    format!("  {message}"),
                    Style::default().fg(Color::Yellow),
                ));
            } else {
                for (key, description) in HELP
                    .iter()
                    .filter(|(key, _)| ["e", "u / r", "m / M", "f", "q"].contains(key))
                {
                    footer.push(Span::styled(format!("  {key}"), key_style));
                    footer.push(Span::raw(format!(" {description}")));
                }
            }
            let status_bar = Paragraph::new(Line::from(footer))
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
            // Reserve one cell for the frozen pane separator lines
            let available_width = available_width.saturating_sub((frozen_columns > 0) as u16);
            let row_count = row_count.saturating_sub((frozen_rows > 0) as u16);
            let frozen_width: u16 = (1..=frozen_columns).map(column_char_width).sum();
            let frozen_height: u16 = (1..=frozen_rows).map(row_line_height).sum();
            let scroll_width = available_width.saturating_sub(frozen_width);
            let scroll_height = row_count.saturating_sub(frozen_height);

            // The scrolled region starts after the frozen panes.
            minimum_column_index = minimum_column_index.max(frozen_columns + 1);
            minimum_row_index = minimum_row_index.max(frozen_rows + 1);

            // We want to make sure the moving end of the selection is fully
            // visible. A cell inside the frozen panes is always visible.
            if end_column > frozen_columns {
                if end_column < minimum_column_index {
                    minimum_column_index = end_column;
                }
                while minimum_column_index < end_column {
                    let width: u16 = (minimum_column_index..=end_column)
                        .map(column_char_width)
                        .sum();
                    if width <= scroll_width {
                        break;
                    }
                    minimum_column_index += 1;
                }
            }
            if end_row > frozen_rows {
                if end_row < minimum_row_index {
                    minimum_row_index = end_row;
                }
                while minimum_row_index < end_row {
                    let height: u16 = (minimum_row_index..=end_row).map(row_line_height).sum();
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

            // The selection grows over the merged cells it touches
            let selection = selected_area(
                selected_sheet as u32,
                (selected_row_index as i32, end_row as i32),
                (selected_column_index, end_column),
                whole_row,
                whole_column,
                &merged_cells,
            );
            let selected_rows = selection.row..=selection.row + selection.height - 1;
            let selected_columns = selection.column..=selection.column + selection.width - 1;
            let row_in_selection = |row: u16| whole_column || selected_rows.contains(&(row as i32));
            let column_in_selection = |column: i32| whole_row || selected_columns.contains(&column);
            // The style of a cell on screen from its own style and the selection
            let theme = &model.get_model().workbook.theme;
            let cell_display_style =
                |row_index: u16, column_index: i32, cell_style: &ironcalc::base::types::Style| {
                    // With a whole row/column selected the header carries
                    // the orange, so the active cell just gets the tint.
                    let mut style = if selected_row_index == row_index
                        && selected_column_index == column_index
                        && !whole_row
                        && !whole_column
                    {
                        selected_cell_style
                    } else {
                        let bg_rgb = cell_style.fill.color.to_rgb(theme);
                        let bg_color =
                            if row_in_selection(row_index) && column_in_selection(column_index) {
                                selection_bg
                            } else if bg_rgb.is_empty() {
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
                    style
                };
            for (column_index, width) in &visible_columns {
                let column_str = number_to_column(*column_index);
                let in_selection = column_in_selection(*column_index);
                // Orange when the whole column is selected
                let style = if in_selection && whole_column {
                    selected_cell_style
                } else if in_selection {
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
                if *column_index == frozen_columns {
                    row.push(Cell::from("│").style(header_style.fg(frozen_line_color)));
                }
            }
            rows.push(Row::new(row));
            for (row_index, row_height) in &visible_rows {
                let row_index = *row_index;
                let mut row = Vec::new();
                let in_selection = row_in_selection(row_index);
                // Orange when the whole row is selected
                let style = if in_selection && whole_row {
                    selected_cell_style
                } else if in_selection {
                    selected_header_style
                } else if row_index <= frozen_rows {
                    frozen_header_style
                } else {
                    header_style
                };
                row.push(Cell::from(format!("{}", row_index)).style(style));
                for (column_index, width) in &visible_columns {
                    let column_index = *column_index;
                    let cell_style = model
                        .get_cell_style(selected_sheet as u32, row_index as i32, column_index)
                        .unwrap();
                    // Merged cells are painted as a whole over the table
                    // below, so here they are just blank.
                    let value =
                        if merged_cell_containing(&merged_cells, row_index as i32, column_index)
                            .is_some()
                        {
                            String::new()
                        } else {
                            let value = model
                                .get_formatted_cell_value(
                                    selected_sheet as u32,
                                    row_index as i32,
                                    column_index,
                                )
                                .unwrap();
                            // Wrapped text flows into the lines of a tall row
                            if cell_style.alignment.as_ref().is_some_and(|a| a.wrap_text) {
                                wrap_text(&value, *width as usize).join("\n")
                            } else {
                                value
                            }
                        };
                    let style = cell_display_style(row_index, column_index, &cell_style);
                    row.push(Cell::from(value).style(style));
                    if column_index == frozen_columns {
                        row.push(
                            Cell::from("│\n".repeat(*row_height as usize)).style(frozen_line_style),
                        );
                    }
                }
                rows.push(Row::new(row).height(*row_height));
                if row_index == frozen_rows {
                    let mut line = vec![Cell::from("─".repeat(first_row_width as usize))
                        .style(header_style.fg(frozen_line_color))];
                    for (column_index, width) in &visible_columns {
                        line.push(Cell::from("─".repeat(*width as usize)).style(frozen_line_style));
                        if *column_index == frozen_columns {
                            line.push(Cell::from("┼").style(frozen_line_style));
                        }
                    }
                    rows.push(Row::new(line));
                }
            }
            let mut widths = Vec::new();
            widths.push(Constraint::Length(first_row_width));
            for (column_index, width) in &visible_columns {
                widths.push(Constraint::Length(*width));
                if *column_index == frozen_columns {
                    widths.push(Constraint::Length(1));
                }
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

            // Merged cells are painted over the table as a single block
            // spanning their visible rows and columns, with the anchor's
            // content and style. Like in the web UI a merged cell reaching
            // from a frozen pane into the scrolled region is painted once per
            // pane, clipped to it, so the pane separator stays visible.
            let table_area = spreadsheet_chunks[1];
            // Screen position and size of every visible column and row
            let mut column_tracks: Vec<Track> = Vec::new();
            let mut x = table_area.x + first_row_width;
            for (column_index, width) in &visible_columns {
                column_tracks.push(Track {
                    index: *column_index,
                    start: x,
                    size: *width,
                });
                x += width;
                if *column_index == frozen_columns {
                    x += 1;
                }
            }
            let mut row_tracks: Vec<Track> = Vec::new();
            let mut y = table_area.y + 1;
            for (row_index, height) in &visible_rows {
                row_tracks.push(Track {
                    index: *row_index as i32,
                    start: y,
                    size: *height,
                });
                y += height;
                if *row_index == frozen_rows {
                    y += 1;
                }
            }
            for m in &merged_cells {
                let sheet = selected_sheet as u32;
                let value = model
                    .get_formatted_cell_value(sheet, m.row, m.column)
                    .unwrap_or_default();
                let cell_style = model.get_cell_style(sheet, m.row, m.column).unwrap();
                let style = cell_display_style(m.row as u16, m.column, &cell_style);
                // The text is laid out over the whole merged cell, aligned
                // over its full width (even the part off screen) and wrapped
                // over its full height when the style says so. Every pane
                // shows the slice of lines and columns that falls inside it.
                let full_width: usize = (m.column..=m.last_column())
                    .map(|column| column_char_width(column) as usize)
                    .sum();
                let wraps = cell_style.alignment.as_ref().is_some_and(|a| a.wrap_text);
                let lines = if wraps {
                    wrap_text(&value, full_width)
                } else {
                    vec![value.clone()]
                };
                let line_start = |line: &str| -> usize {
                    let length = line.chars().count();
                    match cell_style.alignment.as_ref().map(|a| &a.horizontal) {
                        Some(HorizontalAlignment::Center)
                        | Some(HorizontalAlignment::CenterContinuous) => {
                            full_width.saturating_sub(length) / 2
                        }
                        Some(HorizontalAlignment::Right) => full_width.saturating_sub(length),
                        _ => 0,
                    }
                };
                let column_panes =
                    pane_extents(&column_tracks, m.column, m.last_column(), frozen_columns);
                let row_panes = pane_extents(&row_tracks, m.row, m.last_row(), frozen_rows as i32);
                for column_pane in &column_panes {
                    for row_pane in &row_panes {
                        let area = Rect {
                            x: column_pane.start,
                            y: row_pane.start,
                            width: column_pane.size,
                            height: row_pane.size,
                        }
                        .intersection(table_area);
                        if area.is_empty() {
                            continue;
                        }
                        rect.render_widget(Block::default().style(style), area);
                        // The pane shows the window of the merged cell that
                        // starts `offset` columns and `line_offset` lines
                        // into it
                        let offset: usize = (m.column..column_pane.first)
                            .map(|column| column_char_width(column) as usize)
                            .sum();
                        let line_offset: usize = (m.row..row_pane.first)
                            .map(|row| row_line_height(row as u16) as usize)
                            .sum();
                        for (line_index, line) in lines
                            .iter()
                            .enumerate()
                            .skip(line_offset)
                            .take(area.height as usize)
                        {
                            let text_start = line_start(line);
                            let from = text_start.max(offset);
                            let to = (text_start + line.chars().count())
                                .min(offset + area.width as usize);
                            if from >= to {
                                continue;
                            }
                            let text: String = line
                                .chars()
                                .skip(from - text_start)
                                .take(to - from)
                                .collect();
                            let text_area = Rect {
                                x: area.x + (from - offset) as u16,
                                y: area.y + (line_index - line_offset) as u16,
                                width: (to - from) as u16,
                                height: 1,
                            }
                            .intersection(area);
                            rect.render_widget(Paragraph::new(text).style(style), text_area);
                        }
                    }
                }
            }
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
                            format!(" {:>10}", key),
                            Style::default()
                                .fg(ironcalc_orange)
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
                    Event::Input(event) => {
                        let shift = event.modifiers.contains(KeyModifiers::SHIFT);
                        status_message = None;
                        // The edges of the merged cell containing a cell (or
                        // the cell itself): moving away from a merged cell
                        // starts past its far edge so one keystroke crosses
                        // the whole merged range.
                        let merge_edges = |row: u16, column: i32| -> (u16, u16, i32, i32) {
                            match merged_cell_containing(&merged_cells, row as i32, column) {
                                Some(m) => {
                                    (m.row as u16, m.last_row() as u16, m.column, m.last_column())
                                }
                                None => (row, row, column, column),
                            }
                        };
                        match event.code {
                            KeyCode::Char('q') => {
                                popup_open = true;
                                cursor_mode = CursorMode::Popup;
                            }
                            // Shift+arrows extend the selection: rows unless whole
                            // columns are selected, columns unless whole rows are.
                            KeyCode::Up if shift => {
                                let (top, _, _, _) = merge_edges(end_row, end_column);
                                if !whole_column && top > 1 {
                                    end_row = top - 1;
                                }
                            }
                            KeyCode::Down if shift => {
                                let (_, bottom, _, _) = merge_edges(end_row, end_column);
                                if !whole_column {
                                    end_row = bottom + 1;
                                }
                            }
                            KeyCode::Left if shift => {
                                let (_, _, left, _) = merge_edges(end_row, end_column);
                                if !whole_row && left > 1 {
                                    end_column = left - 1;
                                }
                            }
                            KeyCode::Right if shift => {
                                let (_, _, _, right) = merge_edges(end_row, end_column);
                                if !whole_row {
                                    end_column = right + 1;
                                }
                            }
                            // Going up past row 1 selects the whole column and
                            // going left past column A the whole row; with both,
                            // every cell is selected. Down/Right step back out.
                            // Landing inside a merged cell selects its anchor
                            // (see the top of the loop).
                            KeyCode::Down => {
                                let (_, bottom, _, _) =
                                    merge_edges(selected_row_index, selected_column_index);
                                if whole_column {
                                    whole_column = false;
                                    selected_row_index = 1;
                                } else {
                                    selected_row_index = bottom + 1;
                                }
                            }
                            KeyCode::Up => {
                                let (top, _, _, _) =
                                    merge_edges(selected_row_index, selected_column_index);
                                if top > 1 && !whole_column {
                                    selected_row_index = top - 1;
                                } else {
                                    whole_column = true;
                                }
                            }
                            KeyCode::Right => {
                                let (_, _, _, right) =
                                    merge_edges(selected_row_index, selected_column_index);
                                if whole_row {
                                    whole_row = false;
                                    selected_column_index = 1;
                                } else {
                                    selected_column_index = right + 1;
                                }
                            }
                            KeyCode::Left => {
                                let (_, _, left, _) =
                                    merge_edges(selected_row_index, selected_column_index);
                                if left > 1 && !whole_row {
                                    selected_column_index = left - 1;
                                } else {
                                    whole_row = true;
                                }
                            }
                            // Merge the selection (M also centers it), or
                            // unmerge when it already holds merged cells.
                            KeyCode::Char(c @ ('m' | 'M')) => {
                                let sheet = selected_sheet as u32;
                                let range = selected_area(
                                    sheet,
                                    (selected_row_index as i32, end_row as i32),
                                    (selected_column_index, end_column),
                                    whole_row,
                                    whole_column,
                                    &merged_cells,
                                );
                                let intersects_merged = merged_cells.iter().any(|m| {
                                    m.intersects(range.row, range.column, range.width, range.height)
                                });
                                if whole_row || whole_column {
                                    status_message =
                                        Some("Cannot merge whole rows or columns".to_string());
                                } else if intersects_merged {
                                    if let Err(message) = model.unmerge_cells(&range) {
                                        status_message = Some(message);
                                    }
                                } else if range.width * range.height < 2 {
                                    status_message =
                                        Some("Select more than one cell to merge".to_string());
                                } else {
                                    let result = if c == 'M' {
                                        model.merge_cells_center(&range)
                                    } else {
                                        model.merge_cells(&range)
                                    };
                                    match result {
                                        Ok(()) => {
                                            // Select the new merged cell: the
                                            // anchor with the focus at the far corner
                                            selected_row_index = range.row as u16;
                                            selected_column_index = range.column;
                                            end_row = (range.row + range.height - 1) as u16;
                                            end_column = range.column + range.width - 1;
                                        }
                                        Err(message) => status_message = Some(message),
                                    }
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
                                    let width =
                                        (width - COLUMN_WIDTH_FACTOR).max(COLUMN_WIDTH_FACTOR);
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
                                    let range = selected_area(
                                        sheet,
                                        (row, end_row as i32),
                                        (column, end_column),
                                        whole_row,
                                        whole_column,
                                        &merged_cells,
                                    );
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
                        }
                        // Moving without Shift collapses the selection
                        if !shift
                            && matches!(
                                event.code,
                                KeyCode::Up
                                    | KeyCode::Down
                                    | KeyCode::Left
                                    | KeyCode::Right
                                    | KeyCode::PageUp
                                    | KeyCode::PageDown
                            )
                        {
                            end_row = selected_row_index;
                            end_column = selected_column_index;
                        }
                    }
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
                        let range = selected_area(
                            selected_sheet as u32,
                            (selected_row_index as i32, end_row as i32),
                            (selected_column_index, end_column),
                            whole_row,
                            whole_column,
                            &merged_cells,
                        );
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

/// The selected range: the cells between `rows` and `columns` (each a pair of
/// opposite ends), widened to whole rows/columns or every cell
fn selection_area(
    sheet: u32,
    rows: (i32, i32),
    columns: (i32, i32),
    whole_row: bool,
    whole_column: bool,
) -> Area {
    let (column, width) = if whole_row {
        (1, LAST_COLUMN)
    } else {
        (columns.0.min(columns.1), (columns.0 - columns.1).abs() + 1)
    };
    let (row, height) = if whole_column {
        (1, LAST_ROW)
    } else {
        (rows.0.min(rows.1), (rows.0 - rows.1).abs() + 1)
    };
    Area {
        sheet,
        row,
        column,
        width,
        height,
    }
}

/// A visible row or column: its index and its position and size on screen
struct Track {
    index: i32,
    start: u16,
    size: u16,
}

/// The screen extent of a run of tracks: the index of its first track and
/// its position and size on screen
#[derive(Debug, PartialEq)]
struct Extent {
    first: i32,
    start: u16,
    size: u16,
}

/// The screen extents of the tracks with index in `first..=last`, one per
/// pane: the frozen tracks (index <= frozen) and the scrolled ones. Panes
/// with no such track are left out.
fn pane_extents(tracks: &[Track], first: i32, last: i32, frozen: i32) -> Vec<Extent> {
    let mut extents = Vec::new();
    for in_frozen_pane in [true, false] {
        let pane: Vec<&Track> = tracks
            .iter()
            .filter(|track| (track.index <= frozen) == in_frozen_pane)
            .filter(|track| track.index >= first && track.index <= last)
            .collect();
        if let (Some(head), Some(tail)) = (pane.first(), pane.last()) {
            extents.push(Extent {
                first: head.index,
                start: head.start,
                size: tail.start + tail.size - head.start,
            });
        }
    }
    extents
}

/// Word-wraps `text` to lines of at most `width` characters: lines break at
/// spaces (and at explicit newlines), words longer than a line are split.
fn wrap_text(text: &str, width: usize) -> Vec<String> {
    let width = width.max(1);
    let mut lines = Vec::new();
    for paragraph in text.lines() {
        let mut line = String::new();
        let mut line_length = 0;
        for word in paragraph.split(' ') {
            let mut word: Vec<char> = word.chars().collect();
            // A word that does not fit the rest of the line starts a new one
            if line_length > 0 && line_length + 1 + word.len() > width {
                lines.push(std::mem::take(&mut line));
                line_length = 0;
            }
            if line_length > 0 {
                line.push(' ');
                line_length += 1;
            }
            // Split the words longer than a line
            while word.len() > width - line_length {
                let rest = word.split_off(width - line_length);
                line.extend(&word);
                lines.push(std::mem::take(&mut line));
                line_length = 0;
                word = rest;
            }
            line_length += word.len();
            line.extend(&word);
        }
        lines.push(line);
    }
    lines
}

/// The merged cell that contains (row, column), if any
fn merged_cell_containing(
    merged_cells: &[MergedCell],
    row: i32,
    column: i32,
) -> Option<&MergedCell> {
    merged_cells.iter().find(|m| m.contains(row, column))
}

/// Grows `area` until it fully contains every merged cell it touches, so a
/// selection never covers a merged cell partially (growing over one merged
/// cell can graze another, hence the loop). Mirrors the engine's
/// `grow_range_over_merged_cells`.
fn grow_area_over_merged_cells(area: Area, merged_cells: &[MergedCell]) -> Area {
    let mut min_row = area.row;
    let mut max_row = area.row + area.height - 1;
    let mut min_column = area.column;
    let mut max_column = area.column + area.width - 1;
    loop {
        let mut changed = false;
        for m in merged_cells {
            if !m.intersects(
                min_row,
                min_column,
                max_column - min_column + 1,
                max_row - min_row + 1,
            ) {
                continue;
            }
            if m.row < min_row {
                min_row = m.row;
                changed = true;
            }
            if m.last_row() > max_row {
                max_row = m.last_row();
                changed = true;
            }
            if m.column < min_column {
                min_column = m.column;
                changed = true;
            }
            if m.last_column() > max_column {
                max_column = m.last_column();
                changed = true;
            }
        }
        if !changed {
            break;
        }
    }
    Area {
        sheet: area.sheet,
        row: min_row,
        column: min_column,
        width: max_column - min_column + 1,
        height: max_row - min_row + 1,
    }
}

/// The selected range as the user sees it: [selection_area] grown over the
/// merged cells it touches. Whole row/column selections are left alone: like
/// in the engine they may slice through merged cells.
fn selected_area(
    sheet: u32,
    rows: (i32, i32),
    columns: (i32, i32),
    whole_row: bool,
    whole_column: bool,
    merged_cells: &[MergedCell],
) -> Area {
    let area = selection_area(sheet, rows, columns, whole_row, whole_column);
    if whole_row || whole_column {
        area
    } else {
        grow_area_over_merged_cells(area, merged_cells)
    }
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selection_area_spans_both_ends() {
        let area = selection_area(0, (6, 3), (2, 2), false, false);
        assert_eq!(
            (area.row, area.height, area.column, area.width),
            (3, 4, 2, 1)
        );
        let area = selection_area(0, (1, 1), (5, 3), false, true);
        assert_eq!(
            (area.row, area.height, area.column, area.width),
            (1, LAST_ROW, 3, 3)
        );
        let area = selection_area(0, (2, 4), (1, 1), true, false);
        assert_eq!(
            (area.row, area.height, area.column, area.width),
            (2, 3, 1, LAST_COLUMN)
        );
    }

    fn merged(row: i32, column: i32, width: i32, height: i32) -> MergedCell {
        MergedCell {
            row,
            column,
            width,
            height,
        }
    }

    fn area(row: i32, column: i32, width: i32, height: i32) -> Area {
        Area {
            sheet: 0,
            row,
            column,
            width,
            height,
        }
    }

    #[test]
    fn selection_grows_over_merged_cells() {
        // B2:C3 merged and D3:D5 merged: touching B2 drags in C3, which
        // touches D3:D5, which drags in rows 4 and 5.
        let merged_cells = [merged(2, 2, 2, 2), merged(3, 4, 1, 3)];
        let grown = grow_area_over_merged_cells(area(2, 2, 1, 1), &merged_cells);
        assert_eq!(
            (grown.row, grown.column, grown.width, grown.height),
            (2, 2, 2, 2)
        );
        let grown = grow_area_over_merged_cells(area(3, 3, 1, 1), &merged_cells);
        assert_eq!(
            (grown.row, grown.column, grown.width, grown.height),
            (2, 2, 2, 2)
        );
        let grown = grow_area_over_merged_cells(area(3, 2, 3, 1), &merged_cells);
        assert_eq!(
            (grown.row, grown.column, grown.width, grown.height),
            (2, 2, 3, 4)
        );
        // Nothing merged in A1
        let grown = grow_area_over_merged_cells(area(1, 1, 1, 1), &merged_cells);
        assert_eq!(
            (grown.row, grown.column, grown.width, grown.height),
            (1, 1, 1, 1)
        );
    }

    #[test]
    fn whole_rows_and_columns_slice_merged_cells() {
        let merged_cells = [merged(2, 2, 2, 2)];
        let selected = selected_area(0, (2, 2), (1, 1), true, false, &merged_cells);
        assert_eq!((selected.row, selected.height), (2, 1));
        let selected = selected_area(0, (2, 2), (2, 2), false, false, &merged_cells);
        assert_eq!(
            (
                selected.row,
                selected.column,
                selected.width,
                selected.height
            ),
            (2, 2, 2, 2)
        );
        assert!(merged_cell_containing(&merged_cells, 3, 3).is_some());
        assert!(merged_cell_containing(&merged_cells, 4, 3).is_none());
    }

    #[test]
    fn pane_extents_split_at_the_frozen_pane() {
        // Columns A (frozen), then C and D visible after scrolling, with the
        // one-cell separator after the frozen pane.
        let tracks = [
            Track {
                index: 1,
                start: 3,
                size: 4,
            },
            Track {
                index: 3,
                start: 8,
                size: 5,
            },
            Track {
                index: 4,
                start: 13,
                size: 5,
            },
        ];
        let extent = |first, start, size| Extent { first, start, size };
        // A merged cell over A:D is painted in both panes
        assert_eq!(
            pane_extents(&tracks, 1, 4, 1),
            vec![extent(1, 3, 4), extent(3, 8, 10)]
        );
        // One over C:D only in the scrolled pane
        assert_eq!(pane_extents(&tracks, 3, 4, 1), vec![extent(3, 8, 10)]);
        // One over B:C is only partially visible
        assert_eq!(pane_extents(&tracks, 2, 3, 1), vec![extent(3, 8, 5)]);
        // Without frozen columns everything is one pane
        assert_eq!(pane_extents(&tracks, 1, 4, 0), vec![extent(1, 3, 15)]);
        assert!(pane_extents(&tracks, 5, 6, 0).is_empty());
    }

    #[test]
    fn wrap_text_breaks_at_spaces_and_long_words() {
        assert_eq!(
            wrap_text("Hello merged world", 10),
            ["Hello", "merged", "world"]
        );
        assert_eq!(
            wrap_text("Hello merged world", 12),
            ["Hello merged", "world"]
        );
        assert_eq!(wrap_text("Hello merged world", 30), ["Hello merged world"]);
        assert_eq!(wrap_text("abcdefghijkl", 5), ["abcde", "fghij", "kl"]);
        assert_eq!(wrap_text("ab cdefghijkl", 5), ["ab", "cdefg", "hijkl"]);
        assert_eq!(wrap_text("one\ntwo three", 5), ["one", "two", "three"]);
        assert!(wrap_text("", 5).is_empty());
        assert_eq!(wrap_text("a", 0), ["a"]);
    }
}
