use crossterm::{
    event::{self, DisableMouseCapture, Event as CEvent, KeyCode},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, LeaveAlternateScreen},
};
use ironcalc::{
    base::{expressions::utils::number_to_column, model::Model},
    import::load_model_from_xlsx,
};
use ratatui::{
    backend::CrosstermBackend,
    layout::{Constraint, Direction, Layout},
    style::{Color, Style},
    text::{Line, Span},
    widgets::{Block, BorderType, Borders, Cell, Paragraph, Row, Table},
    Terminal,
};
use std::io;
use std::sync::mpsc;
use std::thread;
use std::time::{Duration, Instant};

use std::env;

enum Event<I> {
    Input(I),
    Tick,
}

#[derive(PartialEq)]
enum CursorMode {
    Navigate,
    Input,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    enable_raw_mode().expect("cannot run in raw mode");
    let args: Vec<String> = env::args().collect();
    let mut model = if args.len() > 1 {
        let file_name = &args[1];
        load_model_from_xlsx(file_name, "en", "UTC").unwrap()
    } else {
        Model::new_empty("model.xlsx", "en", "UTC").unwrap()
    };
    let mut selected_sheet = 0;
    let mut selected_row_index = 1;
    let mut selected_column_index = 1;
    let mut minimum_row_index = 1;
    let mut minimum_column_index = 1;
    let sheet_list_width = 20;
    let column_width: u16 = 11;
    let mut cursor_mode = CursorMode::Navigate;
    let mut input_str = String::new();

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

    let stdout = io::stdout();
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;
    terminal.clear()?;

    let header_style = Style::default().fg(Color::Yellow).bg(Color::White);
    let selected_header_style = Style::default().bg(Color::Yellow).fg(Color::White);

    let cell_style = Style::default().fg(Color::White).bg(Color::Black);
    let selected_cell_style = Style::default().fg(Color::Yellow).bg(Color::White);

    let background_style = Style::default().bg(Color::Black);
    let selected_sheet_style = Style::default().bg(Color::White).fg(Color::LightMagenta);
    let non_selected_sheet_style = Style::default().fg(Color::White);
    let mut sheet_names = model.workbook.get_worksheet_names();
    loop {
        terminal.draw(|rect| {
            let size = rect.size();

            let global_chunks = Layout::default()
                .direction(Direction::Horizontal)
                .constraints([Constraint::Length(sheet_list_width), Constraint::Min(3)].as_ref())
                .split(size);

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
            let sheet_list = Table::new(rows, widths)
                .block(sheets)
                .column_spacing(0);

            rect.render_widget(sheet_list, global_chunks[0]);

            // The spreadsheet is the formula bar at the top and the sheet data
            let spreadsheet_chunks = Layout::default()
                .direction(Direction::Vertical)
                .margin(0)
                .constraints([Constraint::Length(1), Constraint::Min(2)].as_ref())
                .split(global_chunks[1]);

            let spreadsheet_width = size.width - sheet_list_width;
            let spreadsheet_heigh = size.height - 1;
            let row_count = spreadsheet_heigh - 1;

            let first_row_width: u16 = 3;
            let column_count =
                f64::ceil(((spreadsheet_width - first_row_width) as f64) / (column_width as f64))
                    as i32;
            let mut rows = vec![];
            // The first row in the column headers
            let mut row = Vec::new();
            // The first cell in that row is the top left square of the spreadsheet
            row.push(Cell::from(""));
            let mut maximum_column_index = minimum_column_index + column_count - 1;
            let mut maximum_row_index = minimum_row_index + row_count - 1;

            // We want to make sure the selected cell is visible.
            if selected_column_index > maximum_column_index {
                maximum_column_index = selected_column_index;
                minimum_column_index = maximum_column_index - column_count + 1;
            } else if selected_column_index < minimum_column_index {
                minimum_column_index = selected_column_index;
                maximum_column_index = minimum_column_index + column_count - 1;
            }
            if selected_row_index >= maximum_row_index {
                maximum_row_index = selected_row_index;
                minimum_row_index = maximum_row_index - row_count + 1;
            } else if selected_row_index < minimum_row_index {
                minimum_row_index = selected_row_index;
                maximum_row_index = minimum_row_index + row_count - 1;
            }
            for column_index in minimum_column_index..=maximum_column_index {
                let column_str = number_to_column(column_index);
                let style = if column_index == selected_column_index {
                    selected_header_style
                } else {
                    header_style
                };
                row.push(Cell::from(format!("     {}", column_str.unwrap())).style(style));
            }
            rows.push(Row::new(row));
            for row_index in minimum_row_index..=maximum_row_index {
                let mut row = Vec::new();
                let style = if row_index == selected_row_index {
                    selected_header_style
                } else {
                    header_style
                };
                row.push(Cell::from(format!("{}", row_index)).style(style));
                for column_index in minimum_column_index..=maximum_column_index {
                    let value = model
                        .formatted_cell_value(selected_sheet as u32, row_index as i32, column_index)
                        .unwrap();
                    let style = if selected_row_index == row_index
                        && selected_column_index == column_index
                    {
                        selected_cell_style
                    } else {
                        cell_style
                    };
                    row.push(Cell::from(value.to_string()).style(style));
                }
                rows.push(Row::new(row));
            }
            let mut widths = Vec::new();
            widths.push(Constraint::Length(first_row_width));
            for _ in 0..column_count {
                widths.push(Constraint::Length(column_width));
            }
            let spreadsheet = Table::new(rows, widths)
                .block(Block::default().style(Style::default().bg(Color::Black)))
                .column_spacing(0);

            let text = if cursor_mode == CursorMode::Navigate {
                model
                    .cell_formula(
                        selected_sheet as u32,
                        selected_row_index as i32,
                        selected_column_index,
                    )
                    .unwrap()
                    .unwrap_or_else(|| {
                        model
                            .formatted_cell_value(
                                selected_sheet as u32,
                                selected_row_index as i32,
                                selected_column_index,
                            )
                            .unwrap()
                    })
            } else {
                format!("{}|", input_str)
            };

            let formula_bar_text = format!(
                "{}{}: {}",
                number_to_column(selected_column_index).unwrap(),
                selected_row_index,
                text
            );
            let formula_bar = Paragraph::new(vec![Line::from(vec![Span::raw(formula_bar_text)])]);
            rect.render_widget(formula_bar.block(Block::default()), spreadsheet_chunks[0]);
            rect.render_widget(spreadsheet, spreadsheet_chunks[1]);
        })?;

        match cursor_mode {
            CursorMode::Navigate => {
                match rx.recv()? {
                    Event::Input(event) => match event.code {
                        KeyCode::Char('q') => {
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
                            input_str = model
                                .cell_formula(
                                    selected_sheet as u32,
                                    selected_row_index as i32,
                                    selected_column_index,
                                )
                                .unwrap()
                                .unwrap_or_default()
                        }
                        KeyCode::Char('+') => {
                            model.new_sheet();
                            model.evaluate();
                            sheet_names = model.workbook.get_worksheet_names();
                        }
                        _ => {
                            // println!("{:?}", event);
                        }
                    },
                    Event::Tick => {}
                }
            }
            CursorMode::Input => match rx.recv()? {
                Event::Input(event) => match event.code {
                    KeyCode::Char(c) => {
                        input_str.push(c);
                    }
                    KeyCode::Backspace => {
                        input_str.pop();
                    }
                    KeyCode::Enter => {
                        cursor_mode = CursorMode::Navigate;
                        let value = input_str.clone();
                        let sheet = selected_sheet as i32;
                        let row = selected_row_index as i32;
                        let column = selected_column_index;
                        model.set_user_input(sheet as u32, row, column, value);
                        model.evaluate();
                    }
                    _ => {}
                },
                Event::Tick => {}
            },
        }
    }

    Ok(())
}
