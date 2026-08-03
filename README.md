# TironCalc

[![Discord chat][discord-badge]][discord-url]

[discord-badge]: https://img.shields.io/discord/1206947691058171904.svg?logo=discord&style=flat-square
[discord-url]: https://discord.gg/zZYWfh3RHJ

TironCalc, or Tiron for friends, is a TUI (Terminal User Interface) for IronCalc. Based on [ratatui](https://github.com/ratatui-org/ratatui)

![TironCalc Screenshot](docs/screenshot.png)

## Build

```
cargo build --release
```

You will find the binary at `./target/release/tiron`.

## Documentation

Start empty project:

```
$ tiron
```

Load an existing Excel file:

```
$ tiron example.xlsx
```
-   `Arrow Keys` to navigate cells
-   `e` to edit a cell and enter the value or formula.
-   `q` to quit and save
-   `+` to add a sheet
-   `s` to go to the next sheet
-   `a` to go to the previous sheet
-   `b` to pick a background color for the current cell
-   `c` to pick a text color for the current cell
-   `B`/`I`/`U`/`S` to toggle bold/italic/underline/strikethrough
-   `<`/`>` (or `,`/`.`) to make the current column narrower/wider
-   `-`/`=` to make the current row shorter/taller
-   `f` to freeze the rows above and the columns to the left of the current
    cell; press it again on the same cell (or on A1) to unfreeze
-   `?` to show the keyboard shortcuts
-   `PgUp/PgDown` to navigate rows faster


## Inspiration

James Gosling of Java fame created [sc](https://en.wikipedia.org/wiki/Sc_(spreadsheet_calculator)) the spreadsheet calculator.

Andrés Martinelli has been maintaining [sc-im](https://github.com/andmarti1424/sc-im), the spreadsheet calculator improvised.

## See also

A really nice similar project

[SheetsUI](https://github.com/zaphar/sheetsui)
