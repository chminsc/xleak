use anyhow::{Context, Result, anyhow};
use calamine::{Data, Range, Reader, Sheets, Table, open_workbook_auto};
use chrono::{Duration, NaiveDate};
use std::fs;
use std::path::Path;

pub struct Workbook {
    source: WorkbookSource,
}

enum WorkbookSource {
    Excel(Sheets<std::io::BufReader<std::fs::File>>),
    Csv { sheet_name: String, content: String },
}

impl Workbook {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();

        if path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("csv"))
        {
            let content = fs::read_to_string(path).context("Failed to read CSV file")?;
            let sheet_name = path
                .file_stem()
                .and_then(|name| name.to_str())
                .filter(|name| !name.is_empty())
                .unwrap_or("CSV")
                .to_string();

            return Ok(Self {
                source: WorkbookSource::Csv {
                    sheet_name,
                    content,
                },
            });
        }

        let sheets = open_workbook_auto(path).context("Failed to open workbook")?;

        Ok(Self {
            source: WorkbookSource::Excel(sheets),
        })
    }

    pub fn sheet_names(&self) -> Vec<String> {
        match &self.source {
            WorkbookSource::Excel(sheets) => sheets.sheet_names(),
            WorkbookSource::Csv { sheet_name, .. } => vec![sheet_name.clone()],
        }
    }

    /// Loads all rows eagerly into memory
    pub fn load_sheet(&mut self, name: &str) -> Result<SheetData> {
        match &mut self.source {
            WorkbookSource::Excel(sheets) => {
                let range = sheets
                    .worksheet_range(name)
                    .with_context(|| format!("Sheet '{name}' not found"))?;

                // Try to load formulas, but don't fail if they're not available
                let formula_range = sheets.worksheet_formula(name).ok();

                Ok(SheetData::from_range_with_formulas(range, formula_range))
            }
            WorkbookSource::Csv {
                sheet_name,
                content,
            } => {
                ensure_csv_sheet_name(name, sheet_name)?;
                Ok(SheetData::from_range_with_formulas(
                    csv_content_to_range(content)?,
                    None,
                ))
            }
        }
    }

    /// Loads only headers; rows fetched on demand
    pub fn load_sheet_lazy(&mut self, name: &str) -> Result<LazySheetData> {
        match &mut self.source {
            WorkbookSource::Excel(sheets) => {
                let range = sheets
                    .worksheet_range(name)
                    .with_context(|| format!("Sheet '{name}' not found"))?;

                // Try to load formulas, but don't fail if they're not available
                let formula_range = sheets.worksheet_formula(name).ok();

                Ok(LazySheetData::from_range_with_formulas(
                    range,
                    formula_range,
                ))
            }
            WorkbookSource::Csv {
                sheet_name,
                content,
            } => {
                ensure_csv_sheet_name(name, sheet_name)?;
                Ok(LazySheetData::from_range_with_formulas(
                    csv_content_to_range(content)?,
                    None,
                ))
            }
        }
    }

    // ===== Table API (Xlsx only) =====

    /// Load table metadata from the workbook (Xlsx only)
    pub fn load_tables(&mut self) -> Result<()> {
        match &mut self.source {
            WorkbookSource::Excel(Sheets::Xlsx(xlsx)) => xlsx
                .load_tables()
                .context("Failed to load table metadata")
                .map_err(|e| anyhow!("{e}")),
            WorkbookSource::Excel(_) | WorkbookSource::Csv { .. } => {
                Err(anyhow!("Tables are only supported in .xlsx files"))
            }
        }
    }

    /// Get all table names in the workbook (Xlsx only)
    pub fn table_names(&self) -> Result<Vec<String>> {
        match &self.source {
            WorkbookSource::Excel(Sheets::Xlsx(xlsx)) => {
                Ok(xlsx.table_names().iter().map(|s| (*s).clone()).collect())
            }
            WorkbookSource::Excel(_) | WorkbookSource::Csv { .. } => {
                Err(anyhow!("Tables are only supported in .xlsx files"))
            }
        }
    }

    /// Get table names in a specific sheet (Xlsx only)
    pub fn table_names_in_sheet(&self, sheet_name: &str) -> Result<Vec<String>> {
        match &self.source {
            WorkbookSource::Excel(Sheets::Xlsx(xlsx)) => Ok(xlsx
                .table_names_in_sheet(sheet_name)
                .iter()
                .map(|s| (*s).clone())
                .collect()),
            WorkbookSource::Excel(_) | WorkbookSource::Csv { .. } => {
                Err(anyhow!("Tables are only supported in .xlsx files"))
            }
        }
    }

    /// Get table data by name (Xlsx only)
    pub fn table_by_name(&mut self, table_name: &str) -> Result<TableData> {
        match &mut self.source {
            WorkbookSource::Excel(Sheets::Xlsx(xlsx)) => {
                let table = xlsx
                    .table_by_name(table_name)
                    .map_err(|e| anyhow!("Table '{table_name}' not found: {e}"))?;

                Ok(TableData::from_calamine_table(table))
            }
            WorkbookSource::Excel(_) | WorkbookSource::Csv { .. } => {
                Err(anyhow!("Tables are only supported in .xlsx files"))
            }
        }
    }
}

fn ensure_csv_sheet_name(requested: &str, sheet_name: &str) -> Result<()> {
    if requested == sheet_name {
        Ok(())
    } else {
        Err(anyhow!(
            "Sheet '{requested}' not found. Available sheets: {sheet_name}"
        ))
    }
}

fn csv_content_to_range(content: &str) -> Result<Range<Data>> {
    let records = parse_csv_records(content)?;

    if records.is_empty() {
        return Ok(Range::empty());
    }

    let width = records.iter().map(Vec::len).max().unwrap_or(0);
    if width == 0 {
        return Ok(Range::empty());
    }

    let mut range = Range::new((0, 0), ((records.len() - 1) as u32, (width - 1) as u32));

    for (row_idx, row) in records.into_iter().enumerate() {
        for (col_idx, value) in row.into_iter().enumerate() {
            range[(row_idx, col_idx)] = if value.is_empty() {
                Data::Empty
            } else {
                Data::String(value)
            };
        }
    }

    Ok(range)
}

fn parse_csv_records(content: &str) -> Result<Vec<Vec<String>>> {
    if content.is_empty() {
        return Ok(vec![]);
    }

    let mut rows = Vec::new();
    let mut row = Vec::new();
    let mut field = String::new();
    let mut chars = content.chars().peekable();
    let mut in_quotes = false;

    while let Some(ch) = chars.next() {
        if in_quotes {
            match ch {
                '"' => {
                    if chars.peek() == Some(&'"') {
                        field.push('"');
                        chars.next();
                    } else {
                        in_quotes = false;
                    }
                }
                _ => field.push(ch),
            }
            continue;
        }

        match ch {
            '"' if field.is_empty() => in_quotes = true,
            ',' => {
                row.push(std::mem::take(&mut field));
            }
            '\n' => {
                row.push(std::mem::take(&mut field));
                rows.push(std::mem::take(&mut row));
            }
            '\r' => {
                if chars.peek() == Some(&'\n') {
                    chars.next();
                }
                row.push(std::mem::take(&mut field));
                rows.push(std::mem::take(&mut row));
            }
            _ => field.push(ch),
        }
    }

    if in_quotes {
        return Err(anyhow!("Invalid CSV: unterminated quoted field"));
    }

    if !field.is_empty() || !row.is_empty() || !(content.ends_with('\n') || content.ends_with('\r'))
    {
        row.push(field);
        rows.push(row);
    }

    if let Some(first_row) = rows.first_mut()
        && let Some(first_field) = first_row.first_mut()
    {
        *first_field = first_field.trim_start_matches('\u{feff}').to_string();
    }

    Ok(rows)
}

/// Eagerly-loaded sheet data (loads all rows immediately)
#[derive(Debug, Clone)]
pub struct SheetData {
    pub headers: Vec<String>,
    pub rows: Vec<Vec<CellValue>>,
    pub formulas: Vec<Vec<Option<String>>>, // Parallel structure to rows with formulas
    pub width: usize,
    pub height: usize,
}

/// Lazy-loaded sheet data (loads rows on demand)
pub struct LazySheetData {
    range: Range<Data>,
    formula_range: Option<Range<String>>,
    pub headers: Vec<String>,
    pub width: usize,
    pub height: usize,
}

impl LazySheetData {
    /// Extracts headers only; defers row loading
    pub fn from_range_with_formulas(
        range: Range<Data>,
        formula_range: Option<Range<String>>,
    ) -> Self {
        let (height, width) = range.get_size();

        // Only extract headers (first row) - don't load all data yet
        let headers = if height > 0 {
            range
                .rows()
                .next()
                .map(|row| row.iter().map(SheetData::cell_to_string).collect())
                .unwrap_or_default()
        } else {
            vec![]
        };

        Self {
            range,
            formula_range,
            headers,
            width,
            height: height.saturating_sub(1), // Don't count header row
        }
    }

    /// Zero-indexed row range; header excluded
    pub fn get_rows(
        &self,
        start: usize,
        count: usize,
    ) -> (Vec<Vec<CellValue>>, Vec<Vec<Option<String>>>) {
        let end = (start + count).min(self.height);

        // Extract requested rows (skip header + start rows, take count)
        let rows: Vec<Vec<CellValue>> = self
            .range
            .rows()
            .skip(1 + start) // Skip header + start offset
            .take(end - start)
            .map(|row| row.iter().map(SheetData::datatype_to_cellvalue).collect())
            .collect();

        // Extract formulas for requested rows
        let formulas = self.get_formulas_for_range(start, end);

        (rows, formulas)
    }

    fn get_formulas_for_range(&self, start: usize, end: usize) -> Vec<Vec<Option<String>>> {
        if let Some(ref formula_range) = self.formula_range {
            let formula_start = formula_range.start().unwrap_or((0, 0));
            let total_height = self.height + 1; // Include header in total

            // Create formula grid only for requested rows
            let mut formula_grid: Vec<Vec<Option<String>>> =
                vec![vec![None; self.width]; end - start];

            // Populate formulas at their absolute positions
            for (row_offset, formula_row) in formula_range.rows().enumerate() {
                let absolute_row = formula_start.0 as usize + row_offset;

                if absolute_row > 0 && absolute_row <= total_height {
                    let data_row_idx = absolute_row - 1; // Convert to 0-based data row index

                    // Only process if this row is in our requested range
                    if data_row_idx >= start && data_row_idx < end {
                        let result_idx = data_row_idx - start; // Index in result array

                        for (col_offset, formula_str) in formula_row.iter().enumerate() {
                            let absolute_col = formula_start.1 as usize + col_offset;
                            if absolute_col < self.width && !formula_str.is_empty() {
                                formula_grid[result_idx][absolute_col] = Some(formula_str.clone());
                            }
                        }
                    }
                }
            }

            formula_grid
        } else {
            // No formulas available
            vec![vec![None; self.width]; end - start]
        }
    }

    /// Consumes lazy data and loads all rows into memory
    #[allow(clippy::wrong_self_convention)]
    pub fn to_sheet_data(self) -> SheetData {
        SheetData::from_range_with_formulas(self.range, self.formula_range)
    }
}

#[derive(Debug, Clone)]
pub enum CellValue {
    Empty,
    String(String),
    Int(i64),
    Float(f64),
    Bool(bool),
    Error(String),
    DateTime(f64), // Excel datetime as float
}

impl CellValue {
    #[allow(dead_code)]
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    #[allow(dead_code)]
    pub fn is_numeric(&self) -> bool {
        matches!(self, CellValue::Int(_) | CellValue::Float(_))
    }

    /// Returns unformatted value (for export/clipboard)
    pub fn to_raw_string(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::String(s) => s.clone(),
            CellValue::Int(i) => i.to_string(),
            CellValue::Float(val) => {
                if val.fract() == 0.0 {
                    format!("{val:.0}")
                } else {
                    val.to_string()
                }
            }
            CellValue::Bool(b) => b.to_string(),
            CellValue::Error(e) => format!("#{e}"),
            CellValue::DateTime(dt) => {
                let days = dt.floor() as i64;
                // Excel epoch: December 31, 1899 (Excel serial 0)
                let epoch = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
                // Adjust for Excel's 1900 leap year bug (day 60 = Feb 29, 1900 which didn't exist)
                let adjusted_days = if days > 60 { days - 1 } else { days };
                let date = epoch + Duration::days(adjusted_days);
                let time_fraction = dt.fract();
                let total_seconds = (time_fraction * 86400.0).round() as i64;
                let hours = total_seconds / 3600;
                let minutes = (total_seconds % 3600) / 60;
                let seconds = total_seconds % 60;

                if time_fraction.abs() < 0.0000001 {
                    format!("{}", date.format("%Y-%m-%d"))
                } else {
                    format!(
                        "{} {:02}:{:02}:{:02}",
                        date.format("%Y-%m-%d"),
                        hours,
                        minutes,
                        seconds
                    )
                }
            }
        }
    }
}

pub fn row_uses_percentage_display(row: &[CellValue], marker_columns: &[usize]) -> bool {
    marker_columns.iter().any(|&column_number| {
        column_number
            .checked_sub(1)
            .and_then(|col_idx| row.get(col_idx))
            .map(|cell| cell.to_string().contains('%'))
            .unwrap_or(false)
    })
}

pub fn format_cell_for_display(
    cell: &CellValue,
    row: &[CellValue],
    marker_columns: &[usize],
) -> String {
    if row_uses_percentage_display(row, marker_columns) {
        match cell {
            CellValue::Int(i) => return format!("{:.1}%", (*i as f64) * 100.0),
            CellValue::Float(f) => return format!("{:.1}%", f * 100.0),
            _ => {}
        }
    }

    cell.to_string()
}

/// Excel Table data
#[derive(Debug, Clone)]
pub struct TableData {
    pub name: String,
    pub sheet_name: String,
    pub headers: Vec<String>,
    pub rows: Vec<Vec<CellValue>>,
}

impl TableData {
    pub fn from_calamine_table(table: Table<Data>) -> Self {
        let name = table.name().to_string();
        let sheet_name = table.sheet_name().to_string();
        let headers = table.columns().to_vec();

        let rows: Vec<Vec<CellValue>> = table
            .data()
            .rows()
            .map(|row| row.iter().map(SheetData::datatype_to_cellvalue).collect())
            .collect();

        Self {
            name,
            sheet_name,
            headers,
            rows,
        }
    }
}

impl std::fmt::Display for CellValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CellValue::Empty => write!(f, ""),
            CellValue::String(s) => write!(f, "{s}"),
            CellValue::Int(i) => {
                // Format integers with thousand separators
                let s = i.to_string();
                let negative = s.starts_with('-');
                let digits: String = s.trim_start_matches('-').chars().collect();
                let mut result = String::new();
                for (idx, ch) in digits.chars().rev().enumerate() {
                    if idx > 0 && idx % 3 == 0 {
                        result.push(',');
                    }
                    result.push(ch);
                }
                if negative {
                    result.push('-');
                }
                write!(f, "{}", result.chars().rev().collect::<String>())
            }
            CellValue::Float(val) => {
                // Format floats with thousand separators
                let formatted = if val.fract() == 0.0 {
                    format!("{val:.0}")
                } else {
                    format!("{val:.2}")
                };
                let parts: Vec<&str> = formatted.split('.').collect();
                let int_part = parts[0];
                let negative = int_part.starts_with('-');
                let digits: String = int_part.trim_start_matches('-').chars().collect();
                let mut result = String::new();
                for (idx, ch) in digits.chars().rev().enumerate() {
                    if idx > 0 && idx % 3 == 0 {
                        result.push(',');
                    }
                    result.push(ch);
                }
                if negative {
                    result.push('-');
                }
                let int_formatted: String = result.chars().rev().collect();
                if parts.len() > 1 {
                    write!(f, "{}.{}", int_formatted, parts[1])
                } else {
                    write!(f, "{}", int_formatted)
                }
            }
            CellValue::Bool(b) => {
                // Use lowercase for booleans
                write!(f, "{}", if *b { "true" } else { "false" })
            }
            CellValue::Error(e) => write!(f, "ERROR: {e}"),
            CellValue::DateTime(d) => {
                // Excel dates are days since December 31, 1899 (serial 0)
                // Excel has a leap year bug where 1900 is incorrectly treated as a leap year
                // Days > 60 need adjustment for this bug
                let days = d.floor() as i64;

                // Excel epoch: December 31, 1899 (Excel serial 0)
                let excel_epoch = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();

                // Adjust for Excel's 1900 leap year bug (day 60 = Feb 29, 1900 which didn't exist)
                let adjusted_days = if days > 60 { days - 1 } else { days };

                if let Some(date) = excel_epoch.checked_add_signed(Duration::days(adjusted_days)) {
                    // Check if there's a time component
                    let frac = d.fract();
                    if frac.abs() > 0.000001 {
                        // Has time component
                        let total_seconds = (frac * 86400.0).round() as u32;
                        let hours = total_seconds / 3600;
                        let minutes = (total_seconds % 3600) / 60;
                        let seconds = total_seconds % 60;
                        write!(f, "{} {:02}:{:02}:{:02}", date, hours, minutes, seconds)
                    } else {
                        // Date only
                        write!(f, "{}", date)
                    }
                } else {
                    write!(f, "Date[{days}]")
                }
            }
        }
    }
}

impl SheetData {
    pub fn from_range_with_formulas(
        range: Range<Data>,
        formula_range: Option<Range<String>>,
    ) -> Self {
        let (height, width) = range.get_size();

        // Extract headers from first row if it exists
        let headers = if height > 0 {
            range
                .rows()
                .next()
                .map(|row| row.iter().map(Self::cell_to_string).collect())
                .unwrap_or_default()
        } else {
            vec![]
        };

        // Extract data rows (skip first row as headers)
        let rows: Vec<Vec<CellValue>> = range
            .rows()
            .skip(1)
            .map(|row| row.iter().map(Self::datatype_to_cellvalue).collect())
            .collect();

        // Extract formulas if available
        // Note: Formula range may be sparse (only cells with formulas) and may have different start position
        let formulas: Vec<Vec<Option<String>>> = if let Some(formula_range) = formula_range {
            let formula_start = formula_range.start().unwrap_or((0, 0));

            // Create empty formula structure matching data dimensions
            let mut formula_grid: Vec<Vec<Option<String>>> = vec![vec![None; width]; height];

            // Populate formulas at their absolute positions
            for (row_offset, formula_row) in formula_range.rows().enumerate() {
                let absolute_row = formula_start.0 as usize + row_offset;
                if absolute_row > 0 && absolute_row <= height {
                    // Skip header row (row 0)
                    let data_row_idx = absolute_row - 1; // Convert to 0-based data row index
                    for (col_offset, formula_str) in formula_row.iter().enumerate() {
                        let absolute_col = formula_start.1 as usize + col_offset;
                        if absolute_col < width && !formula_str.is_empty() {
                            formula_grid[data_row_idx][absolute_col] = Some(formula_str.clone());
                        }
                    }
                }
            }

            // Return formula grid matching data rows
            // We already handled header row when populating, so just take the data rows
            formula_grid
                .into_iter()
                .take(height.saturating_sub(1))
                .collect()
        } else {
            // No formulas available, create empty parallel structure
            vec![vec![None; width]; height.saturating_sub(1)]
        };

        Self {
            headers,
            rows,
            formulas,
            width,
            height: height.saturating_sub(1), // Don't count header row
        }
    }

    fn cell_to_string(cell: &Data) -> String {
        match cell {
            Data::Empty => String::new(),
            Data::String(s) => s.clone(),
            Data::Int(i) => i.to_string(),
            Data::Float(f) => {
                if f.fract() == 0.0 {
                    format!("{f:.0}")
                } else {
                    f.to_string()
                }
            }
            Data::Bool(b) => b.to_string(),
            Data::Error(e) => format!("ERROR: {e:?}"),
            Data::DateTime(d) => format!("Date({})", d.as_f64()),
            Data::DateTimeIso(s) => s.clone(),
            Data::DurationIso(s) => s.clone(),
        }
    }

    fn datatype_to_cellvalue(cell: &Data) -> CellValue {
        match cell {
            Data::Empty => CellValue::Empty,
            Data::String(s) => CellValue::String(s.clone()),
            Data::Int(i) => CellValue::Int(*i),
            Data::Float(f) => CellValue::Float(*f),
            Data::Bool(b) => CellValue::Bool(*b),
            Data::Error(e) => CellValue::Error(format!("{e:?}")),
            Data::DateTime(d) => CellValue::DateTime(d.as_f64()),
            Data::DateTimeIso(s) => CellValue::String(s.clone()),
            Data::DurationIso(s) => CellValue::String(s.clone()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cellvalue_display_integer() {
        let val = CellValue::Int(1234567);
        assert_eq!(val.to_string(), "1,234,567");
    }

    #[test]
    fn test_cellvalue_display_negative_integer() {
        let val = CellValue::Int(-1234567);
        assert_eq!(val.to_string(), "-1,234,567");
    }

    #[test]
    fn test_cellvalue_display_float() {
        let val = CellValue::Float(1234567.89);
        assert_eq!(val.to_string(), "1,234,567.89");
    }

    #[test]
    fn test_row_uses_percentage_display() {
        let row = vec![
            CellValue::String("Gross margin %".to_string()),
            CellValue::Empty,
            CellValue::Float(0.222),
        ];
        assert!(row_uses_percentage_display(&row, &[1, 2]));
        assert!(!row_uses_percentage_display(&row, &[2]));
    }

    #[test]
    fn test_format_cell_for_display_percentage_row() {
        let row = vec![
            CellValue::String("Margin".to_string()),
            CellValue::String("%".to_string()),
            CellValue::Float(0.222),
        ];
        assert_eq!(format_cell_for_display(&row[2], &row, &[1, 2]), "22.2%");
    }

    #[test]
    fn test_cellvalue_display_float_whole_number() {
        let val = CellValue::Float(1000.0);
        assert_eq!(val.to_string(), "1,000");
    }

    #[test]
    fn test_cellvalue_display_boolean() {
        assert_eq!(CellValue::Bool(true).to_string(), "true");
        assert_eq!(CellValue::Bool(false).to_string(), "false");
    }

    #[test]
    fn test_cellvalue_display_string() {
        let val = CellValue::String("Hello, World!".to_string());
        assert_eq!(val.to_string(), "Hello, World!");
    }

    #[test]
    fn test_cellvalue_display_empty() {
        let val = CellValue::Empty;
        assert_eq!(val.to_string(), "");
    }

    #[test]
    fn test_cellvalue_display_error() {
        let val = CellValue::Error("DIV/0!".to_string());
        assert_eq!(val.to_string(), "ERROR: DIV/0!");
    }

    #[test]
    fn test_cellvalue_to_raw_string_integer() {
        let val = CellValue::Int(1234567);
        assert_eq!(val.to_raw_string(), "1234567");
    }

    #[test]
    fn test_cellvalue_to_raw_string_float() {
        let val = CellValue::Float(123.45);
        assert_eq!(val.to_raw_string(), "123.45");
    }

    #[test]
    fn test_cellvalue_is_empty() {
        assert!(CellValue::Empty.is_empty());
        assert!(!CellValue::Int(0).is_empty());
        assert!(!CellValue::String("".to_string()).is_empty());
    }

    #[test]
    fn test_cellvalue_is_numeric() {
        assert!(CellValue::Int(123).is_numeric());
        assert!(CellValue::Float(123.45).is_numeric());
        assert!(!CellValue::String("123".to_string()).is_numeric());
        assert!(!CellValue::Empty.is_numeric());
    }

    #[test]
    fn test_datetime_display() {
        // Excel date: January 1, 1900 is day 1
        let val = CellValue::DateTime(1.0);
        let display = val.to_string();
        // Should contain a date in YYYY-MM-DD format
        assert!(display.contains("1900") || display.contains("1899"));
    }

    #[test]
    fn test_datetime_with_time() {
        // Excel datetime with time component
        // Day 1 + 0.5 = 12:00:00 on Jan 1, 1900
        let val = CellValue::DateTime(1.5);
        let display = val.to_string();
        // Should contain both date and time
        assert!(display.contains(":"));
        assert!(display.len() > 10); // Date + time is longer than just date
    }

    #[test]
    fn test_workbook_open_real_file() {
        // Test with actual test file if it exists
        if let Ok(wb) = Workbook::open("tests/fixtures/test_data.xlsx") {
            let sheet_names = wb.sheet_names();
            assert!(!sheet_names.is_empty(), "Should have at least one sheet");
        }
        // If file doesn't exist, test passes (integration test needs real file)
    }

    #[test]
    fn test_parse_csv_records_handles_quotes_and_newlines() {
        let rows = parse_csv_records("Name,Note\nAlice,\"Hello, CSV\"\nBob,\"Line 1\nLine 2\"\n")
            .expect("CSV should parse");

        assert_eq!(rows.len(), 3);
        assert_eq!(rows[1], vec!["Alice", "Hello, CSV"]);
        assert_eq!(rows[2], vec!["Bob", "Line 1\nLine 2"]);
    }

    #[test]
    fn test_workbook_open_csv_file() {
        let sheet_name = format!("xleak_test_csv_{}", std::process::id());
        let path = std::env::temp_dir().join(format!("{sheet_name}.csv"));
        fs::write(
            &path,
            "Name,Age,Note\nAlice,30,\"Hello, CSV\"\nBob,25,\"Line 1\nLine 2\"\n\"Charlie \"\"CJ\"\"\",,\"Uses quotes\"",
        )
        .expect("test CSV should be written");

        let mut wb = Workbook::open(&path).expect("CSV should open");
        assert_eq!(wb.sheet_names(), vec![sheet_name.clone()]);

        let data = wb.load_sheet(&sheet_name).expect("CSV sheet should load");

        assert_eq!(data.headers, vec!["Name", "Age", "Note"]);
        assert_eq!(data.width, 3);
        assert_eq!(data.height, 3);
        assert_eq!(data.rows[0][2].to_raw_string(), "Hello, CSV");
        assert_eq!(data.rows[1][2].to_raw_string(), "Line 1\nLine 2");
        assert_eq!(data.rows[2][0].to_raw_string(), "Charlie \"CJ\"");
        assert!(matches!(data.rows[2][1], CellValue::Empty));

        fs::remove_file(path).expect("test CSV should be removed");
    }

    #[test]
    fn test_sheet_data_structure() {
        // Test SheetData structure can be created
        let sheet = SheetData {
            headers: vec!["Name".to_string(), "Age".to_string()],
            rows: vec![
                vec![CellValue::String("Alice".to_string()), CellValue::Int(30)],
                vec![CellValue::String("Bob".to_string()), CellValue::Int(25)],
            ],
            formulas: vec![vec![None, None], vec![None, None]],
            width: 2,
            height: 2,
        };

        assert_eq!(sheet.width, 2);
        assert_eq!(sheet.height, 2);
        assert_eq!(sheet.headers.len(), 2);
        assert_eq!(sheet.rows.len(), 2);
    }
}
