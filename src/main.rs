use std::collections::{BTreeMap, HashMap};
use std::error::Error;
use std::fs::File;
use std::io::{self, Read, Write};
use std::path::Path;

use chrono::{Duration, NaiveDate, NaiveDateTime, Utc};
use dotenvy::dotenv;
use indicatif::{ProgressBar, ProgressDrawTarget, ProgressStyle};
use rust_xlsxwriter::{
    Chart, ChartFormat, ChartSolidFill, ChartType, Format, Workbook,
};

#[derive(Debug, Default)]
struct TransferRow {
    txn_hash: String,
    block: String,
    time_utc: String,
    from: String,
    to: String,
    token: String,
    token_symbol: String,
    amount_or_token_id: String,
    result: String,
    status: String,
}

fn normalize_header(s: &str) -> String {
    s.replace('\u{00A0}', " ").trim().to_string()
}

fn get<'a>(row: &'a csv::StringRecord, idx: &HashMap<String, usize>, header: &str) -> &'a str {
    idx.get(&normalize_header(header))
        .and_then(|&i| row.get(i))
        .unwrap_or("")
        .trim()
}

fn parse_utc_to_beijing_naive(s: &str) -> Option<NaiveDateTime> {
    let raw = s.trim();
    if raw.is_empty() {
        return None;
    }

    let formats = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"];
    for fmt in formats {
        if let Ok(naive) = NaiveDateTime::parse_from_str(raw, fmt) {
            return Some(naive + Duration::hours(8));
        }
    }

    None
}

fn utc_to_beijing(s: &str) -> String {
    parse_utc_to_beijing_naive(s)
        .map(|dt| dt.format("%Y-%m-%d %H:%M:%S").to_string())
        .unwrap_or_else(|| s.trim().to_string())
}

fn bjt_day_key_from_utc(s: &str) -> String {
    parse_utc_to_beijing_naive(s)
        .map(|dt| dt.date().format("%Y-%m-%d").to_string())
        .unwrap_or_else(|| "UNKNOWN".to_string())
}

fn parse_day_key(s: &str) -> Option<NaiveDate> {
    NaiveDate::parse_from_str(s.trim(), "%Y-%m-%d").ok()
}

fn write_row<W: Write>(w: &mut W, i: usize, r: &TransferRow) -> io::Result<()> {
    writeln!(w, "--- transfer {} ---", i)?;
    writeln!(w, "Txn Hash: {}", r.txn_hash)?;
    writeln!(w, "Block: {}", r.block)?;
    writeln!(w, "Time(BJT): {}", utc_to_beijing(&r.time_utc))?;
    writeln!(w, "From: {}", r.from)?;
    writeln!(w, "To: {}", r.to)?;
    writeln!(w, "Token: {}", r.token)?;
    writeln!(w, "Token Symbol: {}", r.token_symbol)?;
    writeln!(w, "Amount/TokenID: {}", r.amount_or_token_id)?;
    writeln!(w, "Result: {}", r.result)?;
    writeln!(w, "Status: {}", r.status)?;
    Ok(())
}

fn ok_or_broken_pipe(res: io::Result<()>) -> Result<(), Box<dyn Error>> {
    if let Err(e) = res {
        if e.kind() == io::ErrorKind::BrokenPipe {
            return Ok(());
        }
        return Err(e.into());
    }
    Ok(())
}

fn default_csv_path() -> String {
    let home = std::env::var("HOME").unwrap_or_else(|_| ".".to_string());
    format!("{home}/Desktop/Transfers.csv")
}

fn default_xlsx_path() -> String {
    "transfer_daily_summary.xlsx".to_string()
}

fn env_string(key: &str) -> Option<String> {
    std::env::var(key)
        .ok()
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
}

#[derive(Clone, Copy, Debug)]
enum RangeMode {
    All,
    LastNDays(i64),
}

fn parse_days(s: &str) -> Option<i64> {
    let n = s.trim().parse::<i64>().ok()?;
    if n <= 0 {
        return None;
    }
    Some(n.min(3650))
}

fn load_range_mode_from_env() -> RangeMode {
    if let Ok(v) = std::env::var("RANGE_DAYS") {
        if let Some(n) = parse_days(&v) {
            return RangeMode::LastNDays(n);
        }
    }

    RangeMode::All
}

fn load_exclude_today_from_env() -> bool {
    let Ok(v) = std::env::var("EXCLUDE_TODAY") else {
        return true;
    };
    !matches!(
        v.trim().to_ascii_uppercase().as_str(),
        "0" | "FALSE" | "NO" | "N" | "OFF"
    )
}

fn beijing_today() -> NaiveDate {
    (Utc::now().naive_utc() + Duration::hours(8)).date()
}

fn apply_range_mode(
    daily_counts: &BTreeMap<String, usize>,
    mode: RangeMode,
    exclude_today: bool,
) -> BTreeMap<String, usize> {
    let today = if exclude_today {
        Some(beijing_today())
    } else {
        None
    };

    match mode {
        RangeMode::All => {
            let mut out = daily_counts.clone();
            if let Some(today) = today {
                out.remove(&today.format("%Y-%m-%d").to_string());
            }
            out
        }
        RangeMode::LastNDays(days) => {
            let latest = daily_counts
                .keys()
                .filter_map(|k| parse_day_key(k))
                .max();

            let Some(mut end) = latest else {
                return BTreeMap::new();
            };

            if let Some(today) = today
                && end >= today
            {
                end = today - Duration::days(1);
            }

            let days = days.max(1).min(3650);
            let start = end - Duration::days(days - 1);
            let mut out = BTreeMap::new();
            for i in 0..days {
                let d = start + Duration::days(i);
                let k = d.format("%Y-%m-%d").to_string();
                let c = daily_counts.get(&k).copied().unwrap_or(0);
                out.insert(k, c);
            }
            out
        }
    }
}

fn write_daily_summary_xlsx(
    daily_counts: &BTreeMap<String, usize>,
    output_xlsx: &str,
) -> Result<(), Box<dyn Error>> {
    if let Some(parent) = Path::new(output_xlsx).parent() {
        if !parent.as_os_str().is_empty() {
            std::fs::create_dir_all(parent)?;
        }
    }

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("汇总")?;

    let bold = Format::new().set_bold();
    worksheet.write_with_format(0, 0, "日期(BJT)", &bold)?;
    worksheet.write_with_format(0, 1, "转账次数", &bold)?;

    for (i, (day, c)) in daily_counts.iter().enumerate() {
        let row = i as u32 + 1;
        worksheet.write_string(row, 0, day)?;
        worksheet.write_number(row, 1, *c as f64)?;
    }

    worksheet.set_column_width(0, 14)?;
    worksheet.set_column_width(1, 12)?;

    let row_max = daily_counts.len() as u32;
    if row_max >= 1 {
        let mut chart = Chart::new(ChartType::Column);
        chart.set_width(900).set_height(520);

        chart
            .add_series()
            .set_categories(("汇总", 1, 0, row_max, 0))
            .set_values(("汇总", 1, 1, row_max, 1))
            .set_name("每日转账次数")
            .set_format(
                ChartFormat::new()
                    .set_solid_fill(ChartSolidFill::new().set_color("#1F4E79")),
            );

        chart.title().set_name("每日转账次数（北京时间）");
        chart.x_axis().set_name("日期");
        chart.y_axis().set_name("次数");
        chart.set_style(10);

        worksheet.insert_chart_with_offset(1, 3, &chart, 10, 10)?;
    }

    workbook.save(output_xlsx)?;
    Ok(())
}

struct CountingReader<R> {
    inner: R,
    pb: ProgressBar,
}

impl<R: Read> Read for CountingReader<R> {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        let n = self.inner.read(buf)?;
        if n > 0 {
            self.pb.inc(n as u64);
        }
        Ok(n)
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let _ = dotenv();

    let mut path: Option<String> = None;
    let mut summary_only = false;
    let mut xlsx_path: Option<String> = env_string("XLSX_PATH").or_else(|| env_string("OUTPUT_XLSX"));
    let mut range_mode: Option<RangeMode> = None;
    let mut exclude_today: Option<bool> = None;

    let mut args = std::env::args().skip(1).peekable();
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--summary-only" => summary_only = true,
            "--xlsx" => {
                let next = args.peek().map(|s| s.as_str());
                if let Some(v) = next
                    && !v.starts_with("--")
                {
                    xlsx_path = Some(args.next().unwrap());
                } else {
                    xlsx_path = Some(
                        xlsx_path
                            .take()
                            .unwrap_or_else(|| default_xlsx_path()),
                    );
                }
            }
            "--no-xlsx" => {
                xlsx_path = None;
            }
            "--days" => {
                let v = args.next().unwrap_or_else(|| "40".to_string());
                let days = parse_days(&v).unwrap_or(40);
                range_mode = Some(RangeMode::LastNDays(days));
            }
            "--exclude-today" => {
                exclude_today = Some(true);
            }
            "--include-today" => {
                exclude_today = Some(false);
            }
            _ => {
                if path.is_none() {
                    path = Some(arg);
                }
            }
        }
    }
    let path = path.or_else(|| env_string("CSV_PATH")).unwrap_or_else(default_csv_path);

    if !Path::new(&path).exists() {
        return Err(format!("CSV file not found: {}", path).into());
    }

    let total_bytes = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
    let pb_agg = ProgressBar::new(total_bytes);
    pb_agg.set_draw_target(ProgressDrawTarget::stderr_with_hz(12));
    pb_agg.set_style(
        ProgressStyle::with_template(
            "{spinner:.cyan} [{elapsed_precise}] [{bar:40.cyan/blue}] {bytes}/{total_bytes} ({eta}) {msg}",
        )?
        .progress_chars("█▓▒░ "),
    );
    pb_agg.set_message("Reading CSV & aggregating...");

    let file = File::open(&path)?;
    let reader = CountingReader {
        inner: file,
        pb: pb_agg.clone(),
    };

    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(true)
        .flexible(true)
        .from_reader(reader);

    let headers = rdr.headers()?.clone();
    let mut header_idx = HashMap::new();
    for (i, h) in headers.iter().enumerate() {
        header_idx.insert(normalize_header(h), i);
    }

    let stdout = io::stdout();
    let mut out = stdout.lock();

    let mut count = 0usize;
    let mut daily_counts: BTreeMap<String, usize> = BTreeMap::new();
    for result in rdr.records() {
        let record = result?;
        if record.iter().all(|v| v.trim().is_empty()) {
            continue;
        }

        count += 1;
        let row = TransferRow {
            txn_hash: get(&record, &header_idx, "Txn Hash").to_string(),
            block: get(&record, &header_idx, "Block").to_string(),
            time_utc: get(&record, &header_idx, "Time(UTC)").to_string(),
            from: get(&record, &header_idx, "From").to_string(),
            to: get(&record, &header_idx, "To").to_string(),
            token: get(&record, &header_idx, "Token").to_string(),
            token_symbol: get(&record, &header_idx, "Token Symbol").to_string(),
            amount_or_token_id: get(&record, &header_idx, "Amount/TokenID").to_string(),
            result: get(&record, &header_idx, "Result").to_string(),
            status: get(&record, &header_idx, "Status").to_string(),
        };

        let day = bjt_day_key_from_utc(&row.time_utc);
        *daily_counts.entry(day).or_insert(0) += 1;

        if !summary_only {
            if let Err(e) = write_row(&mut out, count, &row) {
                if e.kind() == io::ErrorKind::BrokenPipe {
                    return Ok(());
                }
                return Err(e.into());
            }
        }
    }

    pb_agg.finish_and_clear();

    let mode = range_mode.unwrap_or_else(load_range_mode_from_env);
    let exclude_today = exclude_today.unwrap_or_else(load_exclude_today_from_env);
    let daily_counts = apply_range_mode(&daily_counts, mode, exclude_today);

    ok_or_broken_pipe(writeln!(out, "=== Daily Summary (BJT) ==="))?;
    for (day, c) in &daily_counts {
        ok_or_broken_pipe(writeln!(out, "{}\t{}", day, c))?;
    }

    if let Some(xlsx) = xlsx_path {
        let pb = ProgressBar::new_spinner();
        pb.set_draw_target(ProgressDrawTarget::stderr_with_hz(12));
        pb.set_style(ProgressStyle::with_template("{spinner:.cyan} [{elapsed_precise}] {msg}")?);
        pb.enable_steady_tick(std::time::Duration::from_millis(90));
        pb.set_message("Writing Excel report...");
        write_daily_summary_xlsx(&daily_counts, &xlsx)?;
        pb.finish_and_clear();
        ok_or_broken_pipe(writeln!(out, "已输出 Excel：{}", xlsx))?;
    }
    ok_or_broken_pipe(writeln!(out, "共读取 {} 行（文件：{}）", count, path))?;
    Ok(())
}
