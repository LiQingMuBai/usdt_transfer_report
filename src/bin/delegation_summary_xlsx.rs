use chrono::{FixedOffset, NaiveDate, NaiveDateTime, Utc};
use indicatif::{MultiProgress, ProgressBar, ProgressDrawTarget, ProgressStyle};
use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Format, Workbook};
use std::collections::{BTreeMap, HashMap};
use std::env;
use std::error::Error;
use std::fs::File;

fn normalize_header(s: &str) -> String {
    s.replace('\u{00A0}', " ").trim().to_string()
}

fn get<'a>(row: &'a csv::StringRecord, idx: &HashMap<String, usize>, header: &str) -> &'a str {
    idx.get(&normalize_header(header))
        .and_then(|&i| row.get(i))
        .unwrap_or("")
        .trim()
}

fn normalize_number_like(s: &str) -> String {
    s.replace('\u{00A0}', " ")
        .replace(',', "")
        .split_whitespace()
        .collect::<String>()
}

fn map_resource_quantity(s: &str) -> Option<u32> {
    match normalize_number_like(s).as_str() {
        "131000" => Some(2),
        "65000" => Some(1),
        "2" => Some(2),
        "1" => Some(1),
        _ => None,
    }
}

fn parse_delegation_date_utc8(s: &str) -> Option<NaiveDate> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }
    NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S")
        .ok()
        .map(|dt| dt.date())
}

fn default_input_csv() -> String {
    let home = std::env::var("HOME").unwrap_or_else(|_| ".".to_string());
    format!("{}/Desktop/resource-orders.csv", home)
}

fn default_output_xlsx() -> String {
    "delegation_summary.xlsx".to_string()
}

fn task_progress_style() -> ProgressStyle {
    ProgressStyle::with_template(
        "{spinner:.cyan} [{elapsed_precise}] {msg}\n{wide_bar:.cyan/blue} {pos}/{len} tasks",
    )
    .unwrap()
    .progress_chars("█▓░")
}

fn row_progress_style() -> ProgressStyle {
    ProgressStyle::with_template(
        "  {spinner:.green} [{elapsed_precise}] 解析数据 {pos}/{len} rows ({per_sec})\n  {wide_bar:.green/black} {percent:>3}%",
    )
    .unwrap()
    .progress_chars("█▓░")
}

fn count_records(input_csv: &str) -> Result<u64, Box<dyn Error>> {
    let file = File::open(input_csv)?;
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(true)
        .flexible(true)
        .from_reader(file);

    let mut total = 0u64;
    for result in rdr.records() {
        result?;
        total += 1;
    }
    Ok(total)
}

fn env_non_empty(key: &str) -> Option<String> {
    env::var(key)
        .ok()
        .map(|v| v.trim().to_string())
        .filter(|v| !v.is_empty())
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut args = env::args().skip(1);
    let arg_input_csv = args.next();
    let arg_output_xlsx = args.next();

    dotenvy::dotenv().ok();

    let input_csv = arg_input_csv
        .or_else(|| env_non_empty("INPUT_CSV"))
        .unwrap_or_else(default_input_csv);
    let output_xlsx = arg_output_xlsx
        .or_else(|| env_non_empty("OUTPUT_XLSX"))
        .unwrap_or_else(default_output_xlsx);

    let mp = MultiProgress::new();
    mp.set_draw_target(ProgressDrawTarget::stdout_with_hz(12));

    let utc8 = FixedOffset::east_opt(8 * 3600).ok_or("无效时区偏移")?;
    let today_utc8 = Utc::now().with_timezone(&utc8).date_naive();

    let task_bar = mp.add(ProgressBar::new(5));
    task_bar.set_style(task_progress_style());
    task_bar.set_message("准备读取 CSV");
    task_bar.enable_steady_tick(std::time::Duration::from_millis(120));

    let total_rows = count_records(&input_csv)?;
    task_bar.inc(1);
    task_bar.set_message(format!("读取并汇总派发时间（已扫描 {} 行）", total_rows));

    let file = File::open(&input_csv)?;
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(true)
        .flexible(true)
        .from_reader(file);

    let headers = rdr.headers()?.clone();
    let mut header_idx = HashMap::new();
    for (i, h) in headers.iter().enumerate() {
        header_idx.insert(normalize_header(h), i);
    }

    let mut sums: BTreeMap<NaiveDate, (f64, f64)> = BTreeMap::new();
    let parse_bar = mp.add(ProgressBar::new(total_rows));
    parse_bar.set_style(row_progress_style());
    parse_bar.enable_steady_tick(std::time::Duration::from_millis(120));

    for result in rdr.records() {
        let record = result?;
        parse_bar.inc(1);
        if record.iter().all(|v| v.trim().is_empty()) {
            continue;
        }

        let day = match parse_delegation_date_utc8(get(&record, &header_idx, "Delegation Time (UTC+8)"))
        {
            Some(d) => d,
            None => continue,
        };
        if day == today_utc8 {
            continue;
        }

        let qty = match map_resource_quantity(get(&record, &header_idx, "Resource Quantity")) {
            Some(q) => q,
            None => continue,
        };

        let entry = sums.entry(day).or_insert((0.0, 0.0));
        match qty {
            1 => entry.0 += 1.0,
            2 => entry.1 += 1.0,
            _ => {}
        }
    }
    parse_bar.finish_with_message("  已完成数据汇总");

    if sums.is_empty() {
        return Err("汇总结果为空，无法生成 Excel".into());
    }

    task_bar.inc(1);
    task_bar.set_message(format!("完成 2/5: 已按日期汇总，共 {} 天", sums.len()));
    task_bar.set_message("写入 Excel 表格");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("汇总")?;

    let bold = Format::new().set_bold();
    worksheet.write_with_format(0, 0, "派发时间", &bold)?;
    worksheet.write_with_format(0, 1, "有U能量笔数", &bold)?;
    worksheet.write_with_format(0, 2, "无U能量笔数", &bold)?;
    worksheet.write_with_format(0, 3, "每日转账次数", &bold)?;

    for (i, (day, (sum1, sum2))) in sums.iter().enumerate() {
        let row = i as u32 + 1;
        worksheet.write_string(row, 0, &day.to_string())?;
        worksheet.write_number(row, 1, *sum1)?;
        worksheet.write_number(row, 2, *sum2 * 2.0)?;
        let transfer_count = *sum1 + *sum2;
        worksheet.write_number(row, 3, transfer_count)?;
    }

    let row_max = sums.len() as u32;
    let total_row = row_max + 1;
    let last_data_excel_row = row_max + 1;
    worksheet.write_with_format(total_row, 0, "总计", &bold)?;
    let total_formula_1 = format!("=SUM(B2:B{})", last_data_excel_row);
    worksheet.write_formula_with_format(total_row, 1, total_formula_1.as_str(), &bold)?;
    let total_formula_2 = format!("=SUM(C2:C{})", last_data_excel_row);
    worksheet.write_formula_with_format(total_row, 2, total_formula_2.as_str(), &bold)?;
    let total_formula_3 = format!("=SUM(D2:D{})", last_data_excel_row);
    worksheet.write_formula_with_format(total_row, 3, total_formula_3.as_str(), &bold)?;

    task_bar.inc(1);
    task_bar.set_message("完成 3/5: 汇总数据已写入工作表（含总计）");
    task_bar.set_message("生成柱状图");

    let mut chart = Chart::new(ChartType::Column);
    chart.set_width(900).set_height(520);

    chart
        .add_series()
        .set_categories(("汇总", 1, 0, row_max, 0))
        .set_values(("汇总", 1, 1, row_max, 1))
        .set_name("有U能量笔数")
        .set_format(ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#1F4E79")));

    chart
        .add_series()
        .set_categories(("汇总", 1, 0, row_max, 0))
        .set_values(("汇总", 1, 2, row_max, 2))
        .set_name("无U能量笔数")
        .set_format(ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#7F6000")));

    chart
        .add_series()
        .set_categories(("汇总", 1, 0, row_max, 0))
        .set_values(("汇总", 1, 3, row_max, 3))
        .set_name("每日转账次数")
        .set_format(ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#385723")));

    chart.title().set_name("按派发时间日期汇总");
    chart.x_axis().set_name("派发时间");
    chart.y_axis().set_name("笔数");
    chart.set_style(10);

    worksheet.insert_chart_with_offset(1, 5, &chart, 10, 10)?;
    task_bar.inc(1);
    task_bar.set_message("完成 4/5: 柱状图已插入 Excel");
    task_bar.set_message("保存文件");

    workbook.save(&output_xlsx)?;
    task_bar.inc(1);
    task_bar.finish_with_message("全部任务完成");

    Ok(())
}
