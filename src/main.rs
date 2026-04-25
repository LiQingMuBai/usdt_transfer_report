use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::{self, Write};

#[derive(Debug, Default)]
struct OrderRow {
    order_id: String,
    client_order_id: String,
    resource_type: String,
    source_type: String,
    payment_time_utc8: String,
    receiver_address: String,
    delegation_hash: String,
    delegation_time_utc8: String,
    payment_amount: String,
    activation_amount: String,
    resource_quantity: String,
    staked_amount: String,
    order_status: String,
    activation_status: String,
    confirmation_status: String,
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

fn normalize_number_like(s: &str) -> String {
    s.replace('\u{00A0}', " ")
        .replace(',', "")
        .split_whitespace()
        .collect::<String>()
}

fn map_resource_quantity(s: &str) -> String {
    match normalize_number_like(s).as_str() {
        "131000" => "2".to_string(),
        "65000" => "1".to_string(),
        _ => s.trim().to_string(),
    }
}

fn write_row<W: Write>(w: &mut W, i: usize, r: &OrderRow) -> io::Result<()> {
    writeln!(w, "--- row {} ---", i)?;
    writeln!(w, "Order ID: {}", r.order_id)?;
    writeln!(w, "Client Order ID: {}", r.client_order_id)?;
    writeln!(w, "Resource Type: {}", r.resource_type)?;
    writeln!(w, "Source Type: {}", r.source_type)?;
    writeln!(w, "Payment Time (UTC+8): {}", r.payment_time_utc8)?;
    writeln!(w, "Receiver Address: {}", r.receiver_address)?;
    writeln!(w, "Delegation Hash: {}", r.delegation_hash)?;
    writeln!(w, "Delegation Time (UTC+8): {}", r.delegation_time_utc8)?;
    writeln!(w, "Payment Amount: {}", r.payment_amount)?;
    writeln!(w, "Activation Amount: {}", r.activation_amount)?;
    writeln!(w, "Resource Quantity: {}", r.resource_quantity)?;
    writeln!(w, "Staked Amount: {}", r.staked_amount)?;
    writeln!(w, "Order Status: {}", r.order_status)?;
    writeln!(w, "Activation Status: {}", r.activation_status)?;
    writeln!(w, "Confirmation Status: {}", r.confirmation_status)?;
    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut args = std::env::args().skip(1);
    let path = args.next().unwrap_or_else(|| {
        let home = std::env::var("HOME").unwrap_or_else(|_| ".".to_string());
        format!("{}/Desktop/resource-orders.csv", home)
    });

    let file = File::open(&path)?;
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(true)
        .flexible(true)
        .from_reader(file);

    let headers = rdr.headers()?.clone();
    let mut header_idx = HashMap::new();
    for (i, h) in headers.iter().enumerate() {
        header_idx.insert(normalize_header(h), i);
    }

    let stdout = io::stdout();
    let mut out = stdout.lock();

    let mut count = 0usize;
    for result in rdr.records() {
        let record = result?;
        if record.iter().all(|v| v.trim().is_empty()) {
            continue;
        }

        count += 1;
        let row = OrderRow {
            order_id: get(&record, &header_idx, "Order ID").to_string(),
            client_order_id: get(&record, &header_idx, "Client Order ID").to_string(),
            resource_type: get(&record, &header_idx, "Resource Type").to_string(),
            source_type: get(&record, &header_idx, "Source Type").to_string(),
            payment_time_utc8: get(&record, &header_idx, "Payment Time (UTC+8)").to_string(),
            receiver_address: get(&record, &header_idx, "Receiver Address").to_string(),
            delegation_hash: get(&record, &header_idx, "Delegation Hash").to_string(),
            delegation_time_utc8: get(&record, &header_idx, "Delegation Time (UTC+8)").to_string(),
            payment_amount: get(&record, &header_idx, "Payment Amount").to_string(),
            activation_amount: get(&record, &header_idx, "Activation Amount").to_string(),
            resource_quantity: map_resource_quantity(get(&record, &header_idx, "Resource Quantity")),
            staked_amount: get(&record, &header_idx, "Staked Amount").to_string(),
            order_status: get(&record, &header_idx, "Order Status").to_string(),
            activation_status: get(&record, &header_idx, "Activation Status").to_string(),
            confirmation_status: get(&record, &header_idx, "Confirmation Status").to_string(),
        };

        if let Err(e) = write_row(&mut out, count, &row) {
            if e.kind() == io::ErrorKind::BrokenPipe {
                return Ok(());
            }
            return Err(e.into());
        }
    }

    writeln!(out, "共读取 {} 行（文件：{}）", count, path)?;
    Ok(())
}
