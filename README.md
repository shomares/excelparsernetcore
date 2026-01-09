# Streaming Excel Reader (XLSX)

A **high-performance streaming Excel (.xlsx) reader** for .NET, designed to process **very large files** with **constant memory usage**.

This library reads Excel files directly from their underlying XML structure and provides a **forward-only, pull-based API**. It does **not** rely on the OpenXML SDK and is suitable for ETL pipelines, imports, and large data processing.

---

## ✨ Features

- 🚀 Streaming / forward-only row reading
- 📉 Constant memory usage (1M+ rows supported)
- 🧵 Async file loading
- 📄 Reads directly from `sheet.xml`
- 🔢 Culture-safe numeric parsing (`decimal`)
- 🧠 Optional shared strings caching
- 🏷 Header-based column mapping
- ❌ No OpenXML SDK
- ❌ No LINQ required

---

## 📦 Installation

Clone the repository or reference the project directly:

```bash
git clone https://github.com/shomares/excelparsernetcore/tree/master
```

## 🚀 Basic Usage


```code
var reader = new ReaderExcelSimple();

await reader.ReadFileAsync(
    "Files/bigfile.xlsx",
    "sheet1",
    new ConfigurationReader
    {
        HasHeaders = true,
        UseMemoryForStrings = false
    }
);

foreach (var row in reader.GetNextRow( "sheet1",))
{
    Console.WriteLine(row.ip_address);
}
```

## 🧵 Streaming Design

Rows are read one at a time

No worksheet buffering

Enumeration is single-pass

Memory usage does not grow with file size

⚠ Important
Avoid calling .ToList(), .GroupBy(), or other materializing operations if you want true streaming behavior.

## 🧪 Example Test (Streaming-Safe)

```code
int count = 0;
dynamic? firstRow = null;

foreach (var row in reader.GetNextRow("sheet1"))
{
    if (count == 0)
        firstRow = row;

    count++;
}

Assert.That(count, Is.EqualTo(1000020));
Assert.That(firstRow!.ip_address, Is.EqualTo("86.129.5.175"));
```
## Configuration Options

```code
new ConfigurationReader
{
    HasHeaders = true,
    UseMemoryForStrings = false
}
```
| Option                | Description                                       |
| --------------------- | ------------------------------------------------- |
| `HasHeaders`          | First row contains column names                   |
| `UseMemoryForStrings` | Cache shared strings in memory (faster, more RAM) |


## Limitations

Forward-only reading

Enumeration is single-use

General number format does not expose decimal precision

Dates are numeric Excel serial values

No random row access

## 🧠 Performance Use Cases

ETL pipelines

Data imports

CSV / Excel migrations

Processing files larger than available RAM


## 📄 License

MIT License











