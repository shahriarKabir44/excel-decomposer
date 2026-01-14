# 🚀 XL-Chunker

**XL-Chunker** is a high-performance Node.js CLI tool designed to split massive Excel files (500k+ rows) into manageable, memory-safe CSV chunks.



## ⚡ Why XL-Chunker?

Most Excel parsers try to load the entire file into RAM, which causes "Out of Memory" crashes on large datasets. **XL-Chunker** uses a **Streaming Architecture**, processing data row-by-row to keep memory usage low (typically under 100MB) regardless of file size.

- **Memory Efficient:** Streams data directly from disk to disk.
- **Precision Ranges:** Target specific tabs and cell ranges (e.g., `A1:AA540000`).
- **Smart Chunking:** Automatically breaks datasets into smaller files based on your row limit.
- **Data Integrity:** Correctly handles blank cells, formulas, and sparse rows to prevent column shifting.

---

## 🛠 Installation

Install it globally via NPM to use the CLI command anywhere on your system:

```bash
npm install -g xl-chunker