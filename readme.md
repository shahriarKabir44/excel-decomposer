# 🧩 xl-chunker

**xl-chunker** is a lightweight **Node.js CLI tool** that slices **huge Excel files** into **smaller, manageable CSV chunks** using streaming — without blowing up your memory.

Built for people who deal with **massive spreadsheets** and just want things to *work*.

---

## 🚀 What it Does

- 📂 Reads **large Excel files** efficiently using streaming
- ✂️ Breaks data into **row-based chunks**
- 🧾 Preserves **headers** in every output file
- 🎯 Lets you define an exact **cell range** (e.g. `B4` → `BN542549`)
- 📄 Outputs clean **CSV files**
- 💾 Saves your settings so you can rerun easily

---

## 🛠️ Tech Stack

- **Node.js**
- **ExcelJS (stream reader)** – no full-file memory load
- **CSV output** for maximum compatibility

---

## 📦 Installation (Global)

Install **xl-chunker** globally using npm:

```bash
npm install -g xl-chunker
```

> Requires **Node.js v14+**

---

## ▶️ Usage

Once installed globally, run it from anywhere:

```bash
xl-chunker
```

You’ll be prompted for:

- Source file path & name
- Worksheet (tab) name
- Start & end cell (e.g. `B4` to `BN542549`)
- Chunk size (rows per file)
- Output directory
- Output file name template

---

## 📂 Example Output

```
users_part1.csv
users_part2.csv
users_part3.csv
...
```

Each file automatically includes the **header row**.

---

## 🔁 Reusing Settings

After the first successful run, a `Setting.json` file is generated.

Next time, just provide the path to this file and skip all prompts 🚀

---

## ⚡ Why xl-chunker?

- ❌ No Excel crashes
- ❌ No RAM overload
- ✅ Fast & streaming-based
- ✅ Simple CLI
- ✅ Perfect for automation

Ideal for:
- Data migration
- ETL pipelines
- Upload size limits
- Analytics preprocessing

---

## 🧠 Pro Tip

If Excel struggles to open your file, **xl-chunker won’t**.  
It processes rows one-by-one like a champ 💪

---

## 🧪 Local Development

Clone and test locally:

```bash
git clone https://github.com/your-username/xl-chunker.git
cd xl-chunker
npm install
npm link
xl-chunker
```

---

## 📄 License

MIT — do whatever you want, just don’t blame me 😉
