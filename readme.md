# 📊 Doc-to-Sheet Auto Reporting

A lightweight Google Apps Script tool to help content writers, editors, and managers **automatically log daily word counts** from each writer’s Google Doc into a shared Google Sheet.

---

## 🔍 Problem Statement – What It’s For

Managing multiple writers and tracking daily output manually can be time-consuming and error-prone.

This tool eliminates the hassle by automatically reading linked Google Docs, identifying writing dates, and logging daily word counts into the appropriate cells of a structured spreadsheet.

---

## ⚙️ How It Works

- Reads each writer’s Google Doc (linked in your Sheet).
- Detects the **date** written as a title (required format).
- Calculates the **word count** for each date section (including the date line).
- Matches each date to the corresponding column in your Google Sheet.
- Logs the word count next to the correct writer's name and under the relevant date.

---

## 🧑‍💻 How to Use

### 1. ✅ Rules for Writers

Writers should follow a consistent format within their Docs:

- Begin each day’s entry with a **date title**.
- Write all content for that day **below the date heading**.

#### 📄 Example Doc Content:
June 1, 2025
Today I worked on the new homepage copy and revised two blogs...


- Only the **date** should be formatted as a title.
- All body content should follow in other styles (h1, h2, h3, para).

---

### 2. 🧾 How to Set Up the Google Sheet

#### 🧱 Sheet Structure:

| A (Name)     | B (Doc File Link) | C (June 1, 2025) | D (June 2, 2025) | E (June 3, 2025) |
|--------------|------------------|------------------|------------------|------------------|
| writer_1     | link_1           |                  |                  |                  |
| writer_2     | link_2           |                  |                  |                  |

- **Column A (A2 onward)**: Writer names  
- **Column B (B2 onward)**: Links to their Google Docs  
- **Row 1 (starting from Column C)**: Dates to track

#### ⚙️ Script Configuration:
You can customize where the script begins reading by changing this setting:
```js startCell: "C2" ```

🧠 Notes
✅ Word Count Includes the Date Heading
If a writer adds June 1, 2025 as the heading, those 3 words are counted too.

📅 Automatically Parses the Following Date Formats:
June 1, 2025

1st June 2025

01/06/2025

1/6/25

2025-06-01

June 1

1 June

1 Jun 2025

All dates are internally normalized to MM/DD/YYYY format for matching.

🧩 Customizable Start Cell
Update startCell (e.g., "C2") to match your sheet layout.

🙋 Need Help?
Have a feature suggestion or found an issue?
Open an issue or submit a pull request—collaboration is welcome!

✍️ Author:

Anuj Upreti
Helping writers and editors save time with automation.