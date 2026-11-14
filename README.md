#  Data Cleaning & Structural Validation

A Python project that transforms a raw, messy Excel sales sheet into a reliable, analysis-ready table. It handles **duplicate entries**, **missing values** and **inconsistent formats** while preserving data integrity: every change is logged and nothing is silently deleted.

## 📊 Dataset

`Data_Cleaning_Start.xlsx` – 31 sales records with columns: Date, Client, Contact, Department, Payment, Revenue, Profit, Profit Margin.

### Problems found in the raw data

| Problem | Example |
|---------|---------|
| Extra / uneven spaces | `"     Bill SmITH"`, `"Harley        Fritz"` |
| Inconsistent capitalisation | `"IVAN HINEY"`, `"jaylynn   napp"` |
| Several values in one column | `"AMAZON.COM, INC. (XNAS:AMZN)"`, `"Cloud Tech_Texas"` |
| Missing values | Blank Payment (2 rows), blank Revenue (2 rows) |
| Duplicate rows | Costco, McDonald's, Goldman Sachs entered twice |
| Profit Margin errors | Formula breaks when Revenue is blank |

## ✨ What the script does

1. **Removes exact duplicates** (3 rows) – keeps the first occurrence
2. **Standardises contact names** – trims spaces, Proper Case (`Bill Smith`)
3. **Splits Client** into `Company`, `Exchange`, `Ticker`
4. **Splits Department** into `Department` and `Location`
5. **Missing Payment** → `Unknown`
6. **Missing Revenue** → left blank and highlighted (never guessed); Profit Margin shows blank instead of an error
7. **Recomputes Profit Margin** as a live Excel formula
8. **Sorts by date** and writes everything to `Cleaned_Data.xlsx`

## 📁 Project Structure

```
data-cleaning-validation/
├── data_cleaner.py            # Main script
├── Data_Cleaning_Start.xlsx   # Raw input data
├── Cleaned_Data.xlsx          # Output (Cleaned Data, Change Log, Raw Data sheets)
├── requirements.txt
└── README.md
```

## 🚀 How to Run

```bash
git clone https://github.com/<your-username>/data-cleaning-validation.git
cd data-cleaning-validation
pip install -r requirements.txt
python data_cleaner.py
```

## 📈 Result

| | Rows |
|---|---|
| Raw records | 31 |
| Duplicates removed | 3 |
| **Clean records** | **28** |
| Missing Revenue flagged | 2 |
| Missing Payment set to Unknown | 2 |

## 🛡️ Data Integrity Decisions

- Original data is kept in the **Raw Data** sheet for comparison.
- The **Change Log** sheet lists every edit with the original and new value.
- Missing financial values are flagged, not invented.

## 🔮 Future Improvements

- Typo detection in names (e.g. `Franklin Wrigt`)
- Command-line options for custom input/output files
- Data quality charts

## 👤 Author

**<Your Name>** – [GitHub](https://github.com/<your-username>)
