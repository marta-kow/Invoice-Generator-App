# Invoice PDF Creator

An application for small businesses to generate professional PDF invoices from a Word template and automatically store invoice records in a SQL database.

<img width="512" height="639" alt="obraz" src="https://github.com/user-attachments/assets/06d0801e-fb02-425b-8e76-1088f083d022" /><img width="633" height="842" alt="obraz" src="https://github.com/user-attachments/assets/726b6865-0efb-418f-9104-df3643d99d1b" />

---

## Features

- Fill in customer details, product, quantity, price and VAT rate through a clean dark-mode GUI
- Auto-calculates net amount, tax and gross total
- Populates a `.docx` template with the entered data and converts it to PDF in one click
- Supports three payment terms out of the box: **Immediate**, **Net 7**, **Net 30**
- Saves every invoice record to a SQL database (MySQL) via SQLAlchemy ORM
- Database credentials loaded from a `.env` file to protect private data

---

## Tech Stack

| Layer | Library |
|---|---|
| GUI | CustomTkinter|
| Word templating | python-docx|
| PDF conversion | docx2pdf|
| ORM / Database | SQLAlchemy|
| Env variables | python-dotenv|

---

## How It Works

```
User fills the form
        │
Placeholders replaced in .docx template
        │
Temp .docx saved -> converted to PDF via docx2pdf -> temp .docx deleted
        │
Invoice record written to SQL database (SQLAlchemy ORM)
```

## Database Schema

The `invoices` table is created automatically on first run via `Base.metadata.create_all()`.

| Column | Type | Description |
|---|---|---|
| id | Integer (PK) | Auto-increment |
| customer_name | String | |
| customer_address | String | |
| customer_vat_id | String | |
| invoice_number | String | |
| product | String | |
| quantity | Integer | |
| price | Float | Unit price (net) |
| vat_rate | Float | e.g. 23.0 for 23% |
| netto | Float | quantity × price |
| tax_amount | Float | netto × vat_rate / 100 |
| total | Float | netto + tax_amount |
| payment_terms | String | |
| invoice_date | Date | Set to today on creation |

---


## Limitations

- PDF conversion uses `docx2pdf`, which requires **Microsoft Word on Windows**.
- Currently supports **one line item per invoice**. 
- The `.env.txt` path is hardcoded — this should be replaced with a relative path.

---

## License

MIT
