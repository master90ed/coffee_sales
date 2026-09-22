# Coffee Shop Sales

This project contains transaction-level sales data from a coffee shop. Each row represents one transaction and includes the store, product, quantity, and unit price.

## Dataset

- File: `Coffee Shop Sales.csv`
- Format: CSV
- Grain: One row per transaction

### Columns

| Column | Description |
| --- | --- |
| `transaction_id` | Unique transaction identifier |
| `transaction_date` | Date of the transaction |
| `transaction_time` | Time of the transaction |
| `transaction_qty` | Number of items sold |
| `store_id` | Store identifier |
| `store_location` | Store location |
| `product_id` | Product identifier |
| `unit_price` | Price per item |
| `product_category` | Broad product category |
| `product_type` | Product type |
| `product_detail` | Specific product description |

## Analysis with Python

```python
python analysis.py "Coffee Shop Sales.csv"
```

The script uses pandas, NumPy, Matplotlib, and Seaborn to calculate sales summaries and save these charts in `analysis_outputs/`:

- Revenue by product category
- Revenue by store
- Revenue by transaction hour

To open the charts in separate Matplotlib windows while running locally, use:

```bash
python analysis.py "Coffee Shop Sales.csv" --show-charts
```

It also accepts Excel workbooks:

```bash
python analysis.py sales.xlsx --output-dir charts
```

For a quick interactive analysis:

```python
import pandas as pd

sales = pd.read_csv("Coffee Shop Sales.csv")
sales["revenue"] = sales["transaction_qty"] * sales["unit_price"]

print(sales.head())
print("Total revenue:", sales["revenue"].sum())
print(sales.groupby("product_category")["revenue"].sum().sort_values(ascending=False))
```

## Notes

- Parse `transaction_date` and `transaction_time` as date/time values for time-based analysis.
- Revenue can be estimated as `transaction_qty * unit_price`.
- Keep analysis outputs in separate files or notebooks rather than modifying the source CSV.
