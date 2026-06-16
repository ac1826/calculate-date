# Export Data Processor

Streamlit app for converting Dalian and Tieling export Excel files into the monthly export summary format.

## Run

```powershell
pip install -r requirements.txt
streamlit run app.py
```

## Inputs

- Dalian sales order export: includes customer, product, ordered quantity, price, currency, and date.
- Tieling detail export: includes product, invoice quantity in kg, and sales price in CNY/kg.

## Output

The generated workbook contains:

- 外销市场
- 外销客户
- 产品名称
- 月销量（吨）
- 单价(元/吨)
- 币种
- 出口方

Market subtotals are inserted with Excel formulas, and repeated market/customer/exporter cells are merged to match the sample workbook style.
