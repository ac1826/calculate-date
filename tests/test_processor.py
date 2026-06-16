from pathlib import Path

from openpyxl import load_workbook

from processor import export_summary_workbook, process_files


ROOT = Path(__file__).resolve().parents[1]
DALIAN = Path("F:/llqdocument/大成文件/出口RPA抓取/5月_大连.xlsx")
TIELING = Path("F:/llqdocument/大成文件/出口RPA抓取/5月_铁岭.xlsx")


def test_sample_transform_values():
    tables = process_files(DALIAN, TIELING)

    dalian_row = tables.dalian[tables.dalian["产品名称"] == "新加坡CMM 脆皮鸡肉1000g/12袋/箱"].iloc[0]
    assert dalian_row["月销量（吨）"] == 2.4
    assert dalian_row["单价(元/吨)"] == 3880
    assert dalian_row["币种"] == "USD"

    tieling_row = tables.tieling[tables.tieling["产品名称"] == "麦乐鸡块 500g*30"].iloc[0]
    assert tieling_row["月销量（吨）"] == 3.75
    assert round(tieling_row["单价(元/吨)"], 6) == 17830
    assert tieling_row["币种"] == "RMB"

    dc_extra = tables.dalian[tables.dalian["产品名称"] == "新加坡DC照烧鸡排100g/10枚/12袋/箱"].iloc[0]
    assert dc_extra["月销量（吨）"] == 1.56


def test_export_workbook_has_headers_formulas_and_merges(tmp_path):
    tables = process_files(DALIAN, TIELING)
    output = tmp_path / "summary.xlsx"
    output.write_bytes(export_summary_workbook(tables.final, month_label="5月"))

    wb = load_workbook(output, data_only=False)
    ws = wb.active

    assert ws["A1"].value == "外销市场"
    assert ws["D1"].value == "5月销量\n（吨）"
    assert ws["D19"].value == "=SUM(D2:D18)"
    assert ws["D27"].value == "=SUM(D20:D26)"
    assert "A2:A19" in {str(rng) for rng in ws.merged_cells.ranges}
    assert "G22:G26" in {str(rng) for rng in ws.merged_cells.ranges}
