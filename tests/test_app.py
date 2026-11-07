"""Exercise real Streamlit controls and deserialize actual download workbooks.

Expected prices and P&L use Fraction arithmetic, not production calculators.
No customer ledger, network call, database, or deployed app is used.
"""
import io
import json
import math
import os
import zipfile
from fractions import Fraction
from pathlib import Path

import openpyxl
import pandas as pd
import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest
from streamlit.runtime.uploaded_file_manager import UploadedFile, UploadedFileRec
from streamlit.proto.Common_pb2 import FileURLs

from support import APP, load_helpers


@pytest.fixture
def captures(monkeypatch):
    captured = []
    original = openpyxl.Workbook.save

    def save(wb, target):
        original(wb, target)
        if wb.sheetnames == ["제안서", "매출원가산출표", "예상손익분석표"]:
            buf = io.BytesIO()
            original(wb, buf)
            captured.append(buf.getvalue())

    monkeypatch.setattr(openpyxl.Workbook, "save", save)
    return captured


def new_app(path=APP):
    at = AppTest.from_file(str(path), default_timeout=30).run()
    assert not at.exception
    return at


def seeded_app(path=APP):
    at = new_app(path)
    at.session_state.q_warehouse_df = pd.DataFrame([
        {"구분": "입고", "작업 내용": "검수(기본)", "단위": "건", "매입단가": 10000., "제안단가": 12789., "비고": "가상 검증"},
        {"구분": "보관", "작업 내용": "작업스테이션(*)", "단위": "평", "매입단가": 0., "제안단가": 40000., "비고": "0원 원가 수동단가 유지"}])
    at.session_state.q_express_df = pd.DataFrame([
        {"국가": "TEST-A", "기준구간": "2kg", "매입원가": 10000., "판매 제안단가": 12987., "발송비중(%)": 75., "비고": "가상 A"},
        {"국가": "TEST-B", "기준구간": "2kg", "매입원가": 20000., "판매 제안단가": 23765., "발송비중(%)": 25., "비고": "가상 B"}])
    at.session_state.q_infra_df = pd.DataFrame([
        {"구분": "영업원가", "작업 내용": "가상 센터 고정비", "단위": "식/월", "매입단가": 100000., "비고": "가상 원가"}])
    for key, val in dict(vol_b2c=100, inbound_qty=100, multi_pcs=0, storage_pyeong=0,
                         express_qty=200, dim_qty=0, customs_qty=0).items():
        at.number_input(key=f"in_q_{key}").set_value(val)
    at.run()
    assert not at.exception
    return at


def control(at, prefix):
    return next(s for s in at.slider if s.key.startswith(prefix))


def price(cost, rate, basis):
    return (math.ceil(Fraction(cost) / (1 - Fraction(rate, 100)) / 10) * 10
            if basis == "gross_margin" else round(cost * (1 + rate / 100), -1))


def reprice(at, rate, basis):
    at.radio(key="pricing_basis_tab5").set_value(basis).run()
    at.slider(key="slider_wh_target_margin").set_value(float(rate))
    at.button(key="btn_apply_wh_markup").click().run()
    control(at, "slider_cogs_exp_margin_").set_value(float(rate))
    at.button(key="btn_cogs_recalc_exp").click().run()
    assert not at.exception


def workbook_values(wb):
    return {ws.title: [(c.coordinate, c.value, c.number_format, c.style_id) for row in ws for c in row]
            for ws in wb}


def verify_downloads(at, captures, rate, basis, include_exp, include_infra, sga_rate=7):
    # All three download buttons must reflect this same completed UI run.
    books = [openpyxl.load_workbook(io.BytesIO(b)) for b in captures[-3:]]
    assert len(books) == 3
    assert workbook_values(books[0]) == workbook_values(books[1]) == workbook_values(books[2])
    wb = books[-1]
    p = price(10000, rate, basis)
    p2 = price(20000, rate, basis)
    wh_rev = p * 100
    exp_rev = (Fraction(3, 4) * p + Fraction(1, 4) * p2) * 200 if include_exp else 0
    rev = wh_rev + exp_rev
    direct = 1000000 + (2500000 if include_exp else 0)
    infra = 100000 if include_infra else 0
    sga = round(Fraction(wh_rev) * Fraction(sga_rate, 100))
    total_cost = direct + infra + sga
    op = rev - total_cost
    expected = {"G": rev, "M": total_cost, "N": op, "O": f"{float(op / rev * 100):.1f}%"}
    ws = wb["예상손익분석표"]
    total_row = next(c.row for c in ws["B"] if c.value == "매출액 합계")
    for col, val in expected.items():
        assert ws[f"{col}{total_row}"].value == val
    assert next(m.value for m in at.metric if m.label == "1. 예상 월 매출액") == f"{float(rev):,.0f}원"
    assert next(m.value for m in at.metric if m.label == "2. 예상 월 매입원가") == f"{float(total_cost):,.0f}원"
    assert next(m.value for m in at.metric if m.label == "3. 월 영업이익") == f"{float(op):,.0f}원"
    pnl = next(d.value for d in at.dataframe if "영업이익" in d.value.columns)
    assert pnl["매출액"].sum() == rev
    assert pnl["매입원가"].sum() == total_cost
    assert pnl["영업이익"].sum() == op
    for sheet, lookup_col, lookup, result_col, expected_value in [
        ("제안서", "C", "검수(기본)", "E", p), ("제안서", "B", "TEST-A", "E", p),
        ("매출원가산출표", "C", "검수(기본)", "E", 10000)]:
        row = next(c.row for c in wb[sheet][lookup_col] if c.value == lookup)
        assert wb[sheet][f"{result_col}{row}"].value == expected_value
    assert at.session_state.q_warehouse_df.loc[1, "제안단가"] == 40000
    assert at.session_state.q_warehouse_df.loc[0, "비고"] == "가상 검증"
    return {"basis": basis, "target_percent": rate, "include_express": include_exp,
            "include_infra": include_infra, "sga_percent": sga_rate, "warehouse_unit_price": p,
            "actual_unit_gross_percent": (p - 10000) / p * 100,
            "monthly_revenue": float(rev), "direct_cogs": direct, "infra_cogs": infra,
            "sga": sga, "total_cost": float(total_cost), "operating_profit": float(op),
            "operating_margin_percent": float(op / rev * 100), "status": "PASS"}


@pytest.mark.parametrize("rate", [10, 15, 20, 30])
@pytest.mark.parametrize("basis", ["markup", "gross_margin"])
@pytest.mark.parametrize("include_exp,include_infra", [(False, False), (True, False), (False, True), (True, True)])
def test_matrix_controls_pnl_and_all_excel_downloads(captures, rate, basis, include_exp, include_infra):
    at = seeded_app()
    reprice(at, rate, basis)
    at.checkbox(key="in_q_include_express").set_value(include_exp)
    next(c for c in at.checkbox if c.key.startswith("in_q_include_infra_")).set_value(include_infra)
    at.run()
    assert not at.exception
    result = verify_downloads(at, captures, rate, basis, include_exp, include_infra)
    if directory := os.environ.get("REVIEW_RESULTS_DIR"):
        out = Path(directory)
        out.mkdir(parents=True, exist_ok=True)
        (out / f"{basis}_{rate}_{int(include_exp)}_{int(include_infra)}.json").write_text(json.dumps(result, indent=2))


@pytest.mark.parametrize("prefix,button", [("pop_margin_slider_", "btn_recalc_express_margin"),
    ("slider_cogs_exp_margin_", "btn_cogs_recalc_exp"), ("slider_tab6_margin_", None)])
@pytest.mark.parametrize("basis", ["markup", "gross_margin"])
def test_every_express_entry_point(captures, prefix, button, basis):
    at = seeded_app()
    at.radio(key="pricing_basis_tab4").set_value(basis).run()
    control(at, prefix).set_value(20.)
    if button:
        at.button(key=button).click()
    at.run()
    assert not at.exception
    assert at.session_state.q_express_df["판매 제안단가"].tolist() == [price(10000, 20, basis), price(20000, 20, basis)]
    assert all(r.value == basis for r in at.radio if r.key.startswith("pricing_basis_"))


def test_basis_change_preserves_manual_prices_and_existing_values(captures):
    at = seeded_app()
    wh = at.session_state.q_warehouse_df.copy(deep=True)
    exp = at.session_state.q_express_df.copy(deep=True)
    for tab in ["tab4", "tab5", "tab6"]:
        at.radio(key=f"pricing_basis_{tab}").set_value("gross_margin").run()
        assert not at.exception
        pd.testing.assert_frame_equal(at.session_state.q_warehouse_df, wh)
        pd.testing.assert_frame_equal(at.session_state.q_express_df, exp)
        assert at.session_state.q_express_margin == 8.5
        at.radio(key=f"pricing_basis_{tab}").set_value("markup").run()


@pytest.mark.parametrize("key,val", [("sga_rate", 15.), ("sga_rate", 0.), ("inbound_qty", 200), ("express_qty", 300), ("vol_b2c", 200), ("storage_pyeong", 5)])
def test_tab6_inputs_refresh_earlier_downloads_in_same_run(captures, key, val):
    at = seeded_app()
    reprice(at, 20, "gross_margin")
    at.checkbox(key="in_q_include_express").check().run()
    at.number_input(key=f"in_q_{key}").set_value(val).run()
    assert not at.exception
    books = [openpyxl.load_workbook(io.BytesIO(b)) for b in captures[-3:]]
    assert workbook_values(books[0]) == workbook_values(books[1]) == workbook_values(books[2])


def test_zero_express_volume_and_zero_sell_price_do_not_crash(captures):
    at = seeded_app()
    at.checkbox(key="in_q_include_express").check()
    at.number_input(key="in_q_express_qty").set_value(0).run()
    assert not at.exception
    at.session_state.q_express_df["판매 제안단가"] = 0.
    at.number_input(key="in_q_express_qty").set_value(200).run()
    assert not at.exception
    pnl = next(d.value for d in at.dataframe if "영업이익" in d.value.columns)
    assert pnl.loc[pnl["구분"] == "해외특송", "이익률"].iloc[0] == "0.0%"


def test_manual_editors_keep_bidirectional_links_without_implicit_repricing(captures):
    at = seeded_app()
    def edit(prefix, rows):
        key = prefix + str(at.session_state.quote_editor_version)
        at.session_state[key] = {"edited_rows": rows, "added_rows": [], "deleted_rows": []}
        at.run()
        assert not at.exception

    edit("editor_proposal_wh_", {0: {"제안단가": 13333., "비고": "가상 수정 메모", "단위": "PCS"}})
    edit("editor_cogs_wh_", {0: {"매입단가": 11000.}})
    edit("editor_proposal_exp_", {0: {"판매 제안단가": 14444., "비고": "가상 특송 수정"}})
    edit("editor_cogs_exp_", {0: {"매입원가": 12000., "발송비중(%)": 60.}, 1: {"발송비중(%)": 40.}})
    assert at.session_state.q_warehouse_df.loc[0, "제안단가"] == 13333
    assert at.session_state.q_express_df.loc[0, "판매 제안단가"] == 14444
    wh_views = [d.value for d in at.dataframe if "작업 내용" in d.value and "검수(기본)" in d.value["작업 내용"].values]
    assert any("제안단가" in df and df.loc[0, "제안단가"] == 13333 for df in wh_views)
    assert any("매입단가" in df and df.loc[0, "매입단가"] == 11000 for df in wh_views)
    books = [openpyxl.load_workbook(io.BytesIO(b)) for b in captures[-3:]]
    assert workbook_values(books[0]) == workbook_values(books[1]) == workbook_values(books[2])
    proposal, cogs = books[0]["제안서"], books[0]["매출원가산출표"]
    row = next(c.row for c in proposal["C"] if c.value == "검수(기본)")
    assert proposal[f"D{row}"].value == "PCS"
    assert proposal[f"E{row}"].value == 13333
    assert proposal[f"F{row}"].value == "가상 수정 메모"
    row = next(c.row for c in cogs["C"] if c.value == "검수(기본)")
    assert cogs[f"E{row}"].value == 11000
    assert cogs[f"F{row}"].value == "가상 수정 메모"
    row = next(c.row for c in proposal["B"] if c.value == "TEST-A")
    assert proposal[f"E{row}"].value == 14444
    assert proposal[f"F{row}"].value == "가상 특송 수정"


def test_original_quote_excel_format_and_legacy_defaults_are_preserved(captures):
    # Optional immutable upstream source lets reviewers verify byte-independent
    # workbook values, cell styles, merged ranges, sizing and print settings.
    baseline = os.environ.get("BASELINE_APP")
    if not baseline:
        pytest.skip("Set BASELINE_APP to app.py from upstream fb47fc6 for format comparison")
    old = load_helpers(Path(baseline))
    new = load_helpers()
    meta = {"receiver": "가상 검증", "sga_amt": 321650, "sga_rate": 7., "express_qty": 2000}
    books = []
    for h in [old, new]:
        s = h.st.session_state
        books.append(h.build_quote_workbook(meta, s.q_warehouse_df, s.q_express_df, s.q_express_margin,
                     h.get_quote_vol_dict(), include_exp=True, infra_df=s.q_infra_df, include_infra=True))
    assert workbook_values(books[0]) == workbook_values(books[1])
    for a, b in zip(*books):
        assert str(a.merged_cells) == str(b.merged_cells)
        assert {k: dict(v) for k, v in a.column_dimensions.items()} == {k: dict(v) for k, v in b.column_dimensions.items()}
        assert {k: dict(v) for k, v in a.row_dimensions.items()} == {k: dict(v) for k, v in b.row_dimensions.items()}
        assert a.page_setup == b.page_setup
    serialized = []
    for wb in books:
        buffer = io.BytesIO()
        wb.save(buffer)
        serialized.append(zipfile.ZipFile(io.BytesIO(buffer.getvalue())))
    for member in ["xl/styles.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"]:
        assert serialized[0].read(member) == serialized[1].read(member)


def test_hypothetical_ledger_upload_route_keeps_all_six_tabs(monkeypatch, captures):
    payload = "연도/월,계정,적요,금액(현지 통화),거래처명,구분,상세\n2026/10,운송비,가상 검증,10000,가상 공급처,수출,가상\n"
    file = UploadedFile(UploadedFileRec("hypothetical-test-file", "hypothetical_ledger.csv",
                                       "text/csv", payload.encode("utf-8")), FileURLs())
    def upload(*args, **kwargs):
        file.seek(0)
        return file
    monkeypatch.setattr(st, "file_uploader", upload)
    at = new_app()
    assert len(at.tabs) == 6
    assert at.session_state.raw_df["금액(현지 통화)"].tolist() == [10000]
    at.radio(key="pricing_basis_tab6").set_value("gross_margin").run()
    assert not at.exception
    books = [openpyxl.load_workbook(io.BytesIO(b)) for b in captures[-3:]]
    assert workbook_values(books[0]) == workbook_values(books[1]) == workbook_values(books[2])
