"""Independent final file, formula-reference and display-type checks for M1."""
from pathlib import Path
import json
import pickle
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

import pandas as pd

ROOT=Path(__file__).resolve().parents[2]
OUT=ROOT/"pharmacy/analysis/madrid_pilot_20261006"
sys.path.insert(0,str(ROOT/"visualizacion/scripts"))
from prepare_data import sha256

controls=json.loads((OUT/"controls.json").read_text(encoding="utf-8"))
sources=json.loads((OUT/"sources.json").read_text(encoding="utf-8"))
assert all(sha256(ROOT/p)==h for p,h in sources["input_hashes"].items())
production=ROOT/"pharmacy/output/UNIVERSO_FARMACIAS_CATEGORIZADO_2027.xlsx"
assert sha256(production)==sources["production_sha256"]=="cfc28d19c365c658443926d4f89237dd9a712155f80b236e65755449c8cf486b"
book=OUT/"PILOTO_MADRID_DOBLE_CATEGORIZACION_2027.xlsx"
assert sha256(book)==controls["workbook_sha256"]
original_sources=json.loads((OUT/"original_source_controls.json").read_text(encoding="utf-8"))
assert all(sha256(ROOT/item["file"])==item["sha256"] for item in original_sources)

m=pd.read_csv(OUT/"establishments_madrid.csv",sep=";",decimal=",",dtype={"id":str,"sb_code":str,"postal_code":str,"sscc":str})
assert (m.weighted_population-m.pop65-m.pop75-m.pop85).abs().max()<1e-8
NS={"s":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
with zipfile.ZipFile(book) as z:
    wb=ET.fromstring(z.read("xl/workbook.xml"))
    names=[s.attrib["name"] for s in wb.findall("s:sheets/s:sheet",NS)]
    assert len(names)==11
    sheets={name:ET.fromstring(z.read(f"xl/worksheets/sheet{i+1}.xml")) for i,name in enumerate(names)}
    for name,sheet in sheets.items():
        assert sheet.find("s:autoFilter",NS) is not None,name
        pane=sheet.find("s:sheetViews/s:sheetView/s:pane",NS)
        assert pane is not None and pane.attrib["state"]=="frozen",name

    def unpack(sheet):
        result={}
        for cell in sheet.findall("s:sheetData/s:row/s:c",NS):
            v=cell.find("s:v",NS)
            if cell.attrib.get("t")=="inlineStr":
                value="".join(t.text or "" for t in cell.findall("s:is/s:t",NS))
            elif cell.attrib.get("t")=="str":value=v.text if v is not None else ""
            elif v is not None and v.text is not None:value=float(v.text)
            else:value=None
            result[cell.attrib["r"]]=(value,cell.find("s:f",NS),cell.attrib)
        return result

    point_cells=unpack(sheets["Puntos Madrid"])
    cut_cells=unpack(sheets["Cortes €"])
    headings={ref[:-1]:data[0] for ref,data in point_cells.items() if re.fullmatch(r"[A-Z]+4",ref)}
    sb_column=next(c for c,t in headings.items() if t=="sb code")
    sscc_column=next(c for c,t in headings.items() if t=="sscc")
    refs_checked=0
    for ref,(value,formula,attrs) in point_cells.items():
        if formula is None or headings.get(re.sub(r"\d+","",ref)) not in {"Categoría nacional · M1","Categoría CCAA · M1"}:continue
        text=formula.text
        parts=re.findall(r"([A-Z]+\d+)>=\'Cortes €\'!\$([A-Z]+)\$(\d+),\"([ABC])\"",text)
        assert len(parts)==3,(ref,text)
        expected="D"
        for moneyref,col,row,cat in parts:
            if point_cells[moneyref][0]>=cut_cells[col+row][0]:expected=cat;break
        assert expected==value,(ref,expected,value)
        refs_checked+=1
    assert refs_checked==5894
    for row in range(5,2952):
        assert point_cells[f"{sb_column}{row}"][2]["t"]=="inlineStr"
        assert len(point_cells[f"{sb_column}{row}"][0])==6
        section=point_cells[f"{sscc_column}{row}"][0]
        if section:assert len(section)==10

    crm_rows=sheets["CRM Madrid"].findall("s:sheetData/s:row",NS)
    assert len([r for r in crm_rows if int(r.attrib["r"])>=5])==2953
    crm_cells=unpack(sheets["CRM Madrid"])
    with (ROOT/"pharmacy/cache/categorization/universe.pickle").open("rb") as stream:
        universe=pickle.load(stream)
    original=universe["crm"][universe["crm"].establishment_id.isin(m.id)][universe["original_columns"]].reset_index(drop=True)
    compared=0
    for i,values in enumerate(original.itertuples(index=False,name=None),5):
        for j,v in enumerate(values):
            ref=chr(65+j)+str(i)
            expected=None if pd.isna(v) else v
            actual=crm_cells.get(ref,(None,None,{}))[0]
            assert actual==expected,(i,universe["original_columns"][j],type(actual).__name__,type(expected).__name__)
            compared+=1
    assert compared==41342
    pop_label_column=next(c for c,t in headings.items() if t=="Población atribuida 65+")
    assert point_cells[pop_label_column+"5"][0] is not None
    assert sheets["Puntos Madrid"].find("s:conditionalFormatting",NS) is not None
    calc=wb.find("s:calcPr",NS)
    assert calc is not None and calc.attrib["calcMode"]=="auto"

nb=json.loads((OUT/"madrid_analysis.ipynb").read_text(encoding="utf-8"))
assert nb["nbformat"]==4
code=[c for c in nb["cells"] if c["cell_type"]=="code"]
assert len(code)==7 and all(c["execution_count"] for c in code)
assert not any(o["output_type"]=="error" for c in code for o in c["outputs"])
figures=list(OUT.glob("fig-*.png"))
assert len(figures)==5 and all(p.stat().st_size>15000 for p in figures)
result={"production_excel_original_sha256":sources["production_sha256"],"inputs_match_manifest":True,
        "original_rx_and_structure_unchanged":True,"workbook_matches_export_hash":True,"all_11_sheets_filters_and_freezes":True,
        "independently_evaluated_category_formula_references":refs_checked,"codes_are_text":True,"crm_rows":2953,
        "original_crm_cells_compared_to_source":compared,"weighted_population_matches_ages":True,
        "notebook_executed_code_cells":len(code),"notebook_error_outputs":0,"figure_count":len(figures),
        "visual_review":"11 workbook value previews and 5 executed notebook figures inspected; readable labels and units.",
        "xlsx_sha256":sha256(book),"notebook_sha256":sha256(OUT/"madrid_analysis.ipynb"),
        "scope":"Madrid pilot only. National original/app datasets unchanged."}
(OUT/"final_controls.json").write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps(result,ensure_ascii=False))
