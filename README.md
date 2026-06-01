[acmv_comprehensive_calculator_app.py](https://github.com/user-attachments/files/28453743/acmv_comprehensive_calculator_app.py)
# acmv-engineering-calculator
AC&amp;MV engineering calculator platform 
update requirements
import math
from datetime import datetime
from io import BytesIO
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

try:
    from fpdf import FPDF
except Exception:
    FPDF = None

try:
    from docx import Document
except Exception:
    Document = None

# =====================================================
# ACMV COMPREHENSIVE ENGINEERING CALCULATOR APP
# Developed as preliminary calculation software.
# Final engineering selection must be verified with approved drawings,
# project specifications, consultant requirements and manufacturer catalogues.
# =====================================================

st.set_page_config(page_title="ACMV Comprehensive Calculator", layout="wide")
st.title("ACMV Comprehensive Engineering Calculator")
st.caption(
    "50-row engineering schedules for ACMV preliminary design, checking, reporting and submission support. "
    "Use this as a calculation aid only; final selection must be verified by qualified personnel and approved project documents."
)

# =====================================================
# COMMON FORMULA FUNCTIONS
# =====================================================

def safe_float(value, default=0.0):
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def cmh_to_m3s(cmh: float) -> float:
    return cmh / 3600.0 if cmh else 0.0


def m3s_to_cmh(m3s: float) -> float:
    return m3s * 3600.0 if m3s else 0.0


def cmh_to_lps(cmh: float) -> float:
    return cmh / 3.6 if cmh else 0.0


def lps_to_cmh(lps: float) -> float:
    return lps * 3.6 if lps else 0.0


def lps_to_m3s(lps: float) -> float:
    return lps / 1000.0 if lps else 0.0


def kw_to_btuhr(kw: float) -> float:
    return kw * 3412.142


def kw_to_tr(kw: float) -> float:
    return kw / 3.517 if kw else 0.0


def tr_to_kw(tr: float) -> float:
    return tr * 3.517 if tr else 0.0


def rect_area(width_mm: float, height_mm: float) -> float:
    return (width_mm / 1000.0) * (height_mm / 1000.0) if width_mm > 0 and height_mm > 0 else 0.0


def round_area(dia_mm: float) -> float:
    dia_m = dia_mm / 1000.0
    return math.pi * dia_m**2 / 4.0 if dia_m > 0 else 0.0


def hydraulic_diameter_rect(width_mm: float, height_mm: float) -> float:
    w = width_mm / 1000.0
    h = height_mm / 1000.0
    return (2 * w * h / (w + h)) if w > 0 and h > 0 else 0.0


def velocity_pressure_pa(velocity_ms: float, air_density: float = 1.2) -> float:
    return 0.5 * air_density * velocity_ms**2


def water_velocity(flow_lps: float, dia_mm: float) -> float:
    area = round_area(dia_mm)
    return lps_to_m3s(flow_lps) / area if area > 0 else 0.0


def darcy_weisbach_loss_pa_per_m(flow_lps: float, dia_mm: float, friction_factor: float = 0.02, density: float = 1000.0) -> float:
    v = water_velocity(flow_lps, dia_mm)
    d = dia_mm / 1000.0
    return friction_factor * (density * v**2 / 2.0) / d if d > 0 else 0.0


def pa_to_m_head(pa: float, density: float = 1000.0) -> float:
    return pa / (density * 9.81) if density > 0 else 0.0


def clean_text(value):
    return str(value).replace("Δ", "Delta").replace("²", "2").replace("³", "3").replace("–", "-").replace("—", "-")


def make_refs(n=50) -> List[str]:
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    refs = []
    for i in range(n):
        if i < 26:
            refs.append(letters[i])
        else:
            refs.append("A" + letters[i - 26])
    return refs


# =====================================================
# REPORT DOWNLOAD FUNCTIONS
# =====================================================

def excel_download(df: pd.DataFrame, project_info: Dict, summary: Dict, sheet_name="Calculation"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame([project_info]).to_excel(writer, sheet_name="Project Info", index=False)
        pd.DataFrame([summary]).to_excel(writer, sheet_name="Summary", index=False)
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        workbook = writer.book
        for ws in workbook.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                max_len = 12
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        max_len = max(max_len, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_len + 2, 35)
    output.seek(0)
    return output


def pdf_download(df: pd.DataFrame, project_info: Dict, summary: Dict, title="ACMV Calculation Report"):
    if FPDF is None:
        return None

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 8, clean_text(title), ln=True)

    pdf.set_font("Arial", "", 8)
    pdf.cell(0, 5, f"Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}", ln=True)
    pdf.ln(2)

    pdf.set_font("Arial", "B", 9)
    pdf.cell(0, 5, "Project Information", ln=True)
    pdf.set_font("Arial", "", 7)
    for k, v in project_info.items():
        pdf.cell(0, 4, clean_text(f"{k}: {v}"), ln=True)

    pdf.ln(1)
    pdf.set_font("Arial", "B", 9)
    pdf.cell(0, 5, "Summary", ln=True)
    pdf.set_font("Arial", "", 7)
    for k, v in summary.items():
        pdf.cell(0, 4, clean_text(f"{k}: {v}"), ln=True)

    pdf.ln(2)
    display_cols = list(df.columns[:10])
    widths = [20, 42, 22, 24, 24, 24, 24, 24, 35, 45][: len(display_cols)]
    pdf.set_font("Arial", "B", 6)
    for col, w in zip(display_cols, widths):
        pdf.cell(w, 5, clean_text(col)[:18], border=1)
    pdf.ln()

    pdf.set_font("Arial", "", 5)
    for _, row in df.head(50).iterrows():
        for col, w in zip(display_cols, widths):
            pdf.cell(w, 4, clean_text(row.get(col, ""))[:24], border=1)
        pdf.ln()

    data = bytes(pdf.output(dest="S"))
    return BytesIO(data)


def word_download(df: pd.DataFrame, project_info: Dict, summary: Dict, title="ACMV Calculation Report"):
    if Document is None:
        return None

    doc = Document()
    doc.add_heading(title, level=1)
    doc.add_paragraph(f"Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}")

    doc.add_heading("Project Information", level=2)
    for k, v in project_info.items():
        doc.add_paragraph(f"{k}: {v}")

    doc.add_heading("Calculation Summary", level=2)
    for k, v in summary.items():
        doc.add_paragraph(f"{k}: {v}")

    doc.add_heading("Detailed Calculation", level=2)
    display_cols = list(df.columns[:10])
    table = doc.add_table(rows=1, cols=len(display_cols))
    table.style = "Table Grid"
    for i, h in enumerate(display_cols):
        table.rows[0].cells[i].text = str(h)

    for _, row in df.head(50).iterrows():
        cells = table.add_row().cells
        for i, col in enumerate(display_cols):
            cells[i].text = str(row.get(col, ""))

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def show_downloads(df: pd.DataFrame, summary: Dict, filename_prefix: str, title: str):
    st.subheader("Download Reports")
    c1, c2, c3 = st.columns(3)
    c1.download_button(
        "Download Excel",
        data=excel_download(df, st.session_state.project_info, summary, title),
        file_name=f"{filename_prefix}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{filename_prefix}_xlsx",
    )

    pdf_file = pdf_download(df, st.session_state.project_info, summary, title)
    if pdf_file:
        c2.download_button(
            "Download PDF",
            data=pdf_file,
            file_name=f"{filename_prefix}.pdf",
            mime="application/pdf",
            key=f"{filename_prefix}_pdf",
        )
    else:
        c2.warning("Install PDF support: pip install fpdf2")

    word_file = word_download(df, st.session_state.project_info, summary, title)
    if word_file:
        c3.download_button(
            "Download Word",
            data=word_file,
            file_name=f"{filename_prefix}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"{filename_prefix}_docx",
        )
    else:
        c3.warning("Install Word support: pip install python-docx")


def metric_row(summary: Dict):
    keys = list(summary.keys())[:6]
    cols = st.columns(len(keys) if keys else 1)
    for col, key in zip(cols, keys):
        col.metric(key, summary[key])


def editor_with_state(key: str, df: pd.DataFrame, column_config=None, height=580):
    if key not in st.session_state:
        st.session_state[key] = df
    edited = st.data_editor(
        st.session_state[key],
        num_rows="fixed",
        use_container_width=True,
        height=height,
        column_config=column_config or {},
        key=f"editor_{key}",
    )
    st.session_state[key] = edited
    return edited


# =====================================================
# SESSION PROJECT INFO
# =====================================================
if "project_info" not in st.session_state:
    st.session_state.project_info = {
        "Project Name": "",
        "Building / Block": "",
        "System Name": "",
        "Drawing / Rev": "",
        "Submitted To": "",
        "Prepared By": "Muthu",
        "Date": datetime.now().strftime("%d-%b-%Y"),
    }

if "master_summary" not in st.session_state:
    st.session_state.master_summary = {}


def save_summary(module_name: str, summary: Dict):
    st.session_state.master_summary[module_name] = summary


# =====================================================
# SIDEBAR MENU - 10 MAIN OPTIONS
# =====================================================
st.sidebar.title("ACMV Menu")
module = st.sidebar.radio(
    "Choose calculation category",
    [
        "1. Project Setup",
        "2. Duct Static Pressure",
        "3. Airflow / ACH / Ventilation",
        "4. Cooling Load / Capacity",
        "5. Pipe Sizing / Friction",
        "6. Pump Head / Pump Power",
        "7. Fan Selection / Fan Power",
        "8. Electrical FLA / Motor",
        "9. Chilled Water / Coil Flow",
        "10. AHU Special Calculation",
        "11. Pressure Drop Calculation",
        "12. MV Fan Selection",
        "13. Smoke Spill Fan Calculation",
        "14. Fresh Air Fan Calculation",
        "15. Exhaust Fan Calculation",
        "16. AHU Selection",
        "17. FCU Selection",
        "18. Export Excel / PDF / Word",
        "19. Condensate Drain / Final Summary",
    ],
)

refs = make_refs(50)

# =====================================================
# 1 PROJECT SETUP
# =====================================================
if module.startswith("1."):
    st.header("Project Setup")
    st.write("Fill this once. The same information will appear in all Excel, PDF and Word reports.")

    cols = st.columns(2)
    info = {}
    for idx, (k, v) in enumerate(st.session_state.project_info.items()):
        info[k] = cols[idx % 2].text_input(k, value=v, key=f"project_{k}")
    st.session_state.project_info = info
    st.success("Project information saved for this session.")

    st.subheader("Master Summary From All Modules")
    if st.session_state.master_summary:
        rows = []
        for mod, summary in st.session_state.master_summary.items():
            row = {"Module": mod}
            row.update(summary)
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True)
    else:
        st.info("After you calculate each module, the summary will appear here.")

# =====================================================
# 2 DUCT STATIC PRESSURE
# =====================================================
elif module.startswith("2."):
    st.header("Duct Static Pressure Calculation - 50 Rows")
    with st.expander("Design Limits / Warning Settings", expanded=True):
        c1, c2, c3, c4, c5 = st.columns(5)
        max_main_velocity = c1.number_input("Max Main Duct Velocity (m/s)", value=8.0, min_value=0.1, key="duct_max_main")
        max_branch_velocity = c2.number_input("Max Branch / Terminal Velocity (m/s)", value=5.0, min_value=0.1, key="duct_max_branch")
        high_loss_warning = c3.number_input("High Item Loss Warning (Pa)", value=50.0, min_value=0.0, key="duct_high_loss")
        safety_percent = c4.number_input("Safety Allowance (%)", value=10.0, min_value=0.0, key="duct_safety_pct")
        air_density = c5.number_input("Air Density (kg/m³)", value=1.20, min_value=0.1, key="duct_air_density")

    default = pd.DataFrame([
        {
            "Ref": ref,
            "Description": "Main duct from fan" if i == 0 else "",
            "Type": "Straight" if i == 0 else "Blank",
            "Duct Shape": "Rectangular",
            "Category": "Main",
            "Airflow CMH": 5000.0 if i == 0 else 0.0,
            "Width mm": 800.0 if i == 0 else 0.0,
            "Height/Dia mm": 500.0 if i == 0 else 0.0,
            "Length m": 10.0 if i == 0 else 0.0,
            "Pa/m": 1.0 if i == 0 else 0.0,
            "K Factor": 0.0,
            "Fixed Loss Pa": 0.0,
        }
        for i, ref in enumerate(refs)
    ])

    edited = editor_with_state(
        "duct_df_v2",
        default,
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=["Blank", "Straight", "Fitting", "Fixed Loss"]),
            "Duct Shape": st.column_config.SelectboxColumn("Duct Shape", options=["Rectangular", "Round"]),
            "Category": st.column_config.SelectboxColumn("Category", options=["Main", "Branch", "Terminal", "Fresh Air", "Exhaust"]),
        },
    )

    results = []
    for _, row in edited.iterrows():
        item_type = row.get("Type", "Blank")
        airflow = safe_float(row.get("Airflow CMH"))
        width = safe_float(row.get("Width mm"))
        hd = safe_float(row.get("Height/Dia mm"))
        length = safe_float(row.get("Length m"))
        pa_m = safe_float(row.get("Pa/m"))
        k = safe_float(row.get("K Factor"))
        fixed = safe_float(row.get("Fixed Loss Pa"))
        shape = row.get("Duct Shape", "Rectangular")
        category = row.get("Category", "Main")
        area = round_area(hd) if shape == "Round" else rect_area(width, hd)
        q = cmh_to_m3s(airflow)
        v = q / area if area > 0 else 0.0
        vp = velocity_pressure_pa(v, air_density)
        loss = pa_m * length if item_type == "Straight" else k * vp if item_type == "Fitting" else fixed if item_type == "Fixed Loss" else 0.0
        remarks = []
        if item_type != "Blank":
            if category == "Main" and v > max_main_velocity:
                remarks.append("High main duct velocity")
            if category != "Main" and v > max_branch_velocity:
                remarks.append("High branch/terminal velocity")
            if loss > high_loss_warning:
                remarks.append("High pressure loss")
            if airflow > 0 and area <= 0:
                remarks.append("Check duct size")
        results.append({**row.to_dict(), "Area m²": round(area, 4), "Airflow L/s": round(cmh_to_lps(airflow), 2), "Velocity m/s": round(v, 2), "Velocity Pressure Pa": round(vp, 2), "Pressure Loss Pa": round(loss, 2), "Status / Remarks": "; ".join(remarks) if remarks else ("OK" if item_type != "Blank" else "")})

    result_df = pd.DataFrame(results)
    active = result_df[result_df["Type"] != "Blank"]
    subtotal = active["Pressure Loss Pa"].sum() if not active.empty else 0.0
    safety = subtotal * safety_percent / 100.0
    total = subtotal + safety
    summary = {
        "Subtotal ESP Pa": round(subtotal, 2),
        "Safety Pa": round(safety, 2),
        "Final ESP Pa": round(total, 2),
        "Max Velocity m/s": round(active["Velocity m/s"].max() if not active.empty else 0, 2),
        "Active Items": int(len(active)),
    }
    save_summary("Duct Static Pressure", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "duct_static_pressure_calculation", "Duct Static Pressure Calculation")

# =====================================================
# 3 AIRFLOW / ACH / VENTILATION
# =====================================================
elif module.startswith("3."):
    st.header("Airflow / ACH / Ventilation Schedule - 50 Rows")
    st.write("Use this for room airflow, ACH, fresh air, exhaust and make-up air checking.")
    c1, c2 = st.columns(2)
    min_ach_default = c1.number_input("Default Minimum ACH Warning", value=6.0, min_value=0.0, key="ach_min")
    safety_pct = c2.number_input("Airflow Safety Allowance (%)", value=10.0, min_value=0.0, key="ach_safety")

    default = pd.DataFrame([
        {
            "Ref": ref,
            "Room / Area": "Room 1" if i == 0 else "",
            "Calculation Type": "ACH from Room Volume" if i == 0 else "Blank",
            "Length m": 10.0 if i == 0 else 0.0,
            "Width m": 8.0 if i == 0 else 0.0,
            "Height m": 3.0 if i == 0 else 0.0,
            "Required ACH": 6.0 if i == 0 else 0.0,
            "Known Airflow CMH": 0.0,
            "Occupants": 0.0,
            "L/s per Person": 10.0,
            "Exhaust CMH": 0.0,
        }
        for i, ref in enumerate(refs)
    ])
    edited = editor_with_state(
        "airflow_df_v2",
        default,
        column_config={"Calculation Type": st.column_config.SelectboxColumn("Calculation Type", options=["Blank", "ACH from Room Volume", "Check Existing ACH", "Outdoor Air by Occupancy", "Exhaust / Make-up Air"])}
    )

    rows = []
    for _, row in edited.iterrows():
        typ = row.get("Calculation Type", "Blank")
        volume = safe_float(row.get("Length m")) * safe_float(row.get("Width m")) * safe_float(row.get("Height m"))
        req_ach = safe_float(row.get("Required ACH"))
        known_cmh = safe_float(row.get("Known Airflow CMH"))
        occupants = safe_float(row.get("Occupants"))
        lps_person = safe_float(row.get("L/s per Person"))
        exhaust_cmh = safe_float(row.get("Exhaust CMH"))
        required_cmh = 0.0
        actual_ach = 0.0
        if typ == "ACH from Room Volume":
            required_cmh = volume * req_ach
            actual_ach = req_ach
        elif typ == "Check Existing ACH":
            required_cmh = known_cmh
            actual_ach = known_cmh / volume if volume > 0 else 0.0
        elif typ == "Outdoor Air by Occupancy":
            required_cmh = lps_to_cmh(occupants * lps_person)
            actual_ach = required_cmh / volume if volume > 0 else 0.0
        elif typ == "Exhaust / Make-up Air":
            required_cmh = exhaust_cmh
            actual_ach = exhaust_cmh / volume if volume > 0 else 0.0
        final_cmh = required_cmh * (1 + safety_pct / 100.0) if typ != "Blank" else 0.0
        remarks = []
        if typ != "Blank":
            if volume <= 0:
                remarks.append("Check room dimensions")
            if typ == "Check Existing ACH" and actual_ach < min_ach_default:
                remarks.append("ACH below warning setting")
        rows.append({**row.to_dict(), "Room Volume m³": round(volume, 2), "Calculated Airflow CMH": round(required_cmh, 2), "Calculated Airflow L/s": round(cmh_to_lps(required_cmh), 2), "Actual / Design ACH": round(actual_ach, 2), "Final CMH with Safety": round(final_cmh, 2), "Status / Remarks": "; ".join(remarks) if remarks else ("OK" if typ != "Blank" else "")})
    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Calculation Type"] != "Blank"]
    summary = {"Total CMH": round(active["Calculated Airflow CMH"].sum(), 2), "Total L/s": round(active["Calculated Airflow L/s"].sum(), 2), "Total Final CMH": round(active["Final CMH with Safety"].sum(), 2), "No. of Areas": int(len(active))}
    save_summary("Airflow / ACH", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "airflow_ach_ventilation_calculation", "Airflow ACH Ventilation Calculation")

# =====================================================
# 4 COOLING LOAD / CAPACITY
# =====================================================
elif module.startswith("4."):
    st.header("Cooling Load / Capacity Schedule - 50 Rows")
    st.write("Quick preliminary sensible + latent style load build-up. For final design, use project-approved heat load calculation method.")
    safety_pct = st.number_input("Cooling Capacity Safety Allowance (%)", value=10.0, min_value=0.0, key="cooling_safety")
    default = pd.DataFrame([
        {"Ref": ref, "Area / Zone": "Zone 1" if i == 0 else "", "Type": "Area Method" if i == 0 else "Blank", "Area m²": 50.0 if i == 0 else 0.0, "W/m²": 150.0, "People": 5.0 if i == 0 else 0.0, "W/Person": 120.0, "Equipment W": 1000.0 if i == 0 else 0.0, "Lighting W": 500.0 if i == 0 else 0.0, "Fresh Air CMH": 300.0 if i == 0 else 0.0, "OA Delta h kJ/kg": 20.0}
        for i, ref in enumerate(refs)
    ])
    edited = editor_with_state("cooling_df_v2", default, column_config={"Type": st.column_config.SelectboxColumn("Type", options=["Blank", "Area Method", "Detailed Internal Load", "Fresh Air Load Only", "Fixed kW"])} )
    rows = []
    for _, row in edited.iterrows():
        typ = row.get("Type", "Blank")
        area_kw = safe_float(row.get("Area m²")) * safe_float(row.get("W/m²")) / 1000.0
        people_kw = safe_float(row.get("People")) * safe_float(row.get("W/Person")) / 1000.0
        equip_kw = safe_float(row.get("Equipment W")) / 1000.0
        light_kw = safe_float(row.get("Lighting W")) / 1000.0
        # Outdoor air approximate total load: m_dot x delta h. density approx 1.2 kg/m3.
        oa_kw = cmh_to_m3s(safe_float(row.get("Fresh Air CMH"))) * 1.2 * safe_float(row.get("OA Delta h kJ/kg"))
        fixed_kw = safe_float(row.get("Area m²")) if typ == "Fixed kW" else 0.0
        if typ == "Area Method": total_kw = area_kw + oa_kw
        elif typ == "Detailed Internal Load": total_kw = people_kw + equip_kw + light_kw + oa_kw
        elif typ == "Fresh Air Load Only": total_kw = oa_kw
        elif typ == "Fixed kW": total_kw = fixed_kw
        else: total_kw = 0.0
        final_kw = total_kw * (1 + safety_pct / 100.0) if typ != "Blank" else 0.0
        rows.append({**row.to_dict(), "Area Load kW": round(area_kw, 2), "People Load kW": round(people_kw, 2), "Equipment kW": round(equip_kw, 2), "Lighting kW": round(light_kw, 2), "Fresh Air Load kW": round(oa_kw, 2), "Total Cooling kW": round(total_kw, 2), "Total TR": round(kw_to_tr(total_kw), 2), "Final kW with Safety": round(final_kw, 2), "Final TR": round(kw_to_tr(final_kw), 2), "Status / Remarks": "OK" if typ != "Blank" else ""})
    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Type"] != "Blank"]
    summary = {"Total kW": round(active["Total Cooling kW"].sum(), 2), "Total TR": round(active["Total TR"].sum(), 2), "Final kW": round(active["Final kW with Safety"].sum(), 2), "Final TR": round(active["Final TR"].sum(), 2)}
    save_summary("Cooling Load", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "cooling_load_capacity_calculation", "Cooling Load Capacity Calculation")

# =====================================================
# 5 PIPE SIZING / FRICTION
# =====================================================
elif module.startswith("5."):
    st.header("Pipe Sizing / Friction Schedule - 50 Rows")
    c1, c2, c3 = st.columns(3)
    max_vel = c1.number_input("Max Pipe Velocity Warning (m/s)", value=2.5, min_value=0.1, key="pipe_max_vel")
    friction_factor = c2.number_input("Darcy Friction Factor", value=0.02, min_value=0.001, key="pipe_ff")
    fitting_safety = c3.number_input("Additional Safety / Misc (%)", value=10.0, min_value=0.0, key="pipe_safety")
    default = pd.DataFrame([
        {"Ref": ref, "Description": "CHW main" if i == 0 else "", "Type": "Straight Pipe" if i == 0 else "Blank", "Flow L/s": 2.0 if i == 0 else 0.0, "Pipe ID mm": 50.0 if i == 0 else 0.0, "Length m": 30.0 if i == 0 else 0.0, "Fitting K": 0.0, "No. of Fittings": 0.0, "Fixed Loss mH2O": 0.0}
        for i, ref in enumerate(refs)
    ])
    edited = editor_with_state("pipe_df_v2", default, column_config={"Type": st.column_config.SelectboxColumn("Type", options=["Blank", "Straight Pipe", "Fitting K Loss", "Fixed Loss"])} )
    rows = []
    for _, row in edited.iterrows():
        typ = row.get("Type", "Blank")
        flow = safe_float(row.get("Flow L/s")); dia = safe_float(row.get("Pipe ID mm")); length = safe_float(row.get("Length m")); k = safe_float(row.get("Fitting K")); n = safe_float(row.get("No. of Fittings")); fixed_m = safe_float(row.get("Fixed Loss mH2O"))
        v = water_velocity(flow, dia)
        pa_m = darcy_weisbach_loss_pa_per_m(flow, dia, friction_factor)
        m_per_m = pa_to_m_head(pa_m)
        straight_loss = m_per_m * length
        fitting_loss = pa_to_m_head(k * n * 1000.0 * v**2 / 2.0)
        loss = straight_loss if typ == "Straight Pipe" else fitting_loss if typ == "Fitting K Loss" else fixed_m if typ == "Fixed Loss" else 0.0
        remarks = []
        if typ != "Blank":
            if v > max_vel: remarks.append("High pipe velocity")
            if dia <= 0 and flow > 0: remarks.append("Check pipe ID")
        rows.append({**row.to_dict(), "Velocity m/s": round(v, 2), "Friction Pa/m": round(pa_m, 2), "Friction mH2O/m": round(m_per_m, 4), "Straight Loss mH2O": round(straight_loss, 3), "Fitting Loss mH2O": round(fitting_loss, 3), "Total Loss mH2O": round(loss, 3), "Status / Remarks": "; ".join(remarks) if remarks else ("OK" if typ != "Blank" else "")})
    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Type"] != "Blank"]
    subtotal = active["Total Loss mH2O"].sum() if not active.empty else 0.0
    final = subtotal * (1 + fitting_safety / 100.0)
    summary = {"Subtotal Head m": round(subtotal, 2), "Final Head m": round(final, 2), "Max Velocity m/s": round(active["Velocity m/s"].max() if not active.empty else 0, 2), "Active Items": int(len(active))}
    save_summary("Pipe Sizing / Friction", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "pipe_sizing_friction_calculation", "Pipe Sizing Friction Calculation")

# =====================================================
# 6 PUMP HEAD / PUMP POWER
# =====================================================
elif module.startswith("6."):
    st.header("Pump Head / Pump Power Schedule - 50 Rows")
    c1, c2 = st.columns(2)
    pump_eff = c1.number_input("Pump Efficiency (%)", value=60.0, min_value=1.0, max_value=100.0, key="pump_eff")
    motor_margin = c2.number_input("Motor Margin (%)", value=15.0, min_value=0.0, key="pump_motor_margin")
    default = pd.DataFrame([
        {"Ref": ref, "Description": "Circuit 1" if i == 0 else "", "Type": "Pipe / Fitting Head" if i == 0 else "Blank", "Flow L/s": 2.0 if i == 0 else 0.0, "Head mH2O": 20.0 if i == 0 else 0.0, "Quantity": 1.0, "Diversity %": 100.0}
        for i, ref in enumerate(refs)
    ])
    edited = editor_with_state("pump_df_v2", default, column_config={"Type": st.column_config.SelectboxColumn("Type", options=["Blank", "Pipe / Fitting Head", "Equipment Head", "Static Head", "Control Valve Allowance", "Fixed Head"])} )
    rows=[]
    for _, row in edited.iterrows():
        typ=row.get("Type","Blank"); flow=safe_float(row.get("Flow L/s")); head=safe_float(row.get("Head mH2O")); qty=safe_float(row.get("Quantity"),1); div=safe_float(row.get("Diversity %"),100)/100
        eq_head=head*qty*div if typ!="Blank" else 0.0
        rows.append({**row.to_dict(), "Equivalent Head mH2O": round(eq_head,2), "Status / Remarks": "OK" if typ!="Blank" else ""})
    result_df=pd.DataFrame(rows); active=result_df[result_df["Type"]!="Blank"]
    total_head=active["Equivalent Head mH2O"].sum() if not active.empty else 0.0
    design_flow=active["Flow L/s"].max() if not active.empty else 0.0
    hyd_kw=1000*9.81*lps_to_m3s(design_flow)*total_head/1000
    shaft_kw=hyd_kw/(pump_eff/100) if pump_eff>0 else 0.0
    motor_kw=shaft_kw*(1+motor_margin/100)
    summary={"Design Flow L/s":round(design_flow,2),"Total Head m":round(total_head,2),"Hydraulic kW":round(hyd_kw,2),"Shaft kW":round(shaft_kw,2),"Motor kW":round(motor_kw,2)}
    save_summary("Pump Head / Power", summary)
    metric_row(summary)
    st.dataframe(result_df,use_container_width=True,height=500)
    show_downloads(result_df, summary, "pump_head_power_calculation", "Pump Head Pump Power Calculation")

# =====================================================
# 7 FAN SELECTION / FAN POWER
# =====================================================
elif module.startswith("7."):
    st.header("Fan Selection / Fan Power Schedule - 50 Rows")
    c1,c2=st.columns(2)
    fan_eff=c1.number_input("Fan Total Efficiency (%)",value=55.0,min_value=1.0,max_value=100.0,key="fan_eff")
    motor_margin=c2.number_input("Motor Margin (%)",value=15.0,min_value=0.0,key="fan_motor_margin")
    default=pd.DataFrame([{ "Ref":ref,"Fan Tag":"FAF-1" if i==0 else "","Type":"Supply Fan" if i==0 else "Blank","Airflow CMH":5000.0 if i==0 else 0.0,"ESP Pa":500.0 if i==0 else 0.0,"System Effect Pa":50.0 if i==0 else 0.0,"Filter Dirty Allowance Pa":50.0 if i==0 else 0.0,"Quantity":1.0} for i,ref in enumerate(refs)])
    edited=editor_with_state("fan_df_v2",default,column_config={"Type":st.column_config.SelectboxColumn("Type",options=["Blank","Supply Fan","Return Fan","Exhaust Fan","Fresh Air Fan","Smoke Purge Fan"])} )
    rows=[]
    for _,row in edited.iterrows():
        typ=row.get("Type","Blank"); cmh=safe_float(row.get("Airflow CMH")); esp=safe_float(row.get("ESP Pa")); se=safe_float(row.get("System Effect Pa")); filt=safe_float(row.get("Filter Dirty Allowance Pa")); qty=safe_float(row.get("Quantity"),1)
        total_sp=esp+se+filt if typ!="Blank" else 0.0
        kw=(cmh_to_m3s(cmh)*total_sp)/(fan_eff/100)/1000 if fan_eff>0 and typ!="Blank" else 0.0
        motor_kw=kw*(1+motor_margin/100)
        rows.append({**row.to_dict(),"Total SP Pa":round(total_sp,2),"Fan Power kW Each":round(kw,2),"Motor kW Each":round(motor_kw,2),"Total Airflow CMH":round(cmh*qty,2),"Total Motor kW":round(motor_kw*qty,2),"Status / Remarks":"OK" if typ!="Blank" else ""})
    result_df=pd.DataFrame(rows); active=result_df[result_df["Type"]!="Blank"]
    summary={"Total Airflow CMH":round(active["Total Airflow CMH"].sum(),2),"Max SP Pa":round(active["Total SP Pa"].max() if not active.empty else 0,2),"Total Fan kW":round(active["Fan Power kW Each"].sum(),2),"Total Motor kW":round(active["Total Motor kW"].sum(),2)}
    save_summary("Fan Selection",summary); metric_row(summary); st.dataframe(result_df,use_container_width=True,height=500); show_downloads(result_df,summary,"fan_selection_power_calculation","Fan Selection Power Calculation")

# =====================================================
# 8 ELECTRICAL FLA / MOTOR
# =====================================================
elif module.startswith("8."):
    st.header("Electrical FLA / Motor Schedule - 50 Rows")
    default=pd.DataFrame([{ "Ref":ref,"Equipment Tag":"MOTOR-1" if i==0 else "","Phase":"3 Phase","Motor kW":5.5 if i==0 else 0.0,"Voltage V":415.0,"Power Factor":0.85,"Efficiency %":90.0,"Quantity":1.0,"Starter Type":"DOL"} for i,ref in enumerate(refs)])
    edited=editor_with_state("electrical_df_v2",default,column_config={"Phase":st.column_config.SelectboxColumn("Phase",options=["1 Phase","3 Phase"]),"Starter Type":st.column_config.SelectboxColumn("Starter Type",options=["DOL","Star-Delta","VSD","Soft Starter"])} )
    rows=[]
    for _,row in edited.iterrows():
        kw=safe_float(row.get("Motor kW")); volt=safe_float(row.get("Voltage V")); pf=safe_float(row.get("Power Factor")); eff=safe_float(row.get("Efficiency %")); qty=safe_float(row.get("Quantity"),1); phase=row.get("Phase","3 Phase")
        if phase=="3 Phase": fla=(kw*1000)/(math.sqrt(3)*volt*pf*(eff/100)) if volt>0 and pf>0 and eff>0 else 0
        else: fla=(kw*1000)/(volt*pf*(eff/100)) if volt>0 and pf>0 and eff>0 else 0
        rec_breaker=fla*1.25
        rows.append({**row.to_dict(),"FLA A Each":round(fla,2),"Total FLA A":round(fla*qty,2),"Suggested Breaker Min A":round(rec_breaker,2),"Status / Remarks":"OK" if kw>0 else ""})
    result_df=pd.DataFrame(rows); active=result_df[result_df["Motor kW"]>0]
    summary={"Total Connected kW":round((active["Motor kW"]*active["Quantity"]).sum() if not active.empty else 0,2),"Total FLA A":round(active["Total FLA A"].sum() if not active.empty else 0,2),"Max FLA A":round(active["FLA A Each"].max() if not active.empty else 0,2),"No. of Motors":int(len(active))}
    save_summary("Electrical FLA",summary); metric_row(summary); st.dataframe(result_df,use_container_width=True,height=500); show_downloads(result_df,summary,"electrical_fla_motor_calculation","Electrical FLA Motor Calculation")

# =====================================================
# 9 CHILLED WATER / COIL FLOW
# =====================================================
elif module.startswith("9."):
    st.header("Chilled Water / Coil Flow Schedule - 50 Rows")
    st.write("Formula used: Water flow L/s = kW / (4.186 × Delta T).")
    default=pd.DataFrame([{ "Ref":ref,"Equipment Tag":"AHU-1" if i==0 else "","Type":"AHU Coil" if i==0 else "Blank","Cooling Capacity kW":37.5 if i==0 else 0.0,"CHW Supply °C":7.0,"CHW Return °C":12.0,"Valve Pressure Drop kPa":30.0,"Coil Pressure Drop kPa":40.0,"Quantity":1.0} for i,ref in enumerate(refs)])
    edited=editor_with_state("chw_df_v2",default,column_config={"Type":st.column_config.SelectboxColumn("Type",options=["Blank","AHU Coil","FCU Coil","PAHU Coil","Process Load","Fixed Flow"])} )
    rows=[]
    for _,row in edited.iterrows():
        typ=row.get("Type","Blank"); kw=safe_float(row.get("Cooling Capacity kW")); ts=safe_float(row.get("CHW Supply °C")); tr=safe_float(row.get("CHW Return °C")); valve=safe_float(row.get("Valve Pressure Drop kPa")); coil=safe_float(row.get("Coil Pressure Drop kPa")); qty=safe_float(row.get("Quantity"),1)
        dt=tr-ts
        flow=kw/(4.186*dt) if dt>0 and typ!="Blank" else 0.0
        head=pa_to_m_head((valve+coil)*1000)
        rows.append({**row.to_dict(),"Delta T °C":round(dt,2),"Flow L/s Each":round(flow,3),"Flow CMH Each":round(lps_to_cmh(flow),2),"Total Flow L/s":round(flow*qty,3),"Coil+Valve Head m":round(head,2),"Status / Remarks":"OK" if typ!="Blank" and dt>0 else ("Check Delta T" if typ!="Blank" else "")})
    result_df=pd.DataFrame(rows); active=result_df[result_df["Type"]!="Blank"]
    summary={"Total Cooling kW":round((active["Cooling Capacity kW"]*active["Quantity"]).sum() if not active.empty else 0,2),"Total Flow L/s":round(active["Total Flow L/s"].sum() if not active.empty else 0,2),"Total Flow CMH":round(lps_to_cmh(active["Total Flow L/s"].sum() if not active.empty else 0),2),"Max Head m":round(active["Coil+Valve Head m"].max() if not active.empty else 0,2)}
    save_summary("CHW / Coil Flow",summary); metric_row(summary); st.dataframe(result_df,use_container_width=True,height=500); show_downloads(result_df,summary,"chilled_water_coil_flow_calculation","Chilled Water Coil Flow Calculation")

# =====================================================
# 10 AHU SPECIAL CALCULATION
# =====================================================
elif module.startswith("10."):
    st.header("AHU Special Calculation")

    default = pd.DataFrame([{
        "Ref": ref,
        "Fan No": "",
        "SPL Flow CMH": 0.0,
        "Revised Flow CMH": 0.0,
        "No. of Transition": 1.0,
        "Shaft Length m": 0.0,
        "Shaft Width m": 0.0,
        "Shaft Height m": 0.0,
        "K-Factor": 0.00444,
        "Shock Loss Factor": 1.2,
        "Air Density kg/m3": 1.2
    } for ref in refs])

    edited = editor_with_state("ahu_special_df_v1", default)

    rows = []
    for _, row in edited.iterrows():
        revised_flow = safe_float(row.get("Revised Flow CMH"))
        transition_no = safe_float(row.get("No. of Transition"))
        length = safe_float(row.get("Shaft Length m"))
        width = safe_float(row.get("Shaft Width m"))
        height = safe_float(row.get("Shaft Height m"))
        k_factor = safe_float(row.get("K-Factor"))
        shock_factor = safe_float(row.get("Shock Loss Factor"))
        air_density = safe_float(row.get("Air Density kg/m3"))

        airflow_m3s = revised_flow / 3600 if revised_flow else 0
        perimeter = (width + height) * 2 if width and height else 0
        area = width * height if width and height else 0
        ratio = height / width if width else 0

        if area > 0:
            straight_r = (k_factor * length * perimeter * air_density) / ((area ** 3) * 1.2)
            bend_r = (shock_factor * air_density) / (2 * (area ** 2))
            total_r = straight_r + (transition_no * bend_r)
            static_loss = math.ceil(total_r * airflow_m3s * airflow_m3s) * 1.5
        else:
            straight_r = bend_r = total_r = static_loss = 0

        rows.append({
            **row.to_dict(),
            "Air Flow m3/s": round(airflow_m3s, 3),
            "Perimeter m": round(perimeter, 3),
            "Area m2": round(area, 3),
            "H/W Ratio": round(ratio, 3),
            "Straight Run R": round(straight_r, 6),
            "Bend R": round(bend_r, 6),
            "Total R": round(total_r, 6),
            "Static Pressure Loss Pa": round(static_loss, 2)
        })

    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Revised Flow CMH"] > 0]

    summary = {
        "Total AHU Static Pressure Pa": round(active["Static Pressure Loss Pa"].sum() * 1.5 if not active.empty else 0, 2),
        "No. of Active Rows": len(active)
    }

    save_summary("AHU Special Calculation", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "ahu_special_calculation", "AHU Special Calculation")


# =====================================================
# 11 PRESSURE DROP CALCULATION
# =====================================================
elif module.startswith("11."):
    st.header("Pressure Drop Calculation - 50 Rows")

    default = pd.DataFrame([{
        "Ref": ref,
        "System Element": "",
        "Equipment Label": "",
        "CMH": 0.0,
        "Outlet Area m2": 0.0,
        "Length m / No.": 0.0,
        "Loss Coefficient K": 0.0,
        "Air Density kg/m3": 1.2
    } for ref in refs])

    edited = editor_with_state("pressure_drop_df_v1", default)

    rows = []
    for _, row in edited.iterrows():
        cmh = safe_float(row.get("CMH"))
        area = safe_float(row.get("Outlet Area m2"))
        length_no = safe_float(row.get("Length m / No."))
        k = safe_float(row.get("Loss Coefficient K"))
        density = safe_float(row.get("Air Density kg/m3"))

        q_m3s = cmh / 3600 if cmh else 0
        velocity = q_m3s / area if area > 0 else 0
        velocity_pressure = density * velocity * velocity / 2
        pressure_drop = velocity_pressure * k * length_no

        rows.append({
            **row.to_dict(),
            "Air Flow Q m3/s": round(q_m3s, 3),
            "Velocity m/s": round(velocity, 3),
            "Velocity Pressure Pa": round(velocity_pressure, 2),
            "Pressure Drop Pa": round(pressure_drop, 2)
        })

    result_df = pd.DataFrame(rows)
    active = result_df[result_df["CMH"] > 0]

    summary = {
        "Total Pressure Drop Pa": round(active["Pressure Drop Pa"].sum() if not active.empty else 0, 2),
        "No. of Active Rows": len(active)
    }

    save_summary("Pressure Drop Calculation", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "pressure_drop_calculation", "Pressure Drop Calculation")


# =====================================================
# 12 MV FAN SELECTION
# =====================================================
elif module.startswith("12."):
    st.header("MV Fan Selection - 50 Rows")

    default = pd.DataFrame([{
        "Ref": ref,
        "Fan Tag": "",
        "Location": "",
        "Airflow CMH": 0.0,
        "External Static Pressure Pa": 0.0,
        "Fan Efficiency %": 60.0,
        "Motor Safety Margin %": 15.0
    } for ref in refs])

    edited = editor_with_state("mv_fan_df_v1", default)

    rows = []
    for _, row in edited.iterrows():
        cmh = safe_float(row.get("Airflow CMH"))
        esp = safe_float(row.get("External Static Pressure Pa"))
        eff = safe_float(row.get("Fan Efficiency %"))
        margin = safe_float(row.get("Motor Safety Margin %"))

        q = cmh / 3600 if cmh else 0
        power_kw = (q * esp) / ((eff / 100) * 1000) if eff > 0 else 0
        motor_kw = power_kw * (1 + margin / 100)

        rows.append({
            **row.to_dict(),
            "Airflow m3/s": round(q, 3),
            "Fan Power kW": round(power_kw, 3),
            "Recommended Motor kW": round(motor_kw, 3)
        })

    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Airflow CMH"] > 0]

    summary = {
        "Total Airflow CMH": round(active["Airflow CMH"].sum() if not active.empty else 0, 2),
        "Total Fan Power kW": round(active["Fan Power kW"].sum() if not active.empty else 0, 2)
    }

    save_summary("MV Fan Selection", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "mv_fan_selection", "MV Fan Selection")


# =====================================================
# 13 SMOKE SPILL FAN CALCULATION
# =====================================================
elif module.startswith("13."):
    st.header("Smoke Spill Fan Calculation - 50 Rows")

    default = pd.DataFrame([{
        "Ref": ref,
        "Area / Zone": "",
        "Length m": 0.0,
        "Width m": 0.0,
        "Smoke Exhaust Rate ACH": 10.0,
        "Room Height m": 3.0,
        "Safety Factor %": 10.0
    } for ref in refs])

    edited = editor_with_state("smoke_spill_df_v1", default)

    rows = []
    for _, row in edited.iterrows():
        length = safe_float(row.get("Length m"))
        width = safe_float(row.get("Width m"))
        height = safe_float(row.get("Room Height m"))
        ach = safe_float(row.get("Smoke Exhaust Rate ACH"))
        safety = safe_float(row.get("Safety Factor %"))

        volume = length * width * height
        airflow = volume * ach
        final_airflow = airflow * (1 + safety / 100)

        rows.append({
            **row.to_dict(),
            "Room Volume m3": round(volume, 2),
            "Required Airflow CMH": round(airflow, 2),
            "Final Airflow with Safety CMH": round(final_airflow, 2)
        })

    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Final Airflow with Safety CMH"] > 0]

    summary = {
        "Total Smoke Spill Airflow CMH": round(active["Final Airflow with Safety CMH"].sum() if not active.empty else 0, 2)
    }

    save_summary("Smoke Spill Fan Calculation", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "smoke_spill_fan_calculation", "Smoke Spill Fan Calculation")


# =====================================================
# 14 FRESH AIR FAN CALCULATION
# =====================================================
elif module.startswith("14."):
    st.header("Fresh Air Fan Calculation - 50 Rows")

    default = pd.DataFrame([{
        "Ref": ref,
        "Room / Area": "",
        "No. of Persons": 0.0,
        "Fresh Air per Person L/s": 10.0,
        "Area m2": 0.0,
        "Fresh Air per Area L/s/m2": 0.3,
        "Safety Factor %": 10.0
    } for ref in refs])

    edited = editor_with_state("fresh_air_df_v1", default)

    rows = []
    for _, row in edited.iterrows():
        people = safe_float(row.get("No. of Persons"))
        fa_person = safe_float(row.get("Fresh Air per Person L/s"))
        area = safe_float(row.get("Area m2"))
        fa_area = safe_float(row.get("Fresh Air per Area L/s/m2"))
        safety = safe_float(row.get("Safety Factor %"))

        lps = (people * fa_person) + (area * fa_area)
        cmh = lps * 3.6
        final_cmh = cmh * (1 + safety / 100)

        rows.append({
            **row.to_dict(),
            "Fresh Air L/s": round(lps, 2),
            "Fresh Air CMH": round(cmh, 2),
            "Final Fresh Air CMH": round(final_cmh, 2)
        })

    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Final Fresh Air CMH"] > 0]

    summary = {
        "Total Fresh Air CMH": round(active["Final Fresh Air CMH"].sum() if not active.empty else 0, 2)
    }

    save_summary("Fresh Air Fan Calculation", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "fresh_air_fan_calculation", "Fresh Air Fan Calculation")


# =====================================================
# 15 EXHAUST FAN CALCULATION
# =====================================================
elif module.startswith("15."):
    st.header("Exhaust Fan Calculation - 50 Rows")

    default = pd.DataFrame([{
        "Ref": ref,
        "Room / Area": "",
        "Length m": 0.0,
        "Width m": 0.0,
        "Height m": 3.0,
        "Required ACH": 10.0,
        "Safety Factor %": 10.0
    } for ref in refs])

    edited = editor_with_state("exhaust_fan_df_v1", default)

    rows = []
    for _, row in edited.iterrows():
        length = safe_float(row.get("Length m"))
        width = safe_float(row.get("Width m"))
        height = safe_float(row.get("Height m"))
        ach = safe_float(row.get("Required ACH"))
        safety = safe_float(row.get("Safety Factor %"))

        volume = length * width * height
        cmh = volume * ach
        final_cmh = cmh * (1 + safety / 100)

        rows.append({
            **row.to_dict(),
            "Room Volume m3": round(volume, 2),
            "Exhaust Airflow CMH": round(cmh, 2),
            "Final Exhaust Airflow CMH": round(final_cmh, 2)
        })

    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Final Exhaust Airflow CMH"] > 0]

    summary = {
        "Total Exhaust Airflow CMH": round(active["Final Exhaust Airflow CMH"].sum() if not active.empty else 0, 2)
    }

    save_summary("Exhaust Fan Calculation", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "exhaust_fan_calculation", "Exhaust Fan Calculation")


# =====================================================
# 16 AHU SELECTION
# =====================================================
elif module.startswith("16."):
    st.header("AHU Selection - 50 Rows")

    default = pd.DataFrame([{
        "Ref": ref,
        "AHU Tag": "",
        "Location": "",
        "Cooling Load kW": 0.0,
        "Airflow CMH": 0.0,
        "External Static Pressure Pa": 0.0,
        "CHW Flow L/s": 0.0,
        "Quantity": 1.0
    } for ref in refs])

    edited = editor_with_state("ahu_selection_df_v1", default)

    rows = []
    for _, row in edited.iterrows():
        kw = safe_float(row.get("Cooling Load kW"))
        cmh = safe_float(row.get("Airflow CMH"))
        esp = safe_float(row.get("External Static Pressure Pa"))
        flow = safe_float(row.get("CHW Flow L/s"))
        qty = safe_float(row.get("Quantity"), 1)

        rows.append({
            **row.to_dict(),
            "Total Cooling Load kW": round(kw * qty, 2),
            "Total Airflow CMH": round(cmh * qty, 2),
            "Selection Remarks": "OK" if kw > 0 and cmh > 0 else ""
        })

    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Cooling Load kW"] > 0]

    summary = {
        "Total AHU Cooling Load kW": round(active["Total Cooling Load kW"].sum() if not active.empty else 0, 2),
        "Total AHU Airflow CMH": round(active["Total Airflow CMH"].sum() if not active.empty else 0, 2)
    }

    save_summary("AHU Selection", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "ahu_selection", "AHU Selection")


# =====================================================
# 17 FCU SELECTION
# =====================================================
elif module.startswith("17."):
    st.header("FCU Selection - 50 Rows")

    default = pd.DataFrame([{
        "Ref": ref,
        "FCU Tag": "",
        "Room / Area": "",
        "Cooling Load kW": 0.0,
        "Airflow CMH": 0.0,
        "CHW Flow L/s": 0.0,
        "Quantity": 1.0
    } for ref in refs])

    edited = editor_with_state("fcu_selection_df_v1", default)

    rows = []
    for _, row in edited.iterrows():
        kw = safe_float(row.get("Cooling Load kW"))
        cmh = safe_float(row.get("Airflow CMH"))
        flow = safe_float(row.get("CHW Flow L/s"))
        qty = safe_float(row.get("Quantity"), 1)

        rows.append({
            **row.to_dict(),
            "Total Cooling Load kW": round(kw * qty, 2),
            "Total Airflow CMH": round(cmh * qty, 2),
            "Total CHW Flow L/s": round(flow * qty, 2),
            "Selection Remarks": "OK" if kw > 0 and cmh > 0 else ""
        })

    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Cooling Load kW"] > 0]

    summary = {
        "Total FCU Cooling Load kW": round(active["Total Cooling Load kW"].sum() if not active.empty else 0, 2),
        "Total FCU Airflow CMH": round(active["Total Airflow CMH"].sum() if not active.empty else 0, 2)
    }

    save_summary("FCU Selection", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "fcu_selection", "FCU Selection")


# =====================================================
# 18 EXPORT EXCEL / PDF / WORD
# =====================================================
elif module.startswith("18."):
    st.header("Export Excel / PDF / Word")
    st.info("Individual module export buttons are available inside each calculation module. Full combined export can be added in the next upgrade.")


# =====================================================
# 19 CONDENSATE DRAIN / FINAL SUMMARY
# =====================================================
elif module.startswith("19."):
    st.header("Condensate Drain / Final Summary")
    st.write("Preliminary condensate estimation. Actual drain sizing shall follow project specific and manufacturer details.")

    default = pd.DataFrame([{
        "Ref": ref,
        "Equipment Tag": "FCU-1" if i == 0 else "",
        "Type": "FCU" if i == 0 else "Blank",
        "Cooling Capacity kW": 10.0 if i == 0 else 0.0,
        "Latent Fraction %": 30.0,
        "Drain Pipe Dia mm": 20.0,
        "Pipe Gradient %": 1.0,
        "Quantity": 1.0
    } for i, ref in enumerate(refs)])

    edited = editor_with_state(
        "cond_df_v2",
        default,
        column_config={
            "Type": st.column_config.SelectboxColumn(
                "Type",
                options=["Blank", "FCU", "AHU", "PAHU", "DX Unit", "Other"]
            )
        }
    )

    rows = []
    for _, row in edited.iterrows():
        typ = row.get("Type", "Blank")
        kw = safe_float(row.get("Cooling Capacity kW"))
        latent = safe_float(row.get("Latent Fraction %")) / 100
        dia = safe_float(row.get("Drain Pipe Dia mm"))
        grad = safe_float(row.get("Pipe Gradient %"))
        qty = safe_float(row.get("Quantity"), 1)

        cond_lhr = (kw * latent / 2450) * 3600 if typ != "Blank" else 0.0

        remarks = []
        if typ != "Blank":
            if dia < 20:
                remarks.append("Small drain pipe - verify")
            if grad < 1:
                remarks.append("Low gradient - verify")

        rows.append({
            **row.to_dict(),
            "Condensate L/hr Each": round(cond_lhr, 2),
            "Total Condensate L/hr": round(cond_lhr * qty, 2),
            "Status / Remarks": "; ".join(remarks) if remarks else ("OK" if typ != "Blank" else "")
        })

    result_df = pd.DataFrame(rows)
    active = result_df[result_df["Type"] != "Blank"]

    summary = {
        "Total Condensate L/hr": round(active["Total Condensate L/hr"].sum() if not active.empty else 0, 2),
        "No. of Active Rows": len(active)
    }

    save_summary("Condensate Drain / Final Summary", summary)
    metric_row(summary)
    st.dataframe(result_df, use_container_width=True, height=500)
    show_downloads(result_df, summary, "condensate_drain_final_summary", "Condensate Drain / Final Summary")
