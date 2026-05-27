import math
from datetime import datetime
from io import BytesIO

import streamlit as st

try:
    from fpdf import FPDF
except Exception:
    FPDF = None

# =====================================================
# APP CONFIG
# =====================================================
st.set_page_config(page_title="ACMV Engineering Calculator", layout="wide")
st.title("ACMV Engineering Calculator")
st.caption(
    "Preliminary ACMV engineering calculator for fast design checks. "
    "Final selection must be verified with consultant requirements, authority/code requirements, "
    "manufacturer catalogues, fan curves, pump curves, NPSH, acoustic criteria, and project specifications."
)

# =====================================================
# COMMON FUNCTIONS
# =====================================================
def safe_div(a, b):
    if b == 0 or b is None:
        return 0
    return a / b


def cmh_to_lps(cmh):
    return cmh / 3.6


def cmh_to_m3s(cmh):
    return cmh / 3600


def lps_to_m3s(lps):
    return lps / 1000


def rect_area(width_mm, height_mm):
    return (width_mm / 1000) * (height_mm / 1000)


def round_area(dia_mm):
    dia_m = dia_mm / 1000
    return math.pi * dia_m ** 2 / 4


def velocity_pressure(velocity_ms, air_density=1.2):
    return 0.5 * air_density * velocity_ms ** 2


def nearest_standard_motor(required_kw):
    motors = [
        0.37, 0.55, 0.75, 1.1, 1.5, 2.2, 3.0, 4.0, 5.5, 7.5,
        11, 15, 18.5, 22, 30, 37, 45, 55, 75, 90, 110, 132, 160,
        200, 250, 315, 400
    ]
    for motor in motors:
        if motor >= required_kw:
            return motor
    return None


def motor_fla_kw(kw, voltage, pf, efficiency_percent):
    eff = efficiency_percent / 100
    return safe_div(kw * 1000, math.sqrt(3) * voltage * pf * eff)


def make_pdf(title, lines):
    if FPDF is None:
        return None

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 15)
    pdf.cell(0, 10, title, ln=True)
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 8, f"Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}", ln=True)
    pdf.ln(4)

    for line in lines:
        safe_line = str(line).replace("Δ", "Delta").replace("²", "2").replace("³", "3")
        pdf.multi_cell(190, 7, txt=str(safe_line))

    pdf.ln(4)
    pdf.set_font("Arial", "I", 9)
    pdf.multi_cell(
        0,
        6,
        "Note: This is a preliminary engineering calculation. Final equipment selection shall be verified with approved specifications, manufacturer catalogues, and project requirements.",
    )

    data = bytes(pdf.output(dest="S"))
    return BytesIO(data)


def pdf_button(title, lines, filename):
    pdf_file = make_pdf(title, lines)
    if pdf_file:
        st.download_button(
            label="Save this section as PDF",
            data=pdf_file,
            file_name=filename,
            mime="application/pdf",
            key=f"pdf_{filename}",
        )
    else:
        st.warning("PDF export package not found. In PowerShell run: pip install fpdf2")


# =====================================================
# SIDEBAR NAVIGATION
# =====================================================
st.sidebar.title("ACMV Modules")
module = st.sidebar.radio(
    "Choose calculator",
    [
        "1. Airflow Converter",
        "2. Detailed Heat Load",
        "3. Fresh Air / Ventilation",
        "4. Automatic Duct Sizer",
        "5. Duct Static Pressure",
        "6. Fan Selection + Motor FLA",
        "7. Pipe Sizing",
        "8. Pump Selection",
        "9. Electrical Power",
        "10. Equipment Schedule",
        "11. Cost / Quantity Tracker",
        "12. Report Notes",
    ],
)

# =====================================================
# 1. AIRFLOW CONVERTER
# =====================================================
if module == "1. Airflow Converter":
    st.header("1. Airflow Converter")

    cmh = st.number_input("Airflow (CMH)", min_value=0.0, value=5000.0, step=100.0, key="air_cmh")
    lps = cmh_to_lps(cmh)
    m3s = cmh_to_m3s(cmh)
    cfm = cmh * 0.5886

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CMH", f"{cmh:.2f}")
    c2.metric("L/s", f"{lps:.2f}")
    c3.metric("m³/s", f"{m3s:.4f}")
    c4.metric("CFM", f"{cfm:.2f}")

    lines = [
        f"Airflow: {cmh:.2f} CMH",
        f"L/s = CMH / 3.6 = {lps:.2f} L/s",
        f"m3/s = CMH / 3600 = {m3s:.4f} m3/s",
        f"CFM = CMH x 0.5886 = {cfm:.2f} CFM",
    ]
    pdf_button("Airflow Conversion", lines, "airflow_conversion.pdf")

# =====================================================
# 2. DETAILED HEAT LOAD
# =====================================================
elif module == "2. Detailed Heat Load":
    st.header("2. Detailed Heat Load Calculator")

    area_m2 = st.number_input("Room Area (m²)", min_value=0.0, value=50.0, key="hl_area")
    base_load_wm2 = st.number_input("Base Room Load (W/m²)", min_value=0.0, value=120.0, key="hl_base")
    people = st.number_input("Number of People", min_value=0.0, value=5.0, key="hl_people")
    person_w = st.number_input("Load per Person (W/person)", min_value=0.0, value=120.0, key="hl_person_w")
    lighting_wm2 = st.number_input("Lighting Load (W/m²)", min_value=0.0, value=10.0, key="hl_light")
    equipment_kw = st.number_input("Equipment Load (kW)", min_value=0.0, value=2.0, key="hl_eq")
    fresh_air_lps = st.number_input("Fresh Air (L/s)", min_value=0.0, value=50.0, key="hl_fa")
    outdoor_temp = st.number_input("Outdoor Temp (°C)", value=32.0, key="hl_ot")
    indoor_temp = st.number_input("Indoor Temp (°C)", value=24.0, key="hl_it")
    safety_percent = st.number_input("Safety Allowance (%)", min_value=0.0, value=10.0, key="hl_safety")

    room_kw = area_m2 * base_load_wm2 / 1000
    people_kw = people * person_w / 1000
    lighting_kw = area_m2 * lighting_wm2 / 1000
    fresh_air_kw = 1.2 * 1.006 * (fresh_air_lps / 1000) * (outdoor_temp - indoor_temp)
    total_kw = room_kw + people_kw + lighting_kw + equipment_kw + fresh_air_kw
    final_kw = total_kw * (1 + safety_percent / 100)
    tr = final_kw / 3.517

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Load", f"{total_kw:.2f} kW")
    c2.metric("Final Load", f"{final_kw:.2f} kW")
    c3.metric("Approx. TR", f"{tr:.2f}")

    lines = [
        f"Room load: {room_kw:.2f} kW",
        f"People load: {people_kw:.2f} kW",
        f"Lighting load: {lighting_kw:.2f} kW",
        f"Equipment load: {equipment_kw:.2f} kW",
        f"Fresh air sensible load: {fresh_air_kw:.2f} kW",
        f"Total cooling load: {total_kw:.2f} kW",
        f"Final cooling load with safety: {final_kw:.2f} kW",
        f"Approx TR: {tr:.2f}",
    ]
    pdf_button("Detailed Heat Load Calculation", lines, "detailed_heat_load.pdf")

# =====================================================
# 3. FRESH AIR / VENTILATION
# =====================================================
elif module == "3. Fresh Air / Ventilation":
    st.header("3. Fresh Air / Ventilation Calculator")

    length_m = st.number_input("Room Length (m)", min_value=0.0, value=10.0, key="fa_len")
    width_m = st.number_input("Room Width (m)", min_value=0.0, value=5.0, key="fa_wid")
    height_m = st.number_input("Room Height (m)", min_value=0.0, value=3.0, key="fa_hei")
    ach = st.number_input("Required ACH", min_value=0.0, value=6.0, key="fa_ach")
    persons = st.number_input("Number of Persons", min_value=0.0, value=5.0, key="fa_person")
    lps_per_person = st.number_input("Fresh Air per Person (L/s/person)", min_value=0.0, value=10.0, key="fa_lps_person")

    volume_m3 = length_m * width_m * height_m
    ach_cmh = volume_m3 * ach
    person_lps = persons * lps_per_person
    person_cmh = person_lps * 3.6
    recommended_cmh = max(ach_cmh, person_cmh)
    recommended_lps = recommended_cmh / 3.6

    c1, c2, c3 = st.columns(3)
    c1.metric("Room Volume", f"{volume_m3:.2f} m³")
    c2.metric("ACH Method", f"{ach_cmh:.2f} CMH")
    c3.metric("Recommended FA", f"{recommended_cmh:.2f} CMH")

    lines = [
        f"Room volume: {volume_m3:.2f} m3",
        f"ACH method airflow: {ach_cmh:.2f} CMH",
        f"Person method airflow: {person_cmh:.2f} CMH",
        f"Recommended fresh air: {recommended_cmh:.2f} CMH / {recommended_lps:.2f} L/s",
    ]
    pdf_button("Fresh Air Ventilation Calculation", lines, "fresh_air_ventilation.pdf")

# =====================================================
# 4. AUTOMATIC DUCT SIZER
# =====================================================
elif module == "4. Automatic Duct Sizer":
    st.header("4. Automatic Duct Sizer")

    airflow_cmh = st.number_input("Airflow (CMH)", min_value=0.0, value=5000.0, key="ds_cmh")
    target_velocity = st.number_input("Target Velocity (m/s)", min_value=0.1, value=6.0, key="ds_vel")
    preferred_width = st.number_input("Preferred Duct Width (mm)", min_value=1.0, value=800.0, key="ds_width")

    q_m3s = cmh_to_m3s(airflow_cmh)
    required_area = safe_div(q_m3s, target_velocity)
    required_height = safe_div(required_area, preferred_width / 1000) * 1000
    actual_area = rect_area(preferred_width, required_height)
    actual_velocity = safe_div(q_m3s, actual_area)

    c1, c2, c3 = st.columns(3)
    c1.metric("Required Area", f"{required_area:.3f} m²")
    c2.metric("Suggested Size", f"{preferred_width:.0f} x {required_height:.0f} mm")
    c3.metric("Actual Velocity", f"{actual_velocity:.2f} m/s")

    lines = [
        f"Airflow: {airflow_cmh:.2f} CMH",
        f"Target velocity: {target_velocity:.2f} m/s",
        f"Required area: {required_area:.3f} m2",
        f"Suggested duct size: {preferred_width:.0f} x {required_height:.0f} mm",
        f"Actual velocity: {actual_velocity:.2f} m/s",
    ]
    pdf_button("Automatic Duct Sizing", lines, "automatic_duct_sizing.pdf")

# =====================================================
# 5. DUCT STATIC PRESSURE
# =====================================================
elif module == "5. Duct Static Pressure":
    st.header("5. Duct Static Pressure Calculator")

    airflow_cmh = st.number_input("Airflow (CMH)", min_value=0.0, value=5000.0, key="sp_cmh")
    duct_w = st.number_input("Duct Width (mm)", min_value=1.0, value=800.0, key="sp_w")
    duct_h = st.number_input("Duct Height (mm)", min_value=1.0, value=500.0, key="sp_h")
    length_m = st.number_input("Straight Duct Length (m)", min_value=0.0, value=20.0, key="sp_len")
    pa_per_m = st.number_input("Straight Duct Loss (Pa/m)", min_value=0.0, value=1.0, key="sp_pa_m")
    elbows = st.number_input("No. of Elbows", min_value=0.0, value=4.0, key="sp_elbows")
    k_elbow = st.number_input("K Factor per Elbow", min_value=0.0, value=0.9, key="sp_k")
    filters_pa = st.number_input("Filter / Grille / Damper / Others Loss (Pa)", min_value=0.0, value=80.0, key="sp_filter")
    safety_percent = st.number_input("Safety Allowance (%)", min_value=0.0, value=10.0, key="sp_safety")

    q_m3s = cmh_to_m3s(airflow_cmh)
    area_m2 = rect_area(duct_w, duct_h)
    velocity_ms = safe_div(q_m3s, area_m2)
    vp_pa = velocity_pressure(velocity_ms)
    straight_loss = length_m * pa_per_m
    elbow_loss = elbows * k_elbow * vp_pa
    subtotal = straight_loss + elbow_loss + filters_pa
    final_esp = subtotal * (1 + safety_percent / 100)

    c1, c2, c3 = st.columns(3)
    c1.metric("Velocity", f"{velocity_ms:.2f} m/s")
    c2.metric("Velocity Pressure", f"{vp_pa:.2f} Pa")
    c3.metric("Final ESP", f"{final_esp:.2f} Pa")

    lines = [
        f"Airflow: {airflow_cmh:.2f} CMH",
        f"Duct size: {duct_w:.0f} x {duct_h:.0f} mm",
        f"Area: {area_m2:.3f} m2",
        f"Velocity: {velocity_ms:.2f} m/s",
        f"Velocity pressure: {vp_pa:.2f} Pa",
        f"Straight duct loss: {straight_loss:.2f} Pa",
        f"Elbow loss: {elbow_loss:.2f} Pa",
        f"Other losses: {filters_pa:.2f} Pa",
        f"Final ESP: {final_esp:.2f} Pa",
    ]
    pdf_button("Duct Static Pressure Calculation", lines, "duct_static_pressure.pdf")

# =====================================================
# 6. FAN SELECTION + MOTOR FLA
# =====================================================
elif module == "6. Fan Selection + Motor FLA":
    st.header("6. Fan Selection + Motor FLA")

    fan_cmh = st.number_input("Fan Airflow (CMH)", min_value=0.0, value=5000.0, key="fan_cmh")
    esp_pa = st.number_input("External Static Pressure (Pa)", min_value=0.0, value=500.0, key="fan_esp")
    fan_eff_percent = st.number_input("Fan Total Efficiency (%)", min_value=1.0, value=55.0, key="fan_eff")
    voltage = st.number_input("Voltage (V)", min_value=1.0, value=415.0, key="fan_volt")
    pf = st.number_input("Power Factor", min_value=0.1, value=0.85, key="fan_pf")
    motor_eff_percent = st.number_input("Motor Efficiency (%)", min_value=1.0, value=90.0, key="fan_motor_eff")

    q_m3s = cmh_to_m3s(fan_cmh)
    fan_power_kw = safe_div(q_m3s * esp_pa, fan_eff_percent / 100) / 1000
    motor_required_kw = fan_power_kw * 1.15
    selected_motor = nearest_standard_motor(motor_required_kw)
    fla = motor_fla_kw(selected_motor or motor_required_kw, voltage, pf, motor_eff_percent)

    c1, c2, c3 = st.columns(3)
    c1.metric("Fan Power", f"{fan_power_kw:.2f} kW")
    c2.metric("Selected Motor", f"{selected_motor:.2f} kW" if selected_motor else "Above standard list")
    c3.metric("Estimated FLA", f"{fla:.2f} A")

    lines = [
        f"Fan airflow: {fan_cmh:.2f} CMH",
        f"ESP: {esp_pa:.2f} Pa",
        f"Fan efficiency: {fan_eff_percent:.2f}%",
        f"Fan power: {fan_power_kw:.2f} kW",
        f"Required motor with 15% margin: {motor_required_kw:.2f} kW",
        f"Selected motor: {selected_motor} kW",
        f"Estimated FLA: {fla:.2f} A",
    ]
    pdf_button("Fan Selection and Motor FLA", lines, "fan_selection_motor_fla.pdf")

# =====================================================
# 7. PIPE SIZING
# =====================================================
elif module == "7. Pipe Sizing":
    st.header("7. Pipe Sizing Calculator")

    flow_lps = st.number_input("Flow (L/s)", min_value=0.0, value=2.0, key="pipe_flow")
    dia_mm = st.number_input("Internal Pipe Diameter (mm)", min_value=1.0, value=50.0, key="pipe_dia")
    pipe_length = st.number_input("Pipe Length (m)", min_value=0.0, value=30.0, key="pipe_len")
    friction_rate = st.number_input("Friction Rate (Pa/m)", min_value=0.0, value=250.0, key="pipe_fr")
    elbows = st.number_input("No. of Elbows", min_value=0.0, value=4.0, key="pipe_elbow")
    eq_elbow = st.number_input("Equivalent Length per Elbow (m)", min_value=0.0, value=1.5, key="pipe_eq_elbow")
    valves = st.number_input("No. of Valves", min_value=0.0, value=2.0, key="pipe_valve")
    eq_valve = st.number_input("Equivalent Length per Valve (m)", min_value=0.0, value=3.0, key="pipe_eq_valve")

    area_m2 = round_area(dia_mm)
    velocity_ms = safe_div(lps_to_m3s(flow_lps), area_m2)
    total_equiv_length = pipe_length + elbows * eq_elbow + valves * eq_valve
    friction_pa = friction_rate * total_equiv_length
    friction_m = friction_pa / 9810

    c1, c2, c3 = st.columns(3)
    c1.metric("Pipe Velocity", f"{velocity_ms:.2f} m/s")
    c2.metric("Equivalent Length", f"{total_equiv_length:.2f} m")
    c3.metric("Friction Loss", f"{friction_m:.2f} m")

    if velocity_ms < 0.9:
        st.warning("Velocity is low. Pipe may be oversized.")
    elif velocity_ms > 2.4:
        st.error("Velocity is high. Consider increasing pipe size.")
    else:
        st.success("Velocity is within common CHW design range.")

    lines = [
        f"Flow: {flow_lps:.2f} L/s",
        f"Pipe internal diameter: {dia_mm:.2f} mm",
        f"Velocity: {velocity_ms:.2f} m/s",
        f"Total equivalent length: {total_equiv_length:.2f} m",
        f"Friction loss: {friction_m:.2f} m",
    ]
    pdf_button("Pipe Sizing Calculation", lines, "pipe_sizing.pdf")

# =====================================================
# 8. PUMP SELECTION
# =====================================================
elif module == "8. Pump Selection":
    st.header("8. Pump Selection Calculator")

    flow_lps = st.number_input("Water Flow (L/s)", min_value=0.0, value=2.0, key="pump_flow")
    static_head = st.number_input("Static Head / Height Difference (m)", min_value=0.0, value=10.0, key="pump_static")
    pipe_friction = st.number_input("Pipe Friction Loss (m)", min_value=0.0, value=5.0, key="pump_pipe_friction")
    fittings_loss = st.number_input("Fittings / Valve Loss (m)", min_value=0.0, value=3.0, key="pump_fit")
    equipment_loss = st.number_input("Equipment Loss (m)", min_value=0.0, value=5.0, key="pump_eq")
    safety_percent = st.number_input("Safety Allowance (%)", min_value=0.0, value=10.0, key="pump_safety")
    pump_eff_percent = st.number_input("Pump Efficiency (%)", min_value=1.0, value=60.0, key="pump_eff")

    base_head = static_head + pipe_friction + fittings_loss + equipment_loss
    total_head = base_head * (1 + safety_percent / 100)
    pump_power_kw = safe_div(1000 * 9.81 * lps_to_m3s(flow_lps) * total_head, pump_eff_percent / 100) / 1000
    motor_required_kw = pump_power_kw * 1.15
    selected_motor = nearest_standard_motor(motor_required_kw)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Pump Head", f"{total_head:.2f} m")
    c2.metric("Pump Power", f"{pump_power_kw:.2f} kW")
    c3.metric("Selected Motor", f"{selected_motor:.2f} kW" if selected_motor else "Above standard list")

    lines = [
        f"Flow: {flow_lps:.2f} L/s",
        f"Static head: {static_head:.2f} m",
        f"Pipe friction: {pipe_friction:.2f} m",
        f"Fittings loss: {fittings_loss:.2f} m",
        f"Equipment loss: {equipment_loss:.2f} m",
        f"Total pump head: {total_head:.2f} m",
        f"Pump power: {pump_power_kw:.2f} kW",
        f"Selected motor: {selected_motor} kW",
    ]
    pdf_button("Pump Selection Calculation", lines, "pump_selection.pdf")

# =====================================================
# 9. ELECTRICAL POWER
# =====================================================
elif module == "9. Electrical Power":
    st.header("9. Electrical Power Calculator")

    kw = st.number_input("Motor Power (kW)", min_value=0.0, value=5.5, key="elec_kw")
    voltage = st.number_input("Voltage (V)", min_value=1.0, value=415.0, key="elec_volt")
    pf = st.number_input("Power Factor", min_value=0.1, value=0.85, key="elec_pf")
    eff_percent = st.number_input("Efficiency (%)", min_value=1.0, value=90.0, key="elec_eff")

    fla = motor_fla_kw(kw, voltage, pf, eff_percent)
    mccb_guide = fla * 1.25

    c1, c2 = st.columns(2)
    c1.metric("Estimated FLA", f"{fla:.2f} A")
    c2.metric("MCCB Guide", f"{mccb_guide:.2f} A")

    lines = [
        f"Motor power: {kw:.2f} kW",
        f"Voltage: {voltage:.2f} V",
        f"Power factor: {pf:.2f}",
        f"Efficiency: {eff_percent:.2f}%",
        f"Estimated FLA: {fla:.2f} A",
        f"Minimum MCCB guide: {mccb_guide:.2f} A",
    ]
    pdf_button("Electrical Power Calculation", lines, "electrical_power.pdf")

# =====================================================
# 10. EQUIPMENT SCHEDULE
# =====================================================
elif module == "10. Equipment Schedule":
    st.header("10. Equipment Schedule Generator")

    tag = st.text_input("Equipment Tag", value="AHU-01", key="sch_tag")
    equipment_type = st.selectbox("Equipment Type", ["AHU", "FCU", "Fan", "Pump", "Damper", "Other"], key="sch_type")
    capacity = st.number_input("Capacity (kW / CMH / L/s)", value=37.5, key="sch_capacity")
    airflow = st.number_input("Airflow (CMH)", value=5000.0, key="sch_airflow")
    esp = st.number_input("ESP / Head (Pa or m)", value=500.0, key="sch_esp")
    power = st.number_input("Power (kW)", value=2.2, key="sch_power")
    remarks = st.text_area("Remarks", value="Preliminary selection for review.", key="sch_remarks")

    st.subheader("Generated Equipment Schedule")
    st.write(f"**Equipment Tag:** {tag}")
    st.write(f"**Equipment Type:** {equipment_type}")
    st.write(f"**Capacity:** {capacity}")
    st.write(f"**Airflow:** {airflow:.2f} CMH")
    st.write(f"**ESP / Head:** {esp:.2f}")
    st.write(f"**Power:** {power:.2f} kW")
    st.write(f"**Remarks:** {remarks}")

    lines = [
        f"Equipment tag: {tag}",
        f"Equipment type: {equipment_type}",
        f"Capacity: {capacity}",
        f"Airflow: {airflow:.2f} CMH",
        f"ESP / Head: {esp:.2f}",
        f"Power: {power:.2f} kW",
        f"Remarks: {remarks}",
    ]
    pdf_button("Equipment Schedule", lines, "equipment_schedule.pdf")

# =====================================================
# 11. COST / QUANTITY TRACKER
# =====================================================
elif module == "11. Cost / Quantity Tracker":
    st.header("11. Cost / Quantity Tracker")

    material_qty = st.number_input("Material Quantity", min_value=0.0, value=10.0, key="cost_qty")
    unit_rate = st.number_input("Material Unit Rate ($)", min_value=0.0, value=50.0, key="cost_unit")
    labour_days = st.number_input("Labour Days", min_value=0.0, value=3.0, key="cost_days")
    labour_rate = st.number_input("Labour Rate per Day ($)", min_value=0.0, value=120.0, key="cost_labour")
    other_cost = st.number_input("Other Cost ($)", min_value=0.0, value=100.0, key="cost_other")
    margin_percent = st.number_input("Markup / Margin (%)", min_value=0.0, value=15.0, key="cost_margin")

    material_cost = material_qty * unit_rate
    labour_cost = labour_days * labour_rate
    subtotal = material_cost + labour_cost + other_cost
    final_cost = subtotal * (1 + margin_percent / 100)

    c1, c2, c3 = st.columns(3)
    c1.metric("Material Cost", f"${material_cost:.2f}")
    c2.metric("Subtotal", f"${subtotal:.2f}")
    c3.metric("Final Cost", f"${final_cost:.2f}")

    lines = [
        f"Material quantity: {material_qty:.2f}",
        f"Unit rate: ${unit_rate:.2f}",
        f"Material cost: ${material_cost:.2f}",
        f"Labour cost: ${labour_cost:.2f}",
        f"Other cost: ${other_cost:.2f}",
        f"Subtotal: ${subtotal:.2f}",
        f"Final cost with margin: ${final_cost:.2f}",
    ]
    pdf_button("Cost Quantity Tracker", lines, "cost_quantity_tracker.pdf")

# =====================================================
# 12. REPORT NOTES
# =====================================================
elif module == "12. Report Notes":
    st.header("12. Report / Submission Notes")

    project = st.text_input("Project Name", value="ACMV Calculation Report", key="rep_project")
    prepared_by = st.text_input("Prepared By", value="Muthu", key="rep_by")
    notes = st.text_area(
        "Engineering Notes",
        value="Calculations are preliminary and subject to final manufacturer catalogue verification.",
        key="rep_notes",
    )

    st.subheader("Report Notes Preview")
    st.write(f"**Project:** {project}")
    st.write(f"**Prepared By:** {prepared_by}")
    st.write(f"**Notes:** {notes}")

    lines = [
        f"Project: {project}",
        f"Prepared by: {prepared_by}",
        f"Notes: {notes}",
        "Final verification required: consultant requirements, authority/code requirements, manufacturer catalogues, pump curves, fan curves, NPSH, acoustic criteria, electrical protection, and actual site conditions.",
    ]
    pdf_button("ACMV Report Notes", lines, "acmv_report_notes.pdf")
