#Power factory connection with python
import os, sys
Dig_path=r"C:\Program Files\DIgSILENT\PowerFactory 2024\Python\3.12"
sys.path.append(Dig_path)
os.environ['PATH']+=';'+Dig_path

import powerfactory as pf
import numpy as np
import pandas as pd

# 1. Initialize PowerFactory
app=pf.GetApplication()
if not app:
        raise Exception("Could not open PowerFactory application.")
print(app)

# 2. Activate your project
project_name = "PHC"  # Replace with your actual project name
app.ActivateProject(project_name)

project = app.GetActiveProject()
ldf = app.GetFromStudyCase("ComLdf")
if project is None:
    raise Exception("Please activate your test sgphc Project.")
print(project)


# -----------------------------
# Get NETWORK ELEMENTS
# -----------------------------
buses = app.GetCalcRelevantObjects("*.ElmTerm")
lines = app.GetCalcRelevantObjects("*.ElmLne")
trafos = app.GetCalcRelevantObjects("*.ElmTr2")
loads = app.GetCalcRelevantObjects("*.ElmLod")
ext_grids = app.GetCalcRelevantObjects("*.ElmXnet")
pv_units = app.GetCalcRelevantObjects("*.ElmGenstat")

# System summary 

def run_buswise_comparison():

    app = pf.GetApplication()
    app.ActivateProject("PHC")

    ldf = app.GetFromStudyCase("ComLdf")

    terminals = app.GetCalcRelevantObjects("*.ElmTerm")
    lines = app.GetCalcRelevantObjects("*.ElmLne")
    pv_units = app.GetCalcRelevantObjects("*.ElmGenstat")

    # -----------------------------
    # STORE ORIGINAL PV SETTINGS
    # -----------------------------
    pv_original = {
        pv: {
            "pgini": pv.GetAttribute("pgini"),
            "qgini": pv.GetAttribute("qgini"),
            "outserv": pv.outserv
        }
        for pv in pv_units
    }

    # =============================
    # BASE CASE (NO PV)
    # =============================
    for pv in pv_units:
        pv.SetAttribute("pgini", 0.0)
        pv.SetAttribute("qgini", 0.0)

    ldf.Execute()

    base_voltage = {}
    bus_load = {}

    for term in terminals:
        bus = term.loc_name

        base_voltage[bus] = term.GetAttribute("m:u")

        p_load = sum(
            l.GetAttribute("plini")
            for l in term.GetConnectedElements()
            if l.GetClassName() == "ElmLod"
        )

        q_load = sum(
            l.GetAttribute("qlini")
            for l in term.GetConnectedElements()
            if l.GetClassName() == "ElmLod"
        )

        bus_load[bus] = (p_load, q_load)

    total_ploss_before = sum(l.GetAttribute("m:Ploss") for l in lines)
    total_qloss_before = sum(l.GetAttribute("m:Qloss") for l in lines)

    total_load_p = sum(v[0] for v in bus_load.values())
    total_load_q = sum(v[1] for v in bus_load.values())

    # =============================
    # PV CASE (50 kW EACH)
    # =============================
    injection_mw = 0.05  # 50 kW per PV

    for pv in pv_units:
        pv.SetAttribute("pgini", injection_mw)
        pv.SetAttribute("qgini", 0.0)

    ldf.Execute()

    post_voltage = {}

    # -----------------------------
    # CORRECT PV MAPPING PER BUS
    # -----------------------------
    bus_pv = {term.loc_name: (0.0, 0.0) for term in terminals}

    for pv in pv_units:

        if pv.outserv == 1:
            continue

        cub = pv.GetAttribute("bus1")
        if not cub:
            continue

        term = cub.GetAttribute("cterm")
        if not term:
            continue

        bus = term.loc_name

        p_kw = pv.GetAttribute("pgini") * 1000
        q_kvar = pv.GetAttribute("qgini") * 1000

        bus_pv[bus] = (
            bus_pv[bus][0] + p_kw,
            bus_pv[bus][1] + q_kvar
        )

    # collect post voltages
    for term in terminals:
        bus = term.loc_name
        post_voltage[bus] = term.GetAttribute("m:u")

    total_ploss_after = sum(l.GetAttribute("m:Ploss") for l in lines)
    total_qloss_after = sum(l.GetAttribute("m:Qloss") for l in lines)

    total_pv_p = sum(v[0] for v in bus_pv.values())
    total_pv_q = sum(v[1] for v in bus_pv.values())

    # =============================
    # EXTERNAL INFEED CALCULATION
    # =============================
    external_infeed_kw = total_load_p + total_ploss_after - total_pv_p

    # =============================
    # BUS-WISE TABLE
    # =============================
    data = []

    for bus in base_voltage.keys():

        v_base = base_voltage[bus]
        v_after = post_voltage[bus]
        dv = v_after - v_base
        dv_pct = (dv / v_base) * 100 if v_base else 0

        data.append({
            "Bus": bus,
            "V_base (pu)": v_base,
            "V_after (pu)": v_after,
            "ΔV (pu)": dv,
            "ΔV (%)": dv_pct,
            "Load P (kW)": bus_load[bus][0],
            "Load Q (kVAr)": bus_load[bus][1],
            "PV P (kW)": bus_pv[bus][0],
            "PV Q (kVAr)": bus_pv[bus][1],
        })

    df_bus = pd.DataFrame(data)

    # Natural sorting of buses
    df_bus["Bus_sort"] = df_bus["Bus"].str.extract(r"(\d+)").astype(float)
    df_bus = df_bus.sort_values(by="Bus_sort").drop(columns=["Bus_sort"])

    # =============================
    # CLEAN SYSTEM SUMMARY
    # =============================
    summary = {
        "Total Load P (kW)": total_load_p,
        "Total Load Q (kVAr)": total_load_q,

        "Total PV P (kW)": total_pv_p,
        "Total PV Q (kVAr)": total_pv_q,

        "Loss P Before (kW)": total_ploss_before,
        "Loss P After (kW)": total_ploss_after,
        "Δ Loss P (kW)": total_ploss_after - total_ploss_before,

        "Loss Q Before (kVAr)": total_qloss_before,
        "Loss Q After (kVAr)": total_qloss_after,
        "Δ Loss Q (kVAr)": total_qloss_after - total_qloss_before,

        "External Infeed (kW)": external_infeed_kw,

        "Power Balance Check (kW)": (
            external_infeed_kw + total_pv_p - total_ploss_after - total_load_p
        ),

        "No. of Buses": len(terminals),
        "No. of Lines": len(lines),
        "No. of PV Units": len(pv_units)
    }

    # =============================
    # PRINT RESULTS
    # =============================
    print("\n================ BUS-WISE RESULTS ================")
    print(df_bus.round(4).to_string(index=False))

    print("\n================ SYSTEM SUMMARY ================")
    print("-" * 65)

    for k, v in summary.items():
        if isinstance(v, float):
            print(f"{k:<40} : {v:>12.4f}")
        else:
            print(f"{k:<40} : {v:>12}")

    print("-" * 65)

    # =============================
    # POWER BALANCE CHECK
    # =============================
    tolerance = 1e-3
    balance_error = abs(
        external_infeed_kw + total_pv_p - total_ploss_after - total_load_p
    )

    if balance_error < tolerance:
        print("\n✔ Power balance satisfied")
    else:
        print(f"\n⚠ Power mismatch detected: {balance_error:.6f} kW")

    # =============================
    # RESTORE ORIGINAL PV SETTINGS
    # =============================
    for pv in pv_units:
        pv.SetAttribute("pgini", pv_original[pv]["pgini"])
        pv.SetAttribute("qgini", pv_original[pv]["qgini"])
        pv.outserv = pv_original[pv]["outserv"]

    ldf.Execute()

    return df_bus, summary


if __name__ == "__main__":
    run_buswise_comparison()


#Restore Clean Base Case

def restore_base_case():

    app = pf.GetApplication()
    app.ActivateProject("sgphc")

    ldf = app.GetFromStudyCase("ComLdf")

    pv_units = app.GetCalcRelevantObjects("*.ElmGenstat")
    loads = app.GetCalcRelevantObjects("*.ElmLod")
    lines = app.GetCalcRelevantObjects("*.ElmLne")

    # ---------------------------
    # RESET PV → ZERO
    # ---------------------------
    for pv in pv_units:
        pv.SetAttribute("pgini", 0.0)
        pv.SetAttribute("qgini", 0.0)
        pv.outserv = 0  # ensure in service

    # ---------------------------
    # RUN LOAD FLOW
    # ---------------------------
    ldf.Execute()

    # ---------------------------
    # CHECK VALUES
    # ---------------------------
    total_load = sum(l.GetAttribute("plini") for l in loads)
    total_loss = sum(l.GetAttribute("m:Ploss") for l in lines)
    total_pv = sum(pv.GetAttribute("pgini") for pv in pv_units) * 1000

    print("\n===== BASE CASE RESTORED =====")
    print(f"Total Load (kW): {round(total_load,2)}")
    print(f"Total Loss (kW): {round(total_loss,2)}")
    print(f"Total PV (kW): {round(total_pv,2)}")

    return total_load, total_loss, total_pv


if __name__ == "__main__":
    restore_base_case()