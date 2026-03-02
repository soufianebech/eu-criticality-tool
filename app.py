import re
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

st.set_page_config(page_title="EU Criticality Exercise (Dynamic Materials)", layout="wide")
st.title("EU Criticality Assessment Exercise")

# -----------------------------
# Session state init
# -----------------------------
if "materials" not in st.session_state:
    st.session_state.materials = {}
if "material_order" not in st.session_state:
    st.session_state.material_order = []
if "selected_id" not in st.session_state:
    st.session_state.selected_id = None

# -----------------------------
# Required columns (reference)
# -----------------------------
HHI_REQUIRED_COLS = {
    "Material",
    "Stage (Extraction + Processing)",
    "Scope considered",
    "Country",
    "Supply share",
    "Trade (t)",
    "WGI Scaled",
}

# -----------------------------
# Helpers
# -----------------------------
def round1(x):
    return None if x is None else round(float(x), 1)

def _norm(s: str) -> str:
    s = str(s).strip().lower()
    s = s.replace("\u00a0", " ")   # non-breaking space
    s = s.replace("€", "eur")
    s = re.sub(r"\s+", " ", s)
    return s

def read_excel_with_autoheader(
    xls: pd.ExcelFile,
    sheet_name: str,
    required_tokens: list[str],
    max_scan_rows: int = 40
) -> pd.DataFrame:
    preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=max_scan_rows)
    header_row = None
    for i in range(len(preview)):
        row_vals = [_norm(v) for v in preview.iloc[i].tolist()]
        if all(any(tok in v for v in row_vals) for tok in required_tokens):
            header_row = i
            break
    if header_row is None:
        df_try = pd.read_excel(xls, sheet_name=sheet_name)
        raise ValueError(
            f"Could not detect header row in sheet '{sheet_name}'. Columns read: {list(df_try.columns)}"
        )
    return pd.read_excel(xls, sheet_name=sheet_name, header=header_row)

def read_sheet_safe(xls: pd.ExcelFile, sheet_name: str, required_tokens: list[str]) -> pd.DataFrame:
    """
    1) Try normal read (header=0). If columns contain tokens -> OK
    2) Else fallback to autoheader
    """
    df = pd.read_excel(xls, sheet_name=sheet_name)
    cols = [_norm(c) for c in df.columns]
    if all(any(tok in c for c in cols) for tok in required_tokens):
        return df
    return read_excel_with_autoheader(xls, sheet_name, required_tokens=required_tokens)

def find_col_by_token(df: pd.DataFrame, token: str):
    for c in df.columns:
        if token in _norm(c):
            return c
    return None

def default_sheet_name(xls: pd.ExcelFile, preferred: str) -> str:
    return preferred if preferred in xls.sheet_names else xls.sheet_names[0]

def _to_float(x):
    try:
        return float(x)
    except Exception:
        return None

# ✅ SAFE widget-state setter:
# Only set session_state if the key doesn't exist yet (prevents StreamlitAPIException)
def set_widget_value_if_absent(key: str, value):
    if key not in st.session_state:
        st.session_state[key] = value

# -----------------------------
# Optional inputs from Excel (Others_inputs_EI / Others_inputs_SR)
# -----------------------------
def read_optional_inputs_from_excel(xls: pd.ExcelFile, material_name: str):
    out = {"si_ei": None, "sr_by_stage": {}}
    mkey = material_name.strip().lower()

    # ---- Others_inputs_EI ----
    if "Others_inputs_EI" in xls.sheet_names:
        df = read_sheet_safe(xls, "Others_inputs_EI", required_tokens=["material", "si"])
        cols = {_norm(c): c for c in df.columns}
        col_mat = cols.get("material")
        col_si = cols.get("si_ei") or cols.get("si ei") or cols.get("si_ei ")

        if col_mat and col_si:
            df2 = df.copy()
            df2["_mat"] = df2[col_mat].astype(str).str.strip().str.lower()
            r = df2[df2["_mat"] == mkey]
            if not r.empty:
                out["si_ei"] = _to_float(r.iloc[0][col_si])

    # ---- Others_inputs_SR ----
    if "Others_inputs_SR" in xls.sheet_names:
        df = read_sheet_safe(xls, "Others_inputs_SR", required_tokens=["stage", "material", "import", "eol", "si"])
        cols = {_norm(c): c for c in df.columns}

        col_stage = cols.get("stage")
        col_mat = cols.get("material")
        col_ir = cols.get("import reliance (%)") or cols.get("import reliance") or cols.get("ir")
        col_eol = cols.get("eol (rir)") or cols.get("eol rir") or cols.get("eol_rir")
        col_si = cols.get("si (sr)") or cols.get("si sr") or cols.get("si_sr")

        if col_stage and col_mat and col_ir and col_eol and col_si:
            df2 = df.copy()
            df2["_mat"] = df2[col_mat].astype(str).str.strip().str.lower()
            df2["_stage"] = df2[col_stage].astype(str).str.strip().str.lower()

            dfm = df2[df2["_mat"] == mkey]

            for stg in ["extraction", "processing"]:
                r = dfm[dfm["_stage"] == stg]
                if not r.empty:
                    ir = _to_float(r.iloc[0][col_ir])
                    eol = _to_float(r.iloc[0][col_eol])
                    si = _to_float(r.iloc[0][col_si])

                    # Convert percent to fraction if needed
                    if ir is not None and ir > 1:
                        ir = ir / 100.0

                    out["sr_by_stage"]["Extraction" if stg == "extraction" else "Processing"] = {
                        "ir": ir,
                        "eol_rir": eol,
                        "si_sr": si
                    }

    return out

def apply_optional_inputs_to_material(mat: dict, mid: str, opt: dict, apply_si_ei: bool = True, apply_sr: bool = True):
    """
    Apply optional inputs to:
      - mat dict
      - st.session_state widget keys (only if key not created yet)

    apply_si_ei=False is important in SR stage uploads to avoid StreamlitAPIException.
    """

    # --- SI_EI ---
    if apply_si_ei and opt.get("si_ei") is not None:
        mat["si_ei"] = float(opt["si_ei"])
        set_widget_value_if_absent(f"{mid}_si_ei", float(opt["si_ei"]))

    # --- SR params by stage ---
    if apply_sr:
        for stage_name, vals in opt.get("sr_by_stage", {}).items():
            if stage_name in mat["sr_inputs"]:
                if vals.get("ir") is not None:
                    mat["sr_inputs"][stage_name]["ir"] = float(vals["ir"])
                    set_widget_value_if_absent(f"{mid}_{stage_name}_ir", float(vals["ir"]))
                if vals.get("eol_rir") is not None:
                    mat["sr_inputs"][stage_name]["eol_rir"] = float(vals["eol_rir"])
                    set_widget_value_if_absent(f"{mid}_{stage_name}_eol", float(vals["eol_rir"]))
                if vals.get("si_sr") is not None:
                    mat["sr_inputs"][stage_name]["si_sr"] = float(vals["si_sr"])
                    set_widget_value_if_absent(f"{mid}_{stage_name}_si_sr", float(vals["si_sr"]))

# -----------------------------
# EI computation from Excel
# -----------------------------
def calc_scaled_sum_aq_from_excel(df: pd.DataFrame):
    """
    Σ(A×Q) = Σ(Share × VA)
    Scaling: (Σ(A×Q) × 10) / max(VA)
    """
    msgs = []
    d = df.copy()

    col_share = find_col_by_token(d, "share")
    col_va = find_col_by_token(d, "va")

    if col_share is None or col_va is None:
        msgs.append("EI sheet columns detected: " + ", ".join([str(c) for c in d.columns]))
        raise ValueError("EI sheet must contain columns for Share and VA.")

    share = pd.to_numeric(d[col_share], errors="coerce")
    va = pd.to_numeric(d[col_va], errors="coerce")

    d["A_times_Q"] = share * va

    if d["A_times_Q"].isna().all():
        msgs.append("EI sheet: could not compute (Share × VA). Check numeric values.")
        return None, None, None, msgs

    sum_aq = float(d["A_times_Q"].sum())
    max_va = float(va.max())

    if max_va <= 0:
        msgs.append("EI sheet: max(VA) <= 0, scaling not possible.")
        return sum_aq, max_va, None, msgs

    scaled = (sum_aq * 10.0) / max_va
    return sum_aq, max_va, scaled, msgs

def compute_ei_from_scaled_sum_aq(sum_aq_scaled: float, si_ei: float) -> float:
    return round1(float(sum_aq_scaled) * float(si_ei))

# -----------------------------
# HHI computation from Excel
# -----------------------------
def compute_hhi_wgi_t(df: pd.DataFrame) -> float:
    s = pd.to_numeric(df["Supply share"], errors="coerce")
    wgi = pd.to_numeric(df["WGI Scaled"], errors="coerce")
    t = pd.to_numeric(df["Trade (t)"], errors="coerce")
    return float(((s ** 2) * wgi * t).sum())

def calc_hhi_from_excel(df: pd.DataFrame, stage_name: str):
    msgs = []
    missing = HHI_REQUIRED_COLS - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns in HHI sheet: {sorted(list(missing))}")

    df = df.copy()
    df["Scope considered"] = df["Scope considered"].astype(str).str.strip()
    df["Stage (Extraction + Processing)"] = df["Stage (Extraction + Processing)"].astype(str).str.strip()

    d_stage = df[df["Stage (Extraction + Processing)"].str.lower() == stage_name.lower()].copy()
    if d_stage.empty:
        msgs.append(f"No rows found for stage '{stage_name}'.")
        return None, None, msgs

    d_gs = d_stage[d_stage["Scope considered"].str.lower() == "global"].copy()
    d_eu = d_stage[d_stage["Scope considered"].str.lower() == "eu"].copy()

    hhi_gs = None
    hhi_eu = None

    if not d_gs.empty:
        ss_sum = pd.to_numeric(d_gs["Supply share"], errors="coerce").sum()
        if abs(ss_sum - 1.0) > 0.05:
            msgs.append(f"[Global] Supply share sum = {ss_sum:.3f} (not close to 1).")
        hhi_gs = compute_hhi_wgi_t(d_gs)

    if not d_eu.empty:
        ss_sum = pd.to_numeric(d_eu["Supply share"], errors="coerce").sum()
        # If you want this warning back, uncomment:
        # if abs(ss_sum - 1.0) > 0.05:
        #     msgs.append(f"[EU] Supply share sum = {ss_sum:.3f} (not close to 1).")
        hhi_eu = compute_hhi_wgi_t(d_eu)

    if hhi_gs is None:
        msgs.append("Global scope not available → (HHI_WGI,t)_GS missing.")
    if hhi_eu is None:
        msgs.append("EU scope not available → (HHI_WGI,t)_EU missing.")

    return hhi_gs, hhi_eu, msgs

# -----------------------------
# SR computation (correct formula IR/2)
# -----------------------------
def compute_sr_from_hhi(hhi_gs, hhi_eu, ir: float, eol_rir: float, si_sr: float):
    ir = float(ir)
    w = ir / 2.0

    if (hhi_gs is not None) and (hhi_eu is not None):
        core = (float(hhi_gs) * w) + (float(hhi_eu) * (1.0 - w))
        note = "SR uses both GS and EU sourcing."
    elif hhi_gs is not None:
        core = float(hhi_gs)
        note = "SR uses GS only (EU sourcing missing)."
    elif hhi_eu is not None:
        core = float(hhi_eu)
        note = "SR uses EU only (GS missing)."
    else:
        return None, "SR cannot be computed (both GS and EU missing)."

    sr = core * (1.0 - float(eol_rir)) * float(si_sr)
    return round1(sr), note

def overall_sr(sr_extraction, sr_processing, mode="max"):
    vals = [v for v in [sr_extraction, sr_processing] if v is not None]
    if not vals:
        return None
    if mode == "average":
        return round1(sum(vals) / len(vals))
    return round1(max(vals))

# -----------------------------
# Sidebar (materials list) + thresholds (student input)
# -----------------------------
with st.sidebar:
    st.markdown("### Materials list")

    with st.form("add_material_form", clear_on_submit=True):
        new_name = st.text_input("Add a material (name)", placeholder="e.g., Tungsten, Lithium...")
        add_btn = st.form_submit_button("Add material")
        if add_btn and new_name.strip():
            mat_id = f"mat_{len(st.session_state.material_order) + 1:03d}"
            st.session_state.material_order.append(mat_id)
            st.session_state.materials[mat_id] = {
                "name": new_name.strip(),
                "sum_aq_scaled": 0.0,
                "si_ei": 1.0,
                "stages": ["Extraction"],
                "sr_inputs": {
                    "Extraction": {"hhi_gs": 0.0, "hhi_eu": 0.0, "ir": 0.0, "eol_rir": 0.0, "si_sr": 1.0},
                    "Processing": {"hhi_gs": 0.0, "hhi_eu": 0.0, "ir": 0.0, "eol_rir": 0.0, "si_sr": 1.0},
                },
            }
            st.session_state.selected_id = mat_id

    if st.session_state.material_order:
        labels = {mid: f"{st.session_state.materials[mid]['name']} ({mid})" for mid in st.session_state.material_order}
        st.session_state.selected_id = st.selectbox(
            "Select a material to edit",
            options=st.session_state.material_order,
            format_func=lambda mid: labels[mid],
            index=st.session_state.material_order.index(st.session_state.selected_id)
            if st.session_state.selected_id in st.session_state.material_order
            else 0,
        )

        if st.button("Remove selected material", type="secondary") and st.session_state.selected_id:
            mid_rm = st.session_state.selected_id
            st.session_state.material_order = [x for x in st.session_state.material_order if x != mid_rm]
            st.session_state.materials.pop(mid_rm, None)
            st.session_state.selected_id = st.session_state.material_order[0] if st.session_state.material_order else None

    st.divider()
    sr_mode = st.radio(
        "Overall SR aggregation (if both stages):",
        options=["max"]
    )

    st.divider()
    st.markdown("### EU thresholds (student input)")
    EI_THRESHOLD = st.number_input("EI threshold", min_value=0.0, value=2.8, step=0.1)
    SR_THRESHOLD = st.number_input("SR threshold", min_value=0.0, value=1.0, step=0.1)

# -----------------------------
# Main edit panel
# -----------------------------
mid = st.session_state.selected_id
if not mid:
    st.info("Add at least one material from the sidebar to start.")
    st.stop()

mat = st.session_state.materials[mid]
st.subheader(f"Edit inputs — {mat['name']}")

colA, colB = st.columns([1, 1], gap="large")

# -----------------------------
# EI block (Excel + scaling) + optional sheets auto-fill
# -----------------------------
with colA:
    st.markdown("## Economic Importance (EI)")
    st.caption("Upload an Excel file (recommended sheets: EI_inputs, Others_inputs_EI, Others_inputs_SR).")

    ei_file = st.file_uploader("Upload Excel for EI (+ optional inputs)", type=["xlsx"], key=f"{mid}_ei_file")

    if ei_file is not None:
        xls = pd.ExcelFile(ei_file)

        # ✅ Apply optional inputs here (SI_EI + SR params), BEFORE widgets exist
        opt = read_optional_inputs_from_excel(xls, mat["name"])
        apply_optional_inputs_to_material(mat, mid, opt, apply_si_ei=True, apply_sr=True)

        if opt.get("si_ei") is not None:
            st.success("Auto-filled SI_EI from Others_inputs_EI.")
        if opt.get("sr_by_stage"):
            st.success("Auto-filled IR / EoL-RIR / SI_SR from Others_inputs_SR (when available).")

        ei_sheet = st.selectbox(
            "Select the EI sheet",
            options=xls.sheet_names,
            index=xls.sheet_names.index(default_sheet_name(xls, "EI_inputs")),
            key=f"{mid}_ei_sheet_select"
        )

        try:
            df_ei = read_sheet_safe(xls, ei_sheet, required_tokens=["share", "va"])
            sum_aq, max_va, scaled_sum_aq, msgs = calc_scaled_sum_aq_from_excel(df_ei)

            if scaled_sum_aq is not None:
                mat["sum_aq_scaled"] = float(scaled_sum_aq)
                set_widget_value_if_absent(f"{mid}_sum_aq_scaled", float(scaled_sum_aq))

                st.success(f"EI inputs computed from Excel (sheet: {ei_sheet}).")
                st.write(
                    f"Σ(A×Q) = {round1(sum_aq)} ; max(VA) = {round1(max_va)} ; Scaled Σ(A×Q) = {round1(scaled_sum_aq)}"
                )

            for mmsg in msgs:
                st.warning(mmsg)

        except Exception as e:
            st.error(f"EI Excel error: {e}")

    mat["sum_aq_scaled"] = st.number_input(
        "Scaled Σ(A×Q) (computed or manual)",
        min_value=0.0,
        value=float(mat["sum_aq_scaled"]),
        step=0.01,
        key=f"{mid}_sum_aq_scaled"
    )

    mat["si_ei"] = st.number_input(
        "SI_EI (0–1)",
        min_value=0.0,
        max_value=1.0,
        value=float(mat["si_ei"]),
        step=0.01,
        key=f"{mid}_si_ei"
    )

    ei_value = compute_ei_from_scaled_sum_aq(mat["sum_aq_scaled"], mat["si_ei"])
    st.info(f"EI = {ei_value} (rounded to 1 decimal)")

    st.markdown("## Life-cycle stage(s) available for SR")
    mat["stages"] = st.multiselect(
        "Select stages with available SR data",
        options=["Extraction", "Processing"],
        default=mat.get("stages", ["Extraction"]),
        key=f"{mid}_stages"
    )

# -----------------------------
# SR block (HHI per stage + show SR)
# -----------------------------
with colB:
    st.markdown("## Supply Risk (SR)")
    st.caption("Upload Excel (recommended sheets: SR_inputs + Others_inputs_SR).")

    def stage_block(stage_name: str):
        st.markdown(f"### {stage_name} stage")
        d = mat["sr_inputs"][stage_name]

        hhi_file = st.file_uploader(
            f"Upload Excel for HHI ({stage_name})",
            type=["xlsx"],
            key=f"{mid}_{stage_name}_hhi_file"
        )

        if hhi_file is not None:
            xls = pd.ExcelFile(hhi_file)

            # ✅ Only apply SR params here (never touch SI_EI in stage upload)
            opt = read_optional_inputs_from_excel(xls, mat["name"])
            apply_optional_inputs_to_material(mat, mid, opt, apply_si_ei=False, apply_sr=True)

            hhi_sheet = st.selectbox(
                f"Select the HHI sheet ({stage_name})",
                options=xls.sheet_names,
                index=xls.sheet_names.index(default_sheet_name(xls, "SR_inputs")),
                key=f"{mid}_{stage_name}_hhi_sheet_select"
            )

            try:
                df_hhi = read_sheet_safe(xls, hhi_sheet, required_tokens=["supply share", "wgi"])
                hhi_gs, hhi_eu, msgs = calc_hhi_from_excel(df_hhi, stage_name)

                if hhi_gs is not None:
                    d["hhi_gs"] = float(hhi_gs)
                    set_widget_value_if_absent(f"{mid}_{stage_name}_hhi_gs", float(hhi_gs))
                if hhi_eu is not None:
                    d["hhi_eu"] = float(hhi_eu)
                    set_widget_value_if_absent(f"{mid}_{stage_name}_hhi_eu", float(hhi_eu))

                st.success(f"HHI computed from Excel (sheet: {hhi_sheet}).")
                st.write(f"(HHI_WGI,t)_GS = {round1(hhi_gs)} ; (HHI_WGI,t)_EU = {round1(hhi_eu)}")
                for mmsg in msgs:
                    st.warning(mmsg)

            except Exception as e:
                st.error(f"HHI Excel error: {e}")

        d["hhi_gs"] = st.number_input(
            f"(HHI_WGI,t)_GS ({stage_name})",
            min_value=0.0,
            value=float(d["hhi_gs"]),
            step=0.01,
            key=f"{mid}_{stage_name}_hhi_gs"
        )
        d["hhi_eu"] = st.number_input(
            f"(HHI_WGI,t)_EU ({stage_name})",
            min_value=0.0,
            value=float(d["hhi_eu"]),
            step=0.01,
            key=f"{mid}_{stage_name}_hhi_eu"
        )

        d["ir"] = st.number_input(
            f"Import Reliance IR (0–1) ({stage_name})",
            min_value=0.0,
            max_value=1.0,
            value=float(d["ir"]),
            step=0.01,
            key=f"{mid}_{stage_name}_ir"
        )
        d["eol_rir"] = st.number_input(
            f"EOL-RIR (0–1) ({stage_name})",
            min_value=0.0,
            max_value=1.0,
            value=float(d["eol_rir"]),
            step=0.01,
            key=f"{mid}_{stage_name}_eol"
        )
        d["si_sr"] = st.number_input(
            f"SI_SR (0–1) ({stage_name})",
            min_value=0.0,
            max_value=1.0,
            value=float(d["si_sr"]),
            step=0.01,
            key=f"{mid}_{stage_name}_si_sr"
        )

        hhi_gs_used = d["hhi_gs"] if d["hhi_gs"] > 0 else None
        hhi_eu_used = d["hhi_eu"] if d["hhi_eu"] > 0 else None

        sr_val, sr_note = compute_sr_from_hhi(hhi_gs_used, hhi_eu_used, d["ir"], d["eol_rir"], d["si_sr"])
        if sr_val is None:
            st.info(f"SR ({stage_name}) not computed: {sr_note}")
        else:
            st.info(f"SR ({stage_name}) = {sr_val} — {sr_note}")

        st.divider()
        return sr_val

    sr_ex = None
    sr_pr = None

    if "Extraction" in mat["stages"]:
        sr_ex = stage_block("Extraction")
    if "Processing" in mat["stages"]:
        sr_pr = stage_block("Processing")

# -----------------------------
# Results table (all materials)
# -----------------------------
rows = []
for m_id in st.session_state.material_order:
    m = st.session_state.materials[m_id]

    ei_val = compute_ei_from_scaled_sum_aq(m["sum_aq_scaled"], m["si_ei"])

    sr_ex_val = None
    sr_pr_val = None

    if "Extraction" in m.get("stages", []):
        d = m["sr_inputs"]["Extraction"]
        sr_ex_val, _ = compute_sr_from_hhi(
            d["hhi_gs"] if d["hhi_gs"] > 0 else None,
            d["hhi_eu"] if d["hhi_eu"] > 0 else None,
            d["ir"], d["eol_rir"], d["si_sr"]
        )

    if "Processing" in m.get("stages", []):
        d = m["sr_inputs"]["Processing"]
        sr_pr_val, _ = compute_sr_from_hhi(
            d["hhi_gs"] if d["hhi_gs"] > 0 else None,
            d["hhi_eu"] if d["hhi_eu"] > 0 else None,
            d["ir"], d["eol_rir"], d["si_sr"]
        )

    sr_overall = overall_sr(sr_ex_val, sr_pr_val, mode=sr_mode)
    is_critical = (sr_overall is not None) and (ei_val >= EI_THRESHOLD) and (sr_overall >= SR_THRESHOLD)

    rows.append({
        "Material": m["name"],
        "SR (Extraction)": sr_ex_val,
        "SR (Processing)": sr_pr_val,
        "SR (Overall)": sr_overall,
        "EI": ei_val,
        "Critical?": "YES" if is_critical else "NO",
    })

df = pd.DataFrame(rows)

st.divider()
st.subheader("Results (all materials)")
st.dataframe(df, use_container_width=True)

csv = df.to_csv(index=False).encode("utf-8")
st.download_button("Download results as CSV", data=csv, file_name="criticality_results.csv", mime="text/csv")

# -----------------------------
# Plot EI vs SR overall
# -----------------------------
st.subheader("EI vs SR matrix (overall SR shown)")
fig, ax = plt.subplots()
ax.axvline(EI_THRESHOLD, linestyle="--")
ax.axhline(SR_THRESHOLD, linestyle="--")

plot_df = df.dropna(subset=["SR (Overall)"]).copy()
ax.scatter(plot_df["EI"], plot_df["SR (Overall)"])

for _, r in plot_df.iterrows():
    ax.annotate(r["Material"], (r["EI"], r["SR (Overall)"]),
                textcoords="offset points", xytext=(6, 6), ha="left")

ax.set_xlabel("Economic Importance (EI)")
ax.set_ylabel("Supply Risk (SR) — overall")
ax.set_title(f"EU Criticality Matrix (overall SR = {sr_mode})")
st.pyplot(fig)
