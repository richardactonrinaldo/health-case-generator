"""
Health Analytics Case Generator (HINF1009)

Generates a client request, a synthetic dataset with built-in relationships,
and an instructor answer key with beginner-friendly Excel steps.

Run locally:  streamlit run app.py
"""

import base64
import hashlib
import io
import json
import re
import zlib

import numpy as np
import pandas as pd
import streamlit as st
import anthropic
from cryptography.fernet import Fernet, InvalidToken
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MODEL = "claude-sonnet-5-5"

SCENARIOS = {
    "Emergency department wait times": "An Ontario community hospital ED wants to understand what drives patient wait times and length of stay.",
    "30-day hospital readmissions": "A hospital wants to know which patients are most likely to be readmitted within 30 days of discharge.",
    "Diabetes clinic outcomes": "A primary care diabetes clinic wants to see which patient factors relate to HbA1c control.",
    "Nurse staffing and patient satisfaction": "A hospital wants to know how unit staffing levels relate to patient satisfaction scores and falls.",
    "Flu vaccination uptake": "A public health unit wants to understand flu vaccine uptake across age groups and regions.",
    "Clinic no-show appointments": "An outpatient clinic wants to reduce missed appointments and understand who misses them.",
}

DIFFICULTY = {
    "Beginner: descriptive stats & PivotTables": (
        "Students are beginner Excel users. Use mostly categorical sources for relationships, "
        "with strong, easy-to-see effects. Questions should be answerable with AVERAGE, MEDIAN, "
        "COUNT, COUNTIF, AVERAGEIF, PivotTables and simple column or bar charts."
    ),
    "Intermediate: comparisons & correlation": (
        "Students know basic Excel. Include at least one numeric-to-numeric relationship. "
        "Questions may use COUNTIFS, AVERAGEIFS, CORREL, PivotTables with two fields, "
        "and scatter plots with a trendline."
    ),
}

ARIAL = Font(name="Arial", size=11)
ARIAL_BOLD = Font(name="Arial", size=11, bold=True)
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")

# --------------------------------------------------------------------------
# Claude calls
# --------------------------------------------------------------------------

SPEC_SYSTEM = """You design teaching cases for an introductory health informatics course (HINF1009) in Ontario, Canada, where students analyze data in Microsoft Excel.

Return ONLY a JSON object. No markdown fences, no commentary. Schema:
{
  "client_name": "fictional person",
  "client_role": "job title",
  "organization": "fictional Ontario organization",
  "memo": "150-250 word memo from the client to the student analyst. Realistic and a little ambiguous, states the business problem in plain language. Never reveals the hidden relationships or any answers.",
  "questions": ["3 to 5 specific analysis tasks the student must complete"],
  "variables": [ ... ],
  "relationships": [ ... ]
}

Variable types (exact keys):
- {"name": "PatientID", "type": "id", "prefix": "P"}
- {"name": "...", "type": "numeric", "mean": 50, "sd": 10, "min": 0, "max": 100, "decimals": 0}
- {"name": "...", "type": "categorical", "categories": ["A","B"], "probabilities": [0.6, 0.4]}
- {"name": "...", "type": "binary", "base_rate": 0.15}      (values are "Yes"/"No")
- {"name": "...", "type": "date", "start": "YYYY-MM-DD", "end": "YYYY-MM-DD"}

Relationships are applied in the order listed; a target can be the source of a later relationship. Each has "target", "source", "effect", "description" (plain-language statement of the effect):
- numeric target, numeric source: effect is a number = change in target per 1-unit increase in source
- numeric target, categorical or binary source: effect is {"Category": shift} added to the target (reference category 0 or omitted)
- binary target, numeric source: effect is a number = risk multiplier per 1 standard deviation increase in source (e.g. 1.6)
- binary target, categorical or binary source: effect is {"Category": risk multiplier} (binary source categories are "Yes"/"No")

Rules:
- 6 to 10 variables, exactly one "id" variable, at most one "date" variable.
- 2 to 4 relationships. Sources and targets must be numeric, categorical or binary variables (never id or date). A variable cannot be its own source.
- Names are PascalCase with no spaces (e.g. WaitTimeMinutes).
- Effects must be realistic for healthcare, and large enough to be visible in a PivotTable at a few hundred rows.
- Include 1 or 2 variables with no relationship (distractors).
- The questions must be answerable from the variables."""


KEY_SYSTEM = """You write instructor answer keys for an introductory health informatics course (HINF1009). Students use Microsoft 365 Excel and are beginners.

Write in Markdown with these sections:
## Interpreting the request
## Solution (one subsection per question: Excel steps, expected result, suggested chart)
## Key findings for the client (3 to 5 plain-language bullets)
## Common student mistakes

Rules:
- Use ONLY the numbers in the computed results provided. Never invent or estimate numbers.
- Give exact formulas using the column letters provided and the sheet name Data, e.g. =AVERAGEIF(Data!C:C,"Male",Data!F:F).
- Prefer simple functions: AVERAGE, MEDIAN, MIN, MAX, STDEV.S, COUNT, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, CORREL.
- For PivotTables, say which fields go in Rows, Columns and Values, and the Values summary (Average, Count).
- Keep each step short and concrete."""


def secret(name):
    try:
        return st.secrets[name]
    except Exception:
        return None


def get_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])


def ask_claude(system, user, max_tokens=4000):
    msg = get_client().messages.create(
        model=MODEL,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    return "".join(b.text for b in msg.content if b.type == "text")


def parse_json(text):
    text = re.sub(r"```(?:json)?", "", text).strip()
    start, end = text.find("{"), text.rfind("}")
    return json.loads(text[start : end + 1])


def validate_spec(spec):
    vars_ = {v["name"]: v for v in spec["variables"]}
    allowed = {"id", "numeric", "categorical", "binary", "date"}
    for v in spec["variables"]:
        if v["type"] not in allowed:
            raise ValueError(f"Unknown type {v['type']} for {v['name']}")
        if v["type"] == "categorical" and len(v["categories"]) != len(v["probabilities"]):
            raise ValueError(f"{v['name']}: categories and probabilities differ in length")
    for r in spec["relationships"]:
        for key in ("target", "source"):
            if r[key] not in vars_:
                raise ValueError(f"Relationship refers to unknown variable {r[key]}")
            if vars_[r[key]]["type"] in ("id", "date"):
                raise ValueError(f"{r[key]} cannot be used in a relationship")
        if r["target"] == r["source"]:
            raise ValueError("A variable cannot be its own source")
        source_is_numeric = vars_[r["source"]]["type"] == "numeric"
        if source_is_numeric and not isinstance(r["effect"], (int, float)):
            raise ValueError(f"Effect for numeric source {r['source']} must be a number")
        if not source_is_numeric and not isinstance(r["effect"], dict):
            raise ValueError(f"Effect for categorical source {r['source']} must be an object")
        if vars_[r["target"]]["type"] == "categorical":
            raise ValueError("Targets must be numeric or binary")


def get_spec(scenario_text, difficulty_text):
    user = f"Scenario: {scenario_text}\n\nDifficulty: {difficulty_text}"
    last_error = None
    for _ in range(2):
        prompt = user if last_error is None else f"{user}\n\nYour previous JSON was invalid: {last_error}. Fix it."
        try:
            spec = parse_json(ask_claude(SPEC_SYSTEM, prompt))
            validate_spec(spec)
            return spec
        except (ValueError, KeyError, TypeError, json.JSONDecodeError) as e:
            last_error = str(e)
    raise RuntimeError(f"Could not get a valid data specification: {last_error}")


# --------------------------------------------------------------------------
# Synthetic data
# --------------------------------------------------------------------------

def generate_data(spec, n, seed):
    rng = np.random.default_rng(seed)
    df = pd.DataFrame(index=range(n))
    vars_ = {v["name"]: v for v in spec["variables"]}
    uniforms, multipliers = {}, {}

    for v in spec["variables"]:
        name, t = v["name"], v["type"]
        if t == "id":
            df[name] = [f"{v.get('prefix', 'ID')}{i + 1:05d}" for i in range(n)]
        elif t == "numeric":
            df[name] = rng.normal(v["mean"], v["sd"], n)
        elif t == "categorical":
            p = np.array(v["probabilities"], dtype=float)
            df[name] = rng.choice(v["categories"], size=n, p=p / p.sum())
        elif t == "binary":
            uniforms[name] = rng.random(n)
            multipliers[name] = np.ones(n)
            df[name] = np.where(uniforms[name] < v["base_rate"], "Yes", "No")
        elif t == "date":
            start, end = pd.Timestamp(v["start"]), pd.Timestamp(v["end"])
            days = max((end - start).days, 1)
            df[name] = start + pd.to_timedelta(rng.integers(0, days + 1, n), unit="D")

    for r in spec["relationships"]:
        t, s, effect = r["target"], r["source"], r["effect"]
        t_type, s_type = vars_[t]["type"], vars_[s]["type"]
        x = df[s]

        if t_type == "numeric":
            if s_type == "numeric":
                df[t] = df[t] + float(effect) * (x - x.mean())
            else:
                df[t] = df[t] + x.map(lambda c: float(effect.get(c, 0))).to_numpy()
        elif t_type == "binary":
            if s_type == "numeric":
                z = (x - x.mean()) / (x.std() or 1)
                multipliers[t] = multipliers[t] * np.power(float(effect), z.to_numpy())
            else:
                multipliers[t] = multipliers[t] * x.map(lambda c: float(effect.get(c, 1))).to_numpy()
            p = np.clip(vars_[t]["base_rate"] * multipliers[t], 0, 0.95)
            df[t] = np.where(uniforms[t] < p, "Yes", "No")

    for v in spec["variables"]:
        if v["type"] == "numeric":
            col = df[v["name"]].clip(v.get("min"), v.get("max")).round(int(v.get("decimals", 0)))
            df[v["name"]] = col.astype(int) if int(v.get("decimals", 0)) == 0 else col
        elif v["type"] == "date":
            df[v["name"]] = df[v["name"]].dt.date
    return df


def _text_variant(value, rng):
    if value in ("Yes", "No"):
        return rng.choice([value[0], value.lower(), value.upper()])
    return rng.choice([value.lower(), value.upper(), value + " ", " " + value])


def add_mess(df, spec, seed):
    """Messy copy for a data-cleaning exercise: text variants, blanks, duplicate rows."""
    rng = np.random.default_rng(seed + 1)
    messy = df.copy().astype(object)
    n = len(messy)
    for v in spec["variables"]:
        name, t = v["name"], v["type"]
        if t == "id":
            continue
        if t in ("categorical", "binary"):
            idx = rng.choice(n, size=int(0.05 * n), replace=False)
            messy.loc[idx, name] = [_text_variant(val, rng) for val in messy.loc[idx, name]]
        idx = rng.choice(n, size=max(1, int(0.02 * n)), replace=False)
        messy.loc[idx, name] = None
    dups = messy.sample(n=max(1, int(0.02 * n)), random_state=seed)
    return pd.concat([messy, dups]).sort_index(kind="stable").reset_index(drop=True)


def standard_clean(messy, spec):
    """Clean the student file the way students are taught to: remove duplicate rows,
    fix text variants, and leave blank cells empty (no filling in)."""
    df = messy.copy()
    for v in spec["variables"]:
        name, t = v["name"], v["type"]
        if t in ("categorical", "binary"):
            canon = v["categories"] if t == "categorical" else ["Yes", "No"]
            lookup = {str(c).strip().lower(): c for c in canon}
            if t == "binary":
                lookup.update({"y": "Yes", "n": "No"})
            df[name] = df[name].map(
                lambda x: lookup.get(str(x).strip().lower(), str(x).strip()) if pd.notna(x) else None
            )
        elif t == "numeric":
            df[name] = pd.to_numeric(df[name], errors="coerce")
    df = df.drop_duplicates().reset_index(drop=True)
    return df


# --------------------------------------------------------------------------
# Solution (computed from the clean data, so the key is always correct)
# --------------------------------------------------------------------------

def compute_solution(df, spec):
    vars_ = {v["name"]: v for v in spec["variables"]}
    numeric = [v["name"] for v in spec["variables"] if v["type"] == "numeric"]
    cats = [v["name"] for v in spec["variables"] if v["type"] in ("categorical", "binary")]
    tables = []

    if numeric:
        desc = df[numeric].agg(["count", "mean", "median", "std", "min", "max"]).T.round(2)
        desc.index.name = "Variable"
        tables.append(("Descriptive statistics", desc.reset_index()))

    for c in cats:
        freq = df[c].value_counts().rename_axis(c).reset_index(name="Count")
        freq["Percent"] = (freq["Count"] / df[c].notna().sum() * 100).round(1)
        tables.append((f"Frequency of {c}", freq))

    for r in spec["relationships"]:
        t, s = r["target"], r["source"]
        t_num = vars_[t]["type"] == "numeric"
        s_num = vars_[s]["type"] == "numeric"
        if t_num and not s_num:
            tbl = df.groupby(s)[t].agg(["count", "mean", "median"]).round(2).reset_index()
            title = f"Average {t} by {s}"
        elif t_num and s_num:
            tbl = pd.DataFrame({"Measure": [f"CORREL({s}, {t})"], "Value": [round(df[s].corr(df[t]), 3)]})
            title = f"Correlation of {s} with {t}"
        elif not s_num:
            g = df.dropna(subset=[t]).groupby(s)[t]
            tbl = pd.DataFrame({
                "Count": g.size(),
                f"{t} = Yes": g.apply(lambda x: (x == "Yes").sum()),
                "Percent Yes": g.apply(lambda x: round((x == "Yes").mean() * 100, 1)),
            }).reset_index()
            title = f"{t} rate by {s}"
        else:
            tbl = df.groupby(t)[s].agg(["count", "mean", "median"]).round(2).reset_index()
            title = f"Average {s} by {t}"
        tables.append((title, tbl))
    return tables


def column_letters(df):
    return {col: get_column_letter(i + 1) for i, col in enumerate(df.columns)}


def build_answer_key(spec, df, tables, messy):
    letters = ", ".join(f"{c} = column {l}" for c, l in column_letters(df).items())
    results = "\n\n".join(f"### {title}\n{tbl.to_string(index=False)}" for title, tbl in tables)
    note = (
        "The student file contains messy data: text variants (extra spaces, wrong capitalization, Y/N), "
        "blank cells, and exact duplicate rows. Add a 'Data cleaning' section before the solution "
        "explaining these steps in Excel: Remove Duplicates (Data tab), TRIM and Find & Replace to fix text, "
        "and leave blank cells empty. The computed results below come from the student file after exactly "
        "those steps, so a student who cleans correctly will match them. Blank cells are excluded from "
        "averages and counts; tell students a PivotTable may show a (blank) row, which they should ignore."
        if messy else "The student file is clean."
    )
    user = (
        f"Client memo:\n{spec['memo']}\n\nQuestions:\n"
        + "\n".join(f"{i + 1}. {q}" for i, q in enumerate(spec["questions"]))
        + f"\n\nData sheet: 'Data', {len(df)} rows after cleaning, plus a header row. Columns: {letters}\n\n"
        + f"{note}\n\nComputed results:\n{results}"
    )
    return ask_claude(KEY_SYSTEM, user, max_tokens=6000)


# --------------------------------------------------------------------------
# Case codes: everything needed to rebuild a case exactly, encrypted so students
# can't read the hidden relationships. The key comes from INSTRUCTOR_PASSWORD
# (or the API key if none is set), so changing that password invalidates old codes.
# --------------------------------------------------------------------------

CODE_PREFIX = "HCG1-"


def _fernet():
    secret_text = secret("INSTRUCTOR_PASSWORD") or secret("ANTHROPIC_API_KEY") or "local"
    key = base64.urlsafe_b64encode(hashlib.sha256(("hcg-case-code:" + secret_text).encode()).digest())
    return Fernet(key)


def make_case_code(spec, n_rows, messy, seed, name):
    payload = json.dumps(
        {"spec": spec, "n_rows": n_rows, "messy": messy, "seed": seed, "name": name},
        separators=(",", ":"),
    ).encode()
    return CODE_PREFIX + _fernet().encrypt(zlib.compress(payload, 9)).decode()


def read_case_code(code):
    code = re.sub(r"\s+", "", code or "")
    if not code:
        raise ValueError("Paste a case code into the box first.")
    if not code.startswith(CODE_PREFIX):
        raise ValueError("That doesn't look like a case code (it should start with HCG1-).")
    try:
        data = _fernet().decrypt(code[len(CODE_PREFIX):].encode())
    except InvalidToken:
        raise ValueError("The case code is damaged, or was made when a different instructor password was set.")
    return json.loads(zlib.decompress(data))


def build_case(spec, n_rows, messy, seed, name, with_key, status=None):
    say = status.write if status else (lambda _: None)
    say("Generating synthetic data...")
    clean = generate_data(spec, n_rows, seed)
    student_data = add_mess(clean, spec, seed) if messy else clean
    analysis_data, tables, key_md = clean, None, None
    if with_key:
        say("Computing the solution...")
        analysis_data = standard_clean(student_data, spec) if messy else clean
        tables = compute_solution(analysis_data, spec)
        say("Writing the answer key...")
        key_md = build_answer_key(spec, analysis_data, tables, messy)
    return dict(
        spec=spec, clean=clean, student=student_data, analysis=analysis_data, tables=tables,
        key=key_md, name=name, code=make_case_code(spec, n_rows, messy, seed, name),
    )


# --------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------

def _cell(value):
    if isinstance(value, np.generic):
        value = value.item()
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and value.isdigit():
        return int(value)  # e.g. triage level "3" stored as a number in Excel
    return value


def _write_table(ws, df, start_row=1, title=None):
    row = start_row
    if title:
        ws.cell(row=row, column=1, value=title).font = ARIAL_BOLD
        row += 1
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(row=row, column=j, value=str(col))
        c.font, c.fill = ARIAL_BOLD, HEADER_FILL
    for record in df.itertuples(index=False):
        row += 1
        for j, val in enumerate(record, start=1):
            c = ws.cell(row=row, column=j, value=_cell(val))
            c.font = ARIAL
            if hasattr(c.value, "year"):
                c.number_format = "yyyy-mm-dd"
    return row + 2


def _autowidth(ws):
    for col in ws.columns:
        width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 40)


def _write_text(ws, lines, width=100):
    ws.column_dimensions["A"].width = width
    for i, line in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=line)
        bold = line.startswith("#") or line.startswith("**")
        c.font = ARIAL_BOLD if bold else ARIAL
        c.alignment = Alignment(wrap_text=True, vertical="top")


def student_workbook(spec, data, code=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Client Request"
    lines = [
        f"From: {spec['client_name']}, {spec['client_role']}, {spec['organization']}",
        "",
        *spec["memo"].split("\n"),
        "",
        "**Your tasks**",
        *[f"{i + 1}. {q}" for i, q in enumerate(spec["questions"])],
    ]
    if code:
        lines += ["", "**Case code** (do not edit; your instructor uses it to check your work)", code]
    _write_text(ws, lines)

    ws2 = wb.create_sheet("Data")
    _write_table(ws2, data)
    ws2.freeze_panes = "A2"
    _autowidth(ws2)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def instructor_workbook(spec, clean, tables, key_md, original=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Answer Key"
    _write_text(ws, key_md.split("\n"))

    ws2 = wb.create_sheet("Results")
    row = 1
    for title, tbl in tables:
        row = _write_table(ws2, tbl, start_row=row, title=title)
    _autowidth(ws2)

    ws3 = wb.create_sheet("Hidden Relationships")
    rel = pd.DataFrame([
        {"Target": r["target"], "Source": r["source"], "Built-in effect": r["description"]}
        for r in spec["relationships"]
    ])
    _write_table(ws3, rel)
    _autowidth(ws3)

    ws4 = wb.create_sheet("Cleaned Data" if original is not None else "Data")
    _write_table(ws4, clean)
    ws4.freeze_panes = "A2"
    _autowidth(ws4)

    if original is not None:
        ws5 = wb.create_sheet("Original Data")
        ws5.cell(row=1, column=1,
                 value="Complete data before blanks were added; not recoverable by students.").font = ARIAL_BOLD
        _write_table(ws5, original, start_row=2)
        ws5.freeze_panes = "A3"
        _autowidth(ws5)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

def slug(text):
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")[:40]


def main():
    st.set_page_config(page_title="Health Analytics Case Generator", page_icon="🩺", layout="wide")
    st.title("🩺 Health Analytics Case Generator")
    st.caption("HINF1009: client request → synthetic data → solution")

    with st.sidebar:
        if not secret("ANTHROPIC_API_KEY"):
            st.error("Add ANTHROPIC_API_KEY to the app's secrets to generate cases.")
            st.stop()
        if secret("APP_PASSWORD"):
            if st.text_input("Class password", type="password") != secret("APP_PASSWORD"):
                st.info("Enter the class password to generate cases.")
                st.stop()

        choice = st.selectbox("Scenario", [*SCENARIOS, "Custom scenario..."])
        if choice == "Custom scenario...":
            scenario_text = st.text_area("Describe the scenario", "A long-term care home wants to ...")
        else:
            scenario_text = SCENARIOS[choice]
        difficulty = st.radio("Difficulty", list(DIFFICULTY))
        n_rows = st.slider("Rows of data", 100, 2000, 300, step=50)
        messy = st.checkbox("Add messy data for a cleaning exercise")

        # Instructor mode: without INSTRUCTOR_PASSWORD in secrets (e.g. running locally),
        # everyone sees the answer key. With it set, only the instructor does.
        instructor_pw = secret("INSTRUCTOR_PASSWORD")
        rebuild, code_in = False, ""
        go = st.button("Generate case", type="primary", use_container_width=True)
        with st.expander("Instructor"):
            if instructor_pw:
                entered = st.text_input("Instructor password", type="password")
                is_instructor = entered == instructor_pw
                if entered and not is_instructor:
                    st.error("Incorrect instructor password.")
            else:
                is_instructor = True
            if is_instructor:
                st.caption("Answer key included. Paste a student's case code to rebuild their case with its key.")
                code_in = st.text_area("Case code", height=100)
                rebuild = st.button("Rebuild case", use_container_width=True)

    case = st.session_state.get("case")
    add_key = False
    if is_instructor and case and case["key"] is None:
        st.info("This case was generated without the answer key.")
        add_key = st.button("Add answer key to this case")

    if add_key:
        code_in, rebuild = case["code"], True

    if go or rebuild:
        try:
            with st.status("Building the case...", expanded=True) as status:
                if rebuild:
                    info = read_case_code(code_in)
                    spec, rows, mess, seed, name = (info[k] for k in ("spec", "n_rows", "messy", "seed", "name"))
                else:
                    status.write("Writing the client request and data design...")
                    spec = get_spec(scenario_text, DIFFICULTY[difficulty])
                    rows, mess, name = n_rows, messy, slug(choice)
                    seed = int(np.random.default_rng().integers(1, 1_000_000))
                st.session_state["case"] = build_case(spec, rows, mess, seed, name, is_instructor, status)
                status.update(label="Case ready", state="complete")
        except Exception as e:  # show a friendly error, but keep showing the current case
            st.error(f"Something went wrong: {e}")

    case = st.session_state.get("case")
    if not case:
        st.info("Choose a scenario in the sidebar and select **Generate case**.")
        return

    spec = case["spec"]
    show_key = is_instructor and case["key"] is not None
    st.subheader(f"{spec['organization']}")
    with st.expander("Case code"):
        st.caption("This code is also in the student file. Submit it with your work so your instructor can check it.")
        st.code(case["code"], language=None, wrap_lines=True)

    c1, c2 = st.columns(2)
    c1.download_button(
        "⬇️ Student file (memo + data)",
        student_workbook(spec, case["student"], case["code"]),
        file_name=f"{case['name']}_student.xlsx",
        use_container_width=True,
    )
    if show_key:
        c2.download_button(
            "⬇️ Instructor file (key + results)",
            instructor_workbook(spec, case["analysis"], case["tables"], case["key"],
                                original=case["clean"] if case["analysis"] is not case["clean"] else None),
            file_name=f"{case['name']}_instructor.xlsx",
            use_container_width=True,
        )

    tab_names = ["Client request", "Data"] + (["Answer key", "Hidden relationships"] if show_key else [])
    tabs = st.tabs(tab_names)
    with tabs[0]:
        st.markdown(f"**From:** {spec['client_name']}, {spec['client_role']}")
        st.write(spec["memo"])
        st.markdown("**Tasks**")
        for i, q in enumerate(spec["questions"], start=1):
            st.markdown(f"{i}. {q}")
    with tabs[1]:
        st.dataframe(case["student"], use_container_width=True, height=400)
    if show_key:
        with tabs[2]:
            st.markdown(case["key"])
            for title, tbl in case["tables"]:
                st.markdown(f"**{title}**")
                st.dataframe(tbl, use_container_width=True, hide_index=True)
        with tabs[3]:
            for r in spec["relationships"]:
                st.markdown(f"- **{r['source']} → {r['target']}**: {r['description']}")


if __name__ == "__main__":
    main()
