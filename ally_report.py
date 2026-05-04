import os
import time
from datetime import datetime

import pandas as pd
import requests
from dotenv import load_dotenv

# Load credentials from .env
load_dotenv()

# --- CONFIGURATION ---
CLIENT_ID = os.getenv("ALLY_CLIENT_ID")
API_TOKEN = os.getenv("ALLY_API_TOKEN")
ALLY_REGION = "prod.ally.ac"
MAPPING_FILE = "dept_mapping.csv"

if not API_TOKEN:
    print("❌ Error: ALLY_API_TOKEN not found in .env file.")
    exit()

BASE_URL = f"https://{ALLY_REGION}/api/v2/clients/{CLIENT_ID}"
headers = {"Authorization": f"Bearer {API_TOKEN}"}


def fetch_data(term_name):
    endpoint = f"{BASE_URL}/reports/overall?termName={term_name}"
    while True:
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 202:
            time.sleep(5)
            continue
        elif response.status_code == 200:
            data = response.json()
            if data.get("metadata", {}).get("status") == "Processing":
                time.sleep(5)
                continue
            return pd.DataFrame(data.get("data", []))
        else:
            return pd.DataFrame()


# --- INITIAL SETUP ---
term_a = input("Enter Baseline Term (e.g., 2025SP): ")
term_b = input("Enter Comparison Term (e.g., 2026SP): ")

if not os.path.exists(MAPPING_FILE):
    print(f"❌ Error: {MAPPING_FILE} not found.")
    exit()

folder_name = f"Reports_{datetime.now().strftime('%Y%m%d_%H%M')}"
os.makedirs(folder_name, exist_ok=True)

mapping = pd.read_csv(MAPPING_FILE, dtype={"departmentId": str})
all_schools = mapping["School"].dropna().unique().tolist()

# --- DATA ACQUISITION & FILTERING ---
print(f"🚀 Fetching data for {term_a} and {term_b}...")
df_a_raw = fetch_data(term_a)
df_b_raw = fetch_data(term_b)

# Filter out courses with zero files to ensure Ally scores are meaningful
if not df_a_raw.empty:
    df_a_raw = df_a_raw[df_a_raw["totalFiles"] > 0]
if not df_b_raw.empty:
    df_b_raw = df_b_raw[df_b_raw["totalFiles"] > 0]

for df in [df_a_raw, df_b_raw]:
    df["departmentId"] = df["departmentId"].astype(str)
    df["is_dl"] = df["departmentName"].str.endswith("-DL")

df_a_map = pd.merge(
    df_a_raw, mapping[["departmentId", "School"]], on="departmentId", how="left"
)
df_b_map = pd.merge(
    df_b_raw, mapping[["departmentId", "School"]], on="departmentId", how="left"
)


def get_modality_metrics(df, term_label):
    combined = (df.groupby("School")["overallScore"].mean() * 100).round(1)
    dl = (df[df["is_dl"]].groupby("School")["overallScore"].mean() * 100).round(1)
    non_dl = (df[~df["is_dl"]].groupby("School")["overallScore"].mean() * 100).round(1)
    return pd.DataFrame(
        {
            f"Comb_{term_label}": combined,
            f"DL_{term_label}": dl,
            f"NonDL_{term_label}": non_dl,
        }
    )


m_a = get_modality_metrics(df_a_map, term_a)
m_b = get_modality_metrics(df_b_map, term_b)
inst_context = pd.concat([m_a, m_b], axis=1).fillna(0)

inst_context["Comb_Growth"] = (
    inst_context[f"Comb_{term_b}"] - inst_context[f"Comb_{term_a}"]
).round(1)
inst_context["DL_Growth"] = (
    inst_context[f"DL_{term_b}"] - inst_context[f"DL_{term_a}"]
).round(1)
inst_context["NonDL_Growth"] = (
    inst_context[f"NonDL_{term_b}"] - inst_context[f"NonDL_{term_a}"]
).round(1)

ordered_cols = [
    f"Comb_{term_a}",
    f"Comb_{term_b}",
    "Comb_Growth",
    f"DL_{term_a}",
    f"DL_{term_b}",
    "DL_Growth",
    f"NonDL_{term_a}",
    f"NonDL_{term_b}",
    "NonDL_Growth",
]
inst_context = inst_context[ordered_cols].reset_index()

# --- LOOP THROUGH SCHOOLS ---
for target_school in all_schools:
    # Sub-filter data for the specific school
    school_depts_b = df_b_map[df_b_map["School"] == target_school]

    # SAFETY CHECK: Skip school if no active courses found for this term
    if school_depts_b.empty:
        print(
            f"⚠️ Skipping {target_school}: No active courses (with files) found in {term_b}."
        )
        continue

    print(f"📦 Generating: {target_school}...")
    school_slug = "".join([c for c in target_school if c.isalnum()])[:5]

    school_depts_a = df_a_map[df_a_map["School"] == target_school]

    dept_compare = (
        pd.DataFrame(
            {
                f"{term_a}_%": (
                    school_depts_a.groupby("departmentName")["overallScore"].mean()
                    * 100
                ).round(1),
                f"{term_b}_%": (
                    school_depts_b.groupby("departmentName")["overallScore"].mean()
                    * 100
                ).round(1),
            }
        )
        .reset_index()
        .fillna(0)
    )
    dept_compare["Growth"] = (
        dept_compare[f"{term_b}_%"] - dept_compare[f"{term_a}_%"]
    ).round(1)
    dept_compare = dept_compare.sort_values("Growth", ascending=False)

    granular_data = school_depts_b.copy()
    granular_data["Score_%"] = (granular_data["overallScore"] * 100).round(1)
    granular_data = granular_data[
        [
            "departmentName",
            "courseCode",
            "courseName",
            "Score_%",
            "totalFiles",
            "numberOfStudents",
        ]
    ].sort_values(["departmentName", "Score_%"], ascending=[True, False])

    filename = os.path.join(folder_name, f"Report_{school_slug}_{term_b}.xlsx")
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        workbook = writer.book
        highlight_fmt = workbook.add_format({"bg_color": "#FFEB9C"})
        green_fill_fmt = workbook.add_format(
            {"bg_color": "#E2EFDA", "font_color": "#006100", "bold": True, "border": 1}
        )

        sheets = [
            ("Institutional Context", inst_context, f"TblCtx{school_slug}"),
            ("Discipline Growth", dept_compare, f"TblGrw{school_slug}"),
            ("Granular Course Detail", granular_data, f"TblDet{school_slug}"),
        ]

        for sheet_name, df, t_name in sheets:
            rows, cols = df.shape
            df.to_excel(
                writer, sheet_name=sheet_name, index=False, header=False, startrow=1
            )
            worksheet = writer.sheets[sheet_name]

            # Table is only created if rows > 0, which our safety check ensures
            worksheet.add_table(
                0,
                0,
                rows,
                cols - 1,
                {
                    "name": t_name,
                    "columns": [{"header": c} for c in df.columns],
                    "style": "Table Style Medium 1",
                },
            )

            if sheet_name == "Institutional Context":
                for r_idx in range(rows):
                    row_pos = r_idx + 1
                    for g_col in [3, 6, 9]:
                        val = df.iloc[r_idx, g_col]
                        worksheet.write(row_pos, g_col, val, green_fill_fmt)

                    if df.iloc[r_idx]["School"] == target_school:
                        for c_idx in range(cols):
                            val = df.iloc[r_idx, c_idx]
                            fmt = (
                                green_fill_fmt if c_idx in [3, 6, 9] else highlight_fmt
                            )
                            worksheet.write(row_pos, c_idx, val, fmt)

                worksheet.set_column(0, 0, 35)
                worksheet.set_column(1, cols - 1, 16)
            else:
                worksheet.set_column(0, 0, 35)
                worksheet.set_column(1, cols - 1, 15)

            if sheet_name == "Institutional Context":
                chart = workbook.add_chart({"type": "column"})
                series_configs = [
                    {"idx": 1, "color": "#D9D9D9"},
                    {"idx": 2, "color": "#A6A6A6"},
                    {"idx": 3, "color": "#333333"},
                    {"idx": 4, "color": "#BDD7EE"},
                    {"idx": 5, "color": "#2F5597"},
                    {"idx": 6, "color": "#002060"},
                    {"idx": 7, "color": "#FCE4D6"},
                    {"idx": 8, "color": "#F4B084"},
                    {"idx": 9, "color": "#843C0C"},
                ]
                for config in series_configs:
                    chart.add_series(
                        {
                            "name": [sheet_name, 0, config["idx"]],
                            "categories": [sheet_name, 1, 0, rows, 0],
                            "values": [
                                sheet_name,
                                1,
                                config["idx"],
                                rows,
                                config["idx"],
                            ],
                            "fill": {"color": config["color"]},
                            "overlap": -5,
                            "gap": 150,
                            "data_labels": {"value": True, "font": {"size": 8}},
                        }
                    )
                chart.set_title({"name": f"Performance & Growth: {term_a} vs {term_b}"})
                chart.set_y_axis({"min": 0, "max": 100})
                dynamic_width = max(1200, (rows * 380))
                chart.set_size({"width": dynamic_width, "height": 600})
                chart.set_legend({"position": "bottom"})
                worksheet.insert_chart(rows + 5, 0, chart)

print(f"\n✅ All valid reports generated in: {folder_name}")
