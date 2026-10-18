import os
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from io import BytesIO
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def request_scroll_to_top():
    """Call right before st.rerun() after a major stage transition (saving
    Method Information, submitting Step 1, continuing from Step 2). Without
    this, the browser keeps its scroll pixel-position as a large collapsing
    section shrinks above it, landing mid-way into the next section instead
    of at its top."""
    st.session_state["_scroll_to_top_pending"] = True


def render_raw_html(html_text):
    """Render raw HTML/CSS via st.markdown, stripping each line's leading
    whitespace first. Without this, Streamlit's markdown parser can treat
    indented lines (e.g. closing </div> tags at the end of an f-string) as
    an indented markdown code block instead of passing them through as HTML."""
    stripped = "\n".join(line.strip() for line in html_text.split("\n"))
    st.markdown(stripped, unsafe_allow_html=True)


# -----------------------------
# Session state initialization
# -----------------------------
if "step1_completed" not in st.session_state:
    st.session_state.step1_completed = False

if "mv_plan_inputs" not in st.session_state:
    st.session_state.mv_plan_inputs = {}

if "proposed_characteristics" not in st.session_state:
    st.session_state.proposed_characteristics = []

if "selected_characteristics" not in st.session_state:
    st.session_state.selected_characteristics = []

if "plan_table" not in st.session_state:
    st.session_state.plan_table = []

if "method_info" not in st.session_state:
    st.session_state.method_info = {}

if "proposal_warnings" not in st.session_state:
    st.session_state.proposal_warnings = []

# -----------------------------
# Helper functions
# -----------------------------
def add_characteristic(characteristic_list, characteristic_name):
    if characteristic_name not in characteristic_list:
        characteristic_list.append(characteristic_name)


def safe_key(text):
    return (
        text.lower()
        .replace(" ", "_")
        .replace("/", "_")
        .replace("-", "_")
        .replace("(", "")
        .replace(")", "")
    )


def is_characteristic_saved(characteristic):
    return any(
        row["Performance Characteristic"] == characteristic
        for row in st.session_state.plan_table
    )


def generate_proposed_characteristics(mv_plan_inputs):
    proposed_characteristics = []

    if mv_plan_inputs["method_type"] == "Qualitative":
        add_characteristic(proposed_characteristics, "False positive / false negative rate")

    if mv_plan_inputs["method_type"] == "Quantitative":
        add_characteristic(proposed_characteristics, "Repeatability")
        add_characteristic(proposed_characteristics, "Working range / Reporting range")

    if mv_plan_inputs["matrix_interference"] == "Yes":
        add_characteristic(proposed_characteristics, "Selectivity / Specificity")

    if mv_plan_inputs["method_type"] == "Quantitative":
        if (
            mv_plan_inputs["crm_available"] == "Yes"
            or mv_plan_inputs["pt_available"] == "Yes"
        ):
            add_characteristic(proposed_characteristics, "Bias")

    if mv_plan_inputs["method_type"] == "Quantitative":
        if mv_plan_inputs["spiking_possible"] == "Yes":
            add_characteristic(proposed_characteristics, "Recovery from spiked blank matrix")

        if (
            mv_plan_inputs["spiking_possible"] == "Yes"
            and mv_plan_inputs["stable_samples"] == "Yes"
        ):
            add_characteristic(proposed_characteristics, "Recovery from spiked real matrix")

    if (
        mv_plan_inputs["method_type"] == "Quantitative"
        and mv_plan_inputs["calibration_model"] == "Yes"
    ):
        add_characteristic(proposed_characteristics, "Linearity / Calibration model")

    if mv_plan_inputs["method_type"] == "Quantitative":
        if mv_plan_inputs["reporting_range"] in [
            "Near LOQ / trace level",
            "Wide range (low to high)",
        ]:
            add_characteristic(proposed_characteristics, "LOD / LOQ")
            add_characteristic(proposed_characteristics, "Blank evaluation / blank controls")

    if (
        mv_plan_inputs["method_modified"] == "Yes"
        or mv_plan_inputs["multiple_changes"] == "Yes"
        or (
            mv_plan_inputs["method_type"] == "Quantitative"
            and mv_plan_inputs["method_category"] in [
                "Standard method with major modification",
                "In-house / developed method",
            ]
        )
    ):
        add_characteristic(proposed_characteristics, "Robustness")

    if (
        mv_plan_inputs["reference_method_available"] == "Yes"
        and mv_plan_inputs["can_run_standard"] == "Yes"
    ):
        add_characteristic(proposed_characteristics, "Comparison with reference / comparative method")

    if mv_plan_inputs["mu_estimation"] == "Yes":
        add_characteristic(proposed_characteristics, "Measurement uncertainty")

    if (
        mv_plan_inputs["method_type"] == "Quantitative"
        and (
            mv_plan_inputs["stable_samples"] == "Yes"
            or mv_plan_inputs["crm_available"] == "Yes"
            or mv_plan_inputs["spiking_possible"] == "Yes"
        )
    ):
        add_characteristic(proposed_characteristics, "Intermediate precision")

    return proposed_characteristics


def get_proposal_warnings(mv_plan_inputs):
    warnings = []

    if (
        mv_plan_inputs["method_type"] == "Quantitative"
        and mv_plan_inputs["crm_available"] != "Yes"
        and mv_plan_inputs["pt_available"] != "Yes"
        and mv_plan_inputs["spiking_possible"] != "Yes"
        and not (
            mv_plan_inputs["reference_method_available"] == "Yes"
            and mv_plan_inputs["can_run_standard"] == "Yes"
        )
    ):
        warnings.append(
            "No route to demonstrate trueness was identified (no CRM, PT results, "
            "spiking, or reference method that can be run). Consider sourcing a CRM, "
            "joining a proficiency testing scheme, or arranging an interlaboratory comparison."
        )

    return warnings


def get_fixed_design(characteristic):
    fixed_designs = {
        "Linearity / Calibration model": "Multi-point evaluation",
        "Blank evaluation / blank controls": "One-point / blank-control evaluation",
    }
    return fixed_designs.get(characteristic)


def get_evaluation_options(characteristic, method_type=None):
    if characteristic == "Selectivity / Specificity" and method_type == "Qualitative":
        return [
            "Inclusivity study using a panel of target analytes, variants or strains expected to give a positive result",
            "Exclusivity (cross-reactivity) study using a panel of closely related non-target compounds or organisms expected to give a negative result",
            "Evaluation of results in samples with and without potential matrix interferents or co-occurring components",
            "Analysis of characterised reference materials containing known non-target or interfering components",
        ]

    options = {
        "Repeatability": [
            "Within-run replicate study",
            "Single dataset standard deviation",
            "Difference-based duplicate approach",
            "Pooled standard deviation from multiple datasets",
        ],
        "Bias": [
            "Reference material comparison",
            "Proficiency testing result comparison",
            "Bias evaluation considering laboratory and reference value uncertainties",
        ],
        "Recovery from spiked blank matrix": [
            "Spiked blank recovery study",
            "Spiked pseudo-blank recovery study",
        ],
        "Recovery from spiked real matrix": [
            "Recovery evaluation using a representative real-matrix CRM",
            "Spiking of real samples",
            "Calibration using standard-addition approach",
        ],
        "Linearity / Calibration model": [
            "Comprehensive linearity evaluation (suitability of R², slope, intercept and residuals)",
            "Suitability of linear model using replicates at each calibration point",
        ],
        "LOD / LOQ": [
            "Blank-based approach",
            "Low-level sample approach",
            "Calibration curve approach",
            "Graphical approach",
        ],
        "Blank evaluation / blank controls": [
            "Reagent blank evaluation",
            "Matrix blank evaluation",
            "Field blank evaluation",
            "Blank CRM evaluation",
        ],
        "Robustness": [
            "Experimental design approach",
            "One-factor-at-a-time variation",
            "Routine condition variation assessment",
        ],
        "Comparison with reference / comparative method": [
            "Paired sample comparison",
            "Regression-based method comparison",
            "Direct comparison with reference method",
        ],
        "Intermediate precision": [
            "Between-day precision study",
            "Between-analyst precision study",
            "Between-run precision study",
            "ANOVA-based intermediate precision estimate",
        ],
        "Working range / Reporting range": [
            "Evaluated from linearity studies",
            "Evaluated by validated low and high levels for precision and bias",
        ],
        "Measurement uncertainty": [
            "Estimated from precision and bias data",
            "Use of modelling (bottom-up) approach",
        ],
        "Selectivity / Specificity": [
            "Evaluation of analyte response in the presence and absence of potential interferences",
            "Analysis of a matrix CRM containing known interferences at significant level",
            "Spiking calibration solutions with known interferences and evaluating impact on calibration parameters",
        ],
        "False positive / false negative rate": [
            "Positive/negative sample challenge study",
            "Comparison against confirmed reference results",
            "Classification performance assessment",
        ],
    }

    return options.get(characteristic, ["User-defined evaluation approach"])


# Example wording and example sources shown as guidance for the acceptance criterion.
# Illustrative only: the laboratory sets its own criterion.
CRITERION_GUIDE = {
    "Repeatability": (
        "RSDr ≤ 5 %",
        "AOAC Appendix F (expected RSDr by concentration; HorRat(r) check); "
        "IUPAC harmonised guidelines for single-laboratory validation; "
        "customer or method requirement",
    ),
    "Intermediate precision": (
        "Within-laboratory RSD ≤ 8 %",
        "ISO 5725-3; the sources listed for repeatability; customer or method requirement",
    ),
    "Bias": (
        "Mean CRM result within the certified value and its uncertainty, or PT |z| ≤ 2",
        "CRM certificate; ISO 13528 and ISO/IEC 17043 (PT evaluation)",
    ),
    "Recovery from spiked blank matrix": (
        "Mean recovery 90–110 %",
        "AOAC Appendix F (acceptable recovery by concentration); "
        "customer or regulatory requirement",
    ),
    "Recovery from spiked real matrix": (
        "Mean recovery 90–110 %",
        "AOAC Appendix F (acceptable recovery by concentration); "
        "customer or regulatory requirement",
    ),
    "Linearity / Calibration model": (
        "Residuals show no systematic trend; back-calculated points within ±x %",
        "Eurachem guide The Fitness for Purpose of Analytical Methods; the standard method",
    ),
    "Working range / Reporting range": (
        "Acceptable precision and recovery at both ends of the range",
        "Eurachem guide The Fitness for Purpose of Analytical Methods",
    ),
    "LOD / LOQ": (
        "LOQ ≤ the reporting or regulatory limit",
        "Eurachem guide The Fitness for Purpose of Analytical Methods; "
        "regulatory or customer limit",
    ),
    "Blank evaluation / blank controls": (
        "Blank response below x % of the reporting limit",
        "Eurachem supplement Blanks in Method Validation",
    ),
    "Selectivity / Specificity": (
        "Interferents change the result by no more than ±x %",
        "Eurachem guide The Fitness for Purpose of Analytical Methods; "
        "IUPAC harmonised guidelines for single-laboratory validation",
    ),
    "Comparison with reference / comparative method": (
        "No significant difference between the methods at 95 % confidence",
        "Eurachem guide The Fitness for Purpose of Analytical Methods",
    ),
    "Measurement uncertainty": (
        "Expanded uncertainty ≤ the target uncertainty",
        "Eurachem/CITAC guide Quantifying Uncertainty in Analytical Measurement; ISO 21748",
    ),
    "Robustness": (
        "Results unaffected by deliberate small variations in x and y",
        "Eurachem guide The Fitness for Purpose of Analytical Methods",
    ),
    "False positive / false negative rate": (
        "False negative rate ≤ x % at the threshold",
        "Eurachem/CITAC guide Assessment of performance and uncertainty in qualitative "
        "chemical analysis; ISO 16140 series (microbiology)",
    ),
}

QUALITATIVE_SELECTIVITY_GUIDE = (
    "All target strains or analytes detected; no non-target cross-reactivity",
    "ISO 16140 series (inclusivity and exclusivity studies); "
    "AOAC SMPR Guidelines (inclusivity and exclusivity panels)",
)


def get_criterion_guide(characteristic, method_type=None):
    if characteristic == "Selectivity / Specificity" and method_type == "Qualitative":
        return QUALITATIVE_SELECTIVITY_GUIDE

    return CRITERION_GUIDE.get(characteristic)


# Groups performance characteristics for display (Step 2 selection cards),
# following common Eurachem / ISO 5725 groupings. Order here is the display
# order; a group is shown only if at least one of its characteristics was
# proposed.
CHARACTERISTIC_CATEGORY_ORDER = [
    "Precision",
    "Trueness",
    "Selectivity / Specificity",
    "Range & Calibration",
    "Detection Capability",
    "Robustness",
    "Measurement Uncertainty",
]

CHARACTERISTIC_CATEGORIES = {
    "Repeatability": "Precision",
    "Intermediate precision": "Precision",
    "Bias": "Trueness",
    "Recovery from spiked blank matrix": "Trueness",
    "Recovery from spiked real matrix": "Trueness",
    "Comparison with reference / comparative method": "Trueness",
    "Selectivity / Specificity": "Selectivity / Specificity",
    "False positive / false negative rate": "Selectivity / Specificity",
    "Working range / Reporting range": "Range & Calibration",
    "Linearity / Calibration model": "Range & Calibration",
    "LOD / LOQ": "Detection Capability",
    "Blank evaluation / blank controls": "Detection Capability",
    "Robustness": "Robustness",
    "Measurement uncertainty": "Measurement Uncertainty",
}


def group_characteristics_by_category(characteristics):
    """Group a list of characteristic names into an ordered list of
    (category_label, [characteristics...]) tuples, following
    CHARACTERISTIC_CATEGORY_ORDER. Any characteristic without a known
    category falls into a trailing "Other" group, so nothing is ever
    silently dropped from the list."""
    grouped = {category: [] for category in CHARACTERISTIC_CATEGORY_ORDER}
    other = []

    for characteristic in characteristics:
        category = CHARACTERISTIC_CATEGORIES.get(characteristic)
        if category:
            grouped[category].append(characteristic)
        else:
            other.append(characteristic)

    ordered_groups = [
        (category, items) for category, items in grouped.items() if items
    ]
    if other:
        ordered_groups.append(("Other", other))

    return ordered_groups


def render_coverage_strip(selected_characteristics):
    """Render a checklist-style strip showing which of the standard
    performance-characteristic categories (CHARACTERISTIC_CATEGORY_ORDER) are
    covered by the currently selected characteristics, so the user can spot
    validation coverage gaps at a glance."""
    covered_categories = {
        CHARACTERISTIC_CATEGORIES.get(characteristic)
        for characteristic in selected_characteristics
        if CHARACTERISTIC_CATEGORIES.get(characteristic)
    }

    with st.container(border=True):
        st.markdown("**Validation Coverage**")
        for category in CHARACTERISTIC_CATEGORY_ORDER:
            is_covered = category in covered_categories
            icon = "✅" if is_covered else "⬜"
            text_color = "#1A1A1A" if is_covered else "#9E9E9E"
            render_raw_html(
                f'<div style="color: {text_color}; padding: 2px 0;">'
                f'{icon}&nbsp;&nbsp;{category}</div>'
            )


DESIGN_FIELD_LABELS = {
    "levels": "Number of levels / materials / items",
    "replicates": "Replicates per level / item",
    "days_runs": "Number of days / runs",
}

# Which experimental design fields apply to which characteristic
# (the text is the hover help shown for that field).
DESIGN_FIELDS = {
    "Repeatability": {
        "levels": "Concentration levels tested.",
        "replicates": "Replicates at each level.",
    },
    "Intermediate precision": {
        "levels": "Concentration levels or materials tested.",
        "replicates": "Replicates at each level (per day or run).",
        "days_runs": "Number of days or runs over which the study is spread.",
    },
    "Bias": {
        "levels": "Number of CRMs, PT rounds or levels evaluated.",
        "replicates": "Replicates per CRM, PT sample or level.",
        "days_runs": "Number of days or runs over which the material is measured.",
    },
    "Recovery from spiked blank matrix": {
        "levels": "Spiking levels.",
        "replicates": "Replicates at each spiking level.",
    },
    "Recovery from spiked real matrix": {
        "levels": "Spiking levels or materials.",
        "replicates": "Replicates at each spiking level or material.",
    },
    "Linearity / Calibration model": {
        "levels": "Calibration points.",
        "replicates": "Replicates at each calibration point.",
    },
    "Working range / Reporting range": {
        "levels": "Levels tested across the range.",
        "replicates": "Replicates at each level.",
    },
    "LOD / LOQ": {
        "levels": "Levels tested (1 for a one-point evaluation).",
        "replicates": "Replicates at each level.",
    },
    "Blank evaluation / blank controls": {
        "replicates": "Number of blanks analysed.",
        "days_runs": "Number of days or runs over which the blanks are analysed.",
    },
    "Selectivity / Specificity": {
        "levels": (
            "Quantitative: interferents or matrices tested. "
            "Qualitative: target and non-target items tested."
        ),
        "replicates": "Replicates per interferent, matrix or item.",
    },
    "Comparison with reference / comparative method": {
        "levels": "Number of samples compared.",
        "replicates": "Replicates per sample.",
    },
    "Measurement uncertainty": {},
    "Robustness": {
        "replicates": "Replicates per condition.",
        "days_runs": "Number of experimental runs.",
    },
    "False positive / false negative rate": {
        "levels": "Levels relative to the threshold (for example below, at and above).",
        "replicates": "Samples tested at each level.",
    },
}


def save_plan_row(
    characteristic,
    design,
    criteria,
    criterion,
    criterion_source,
    approaches,
    levels,
    replicates,
    days_runs,
    notes,
):
    new_row = {
        "Performance Characteristic": characteristic,
        "Evaluation design": design,
        "Performance criteria available?": criteria,
        "Acceptance criterion": criterion,
        "Criterion source / justification": criterion_source,
        "Evaluation approach": "; ".join(approaches),
        DESIGN_FIELD_LABELS["levels"]: levels,
        DESIGN_FIELD_LABELS["replicates"]: replicates,
        DESIGN_FIELD_LABELS["days_runs"]: days_runs,
        "Notes": notes,
    }

    st.session_state.plan_table = [
        row
        for row in st.session_state.plan_table
        if row["Performance Characteristic"] != characteristic
    ]

    st.session_state.plan_table.append(new_row)

DESIGN_COLUMNS = list(DESIGN_FIELD_LABELS.values())


def render_workload_estimate(plan_table):
    """Show a traceable estimate of total test results implied by the saved
    plan rows: levels x replicates per characteristic, each defaulting to 1
    when left blank (a one-point design implies 1 level; an unset replicate
    count implies 1 run)."""
    breakdown_rows = []
    total_test_results = 0

    for row in plan_table:
        levels = row.get(DESIGN_FIELD_LABELS["levels"]) or 1
        replicates = row.get(DESIGN_FIELD_LABELS["replicates"]) or 1
        test_results = levels * replicates
        total_test_results += test_results

        breakdown_rows.append({
            "Performance Characteristic": row["Performance Characteristic"],
            "Levels": levels,
            "Replicates": replicates,
            "Estimated test results": test_results,
        })

    st.subheader("Estimated Workload")
    st.metric("Total estimated test results", total_test_results)
    st.dataframe(
        pd.DataFrame(breakdown_rows),
        width="stretch",
        hide_index=True,
    )


def build_plan_dataframe(plan_table):
    plan_df = pd.DataFrame(plan_table)

    for column in DESIGN_COLUMNS:
        if column in plan_df.columns:
            plan_df[column] = plan_df[column].astype("Int64")

    return plan_df


# -----------------------------
# Excel formatting
# -----------------------------
TITLE_FONT = Font(bold=True, size=13, color="1F4E78")
TITLE_FILL = PatternFill("solid", fgColor="DCE6F1")
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP_ALIGNMENT = Alignment(wrap_text=True, vertical="top")
PLAIN_ALIGNMENT = Alignment(vertical="top")
THIN_SIDE = Side(style="thin", color="BFBFBF")
CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

# Columns long enough to need word-wrapping rather than a very wide column.
WRAP_COLUMN_NAMES = {
    "Value",
    "Evaluation approach",
    "Acceptance criterion",
    "Criterion source / justification",
    "Notes",
}

MIN_COLUMN_WIDTH = 12
MAX_COLUMN_WIDTH = 55

LOGO_PATH = "iso4lab_logo.png"
BRAND_TITLE_FONT = Font(bold=True, size=14, color="1B3042")
HEADER_ROWS = 4  # Rows reserved at the top of the sheet for the logo/title banner.


def _insert_header_banner(worksheet):
    """Add the ISO4LAB logo and title banner at the top of the sheet, if the
    logo file is available. Silently skips the logo (title still shown) if
    the file can't be found, so the export never fails because of it."""
    worksheet.row_dimensions[1].height = 18
    worksheet.row_dimensions[2].height = 18
    worksheet.row_dimensions[3].height = 18

    if os.path.exists(LOGO_PATH):
        logo = XLImage(LOGO_PATH)
        logo.width = 65
        logo.height = 61
        worksheet.add_image(logo, "A1")

    title_cell = worksheet.cell(row=2, column=2, value="ISO4LAB Method Validation Planning Tool")
    title_cell.font = BRAND_TITLE_FONT
    worksheet.merge_cells(start_row=2, start_column=2, end_row=2, end_column=6)


def _style_title_cell(worksheet, row, span_columns):
    cell = worksheet.cell(row=row, column=1)
    cell.font = TITLE_FONT
    cell.fill = TITLE_FILL
    if span_columns > 1:
        worksheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span_columns)
        for column in range(2, span_columns + 1):
            worksheet.cell(row=row, column=column).fill = TITLE_FILL


def _style_table(worksheet, header_row, n_rows, columns):
    n_cols = len(columns)

    for column_index in range(1, n_cols + 1):
        header_cell = worksheet.cell(row=header_row, column=column_index)
        header_cell.font = HEADER_FONT
        header_cell.fill = HEADER_FILL
        header_cell.alignment = HEADER_ALIGNMENT
        header_cell.border = CELL_BORDER

    for row_offset in range(1, n_rows + 1):
        row_index = header_row + row_offset
        wrap_in_this_row = False

        for column_index, column_name in enumerate(columns, start=1):
            cell = worksheet.cell(row=row_index, column=column_index)
            cell.border = CELL_BORDER

            if column_name in WRAP_COLUMN_NAMES:
                cell.alignment = WRAP_ALIGNMENT
                wrap_in_this_row = True
            else:
                cell.alignment = PLAIN_ALIGNMENT

        if wrap_in_this_row:
            longest_wrapped_value = max(
                (
                    len(str(worksheet.cell(row=row_index, column=columns.index(name) + 1).value or ""))
                    for name in columns
                    if name in WRAP_COLUMN_NAMES
                ),
                default=0,
            )
            estimated_lines = max(1, -(-longest_wrapped_value // max(MIN_COLUMN_WIDTH, 30)))
            worksheet.row_dimensions[row_index].height = min(15 * estimated_lines, 120)

    if n_rows > 0:
        worksheet.auto_filter.ref = (
            f"{get_column_letter(1)}{header_row}:{get_column_letter(n_cols)}{header_row + n_rows}"
        )


def _apply_column_widths(worksheet, dataframes):
    widths = {}

    for df in dataframes:
        for column_index, column_name in enumerate(df.columns, start=1):
            sample_values = [str(column_name)] + [str(v) for v in df[column_name].dropna().tolist()]
            longest = max((len(v) for v in sample_values), default=0)
            proposed_width = min(max(longest + 2, MIN_COLUMN_WIDTH), MAX_COLUMN_WIDTH)
            widths[column_index] = max(widths.get(column_index, 0), proposed_width)

    for column_index, width in widths.items():
        worksheet.column_dimensions[get_column_letter(column_index)].width = width


def create_excel_output(method_info, plan_table):
    output = BytesIO()

    method_info_rows = [
        {"Field": "Method name", "Value": method_info.get("method_name", "")},
        {"Field": "Method code", "Value": method_info.get("method_code", "")},
        {"Field": "Type of sample(s)", "Value": method_info.get("sample_types", "")},
        {"Field": "Sample preparation / digestion process", "Value": method_info.get("sample_preparation", "")},
        {"Field": "Analyte(s) / measurand(s)", "Value": method_info.get("analytes", "")},
        {"Field": "Analytical finish / reading technique", "Value": method_info.get("measurement_technique", "")},
    ]

    method_info_df = pd.DataFrame(method_info_rows)
    plan_df = build_plan_dataframe(plan_table)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet_name = "Validation Plan"

        method_info_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            startrow=1 + HEADER_ROWS
        )

        worksheet = writer.sheets[sheet_name]
        _insert_header_banner(worksheet)
        worksheet.cell(row=1 + HEADER_ROWS, column=1, value="Method Information")

        method_info_header_row = 2 + HEADER_ROWS
        _style_title_cell(worksheet, row=1 + HEADER_ROWS, span_columns=len(method_info_df.columns))
        _style_table(
            worksheet,
            header_row=method_info_header_row,
            n_rows=len(method_info_df),
            columns=list(method_info_df.columns),
        )

        plan_start_row = len(method_info_df) + 5 + HEADER_ROWS
        worksheet.cell(row=plan_start_row, column=1, value="Initial Validation Plan")

        plan_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            startrow=plan_start_row
        )

        plan_header_row = plan_start_row + 1
        _style_title_cell(worksheet, row=plan_start_row, span_columns=max(len(plan_df.columns), 1))

        if not plan_df.empty:
            _style_table(
                worksheet,
                header_row=plan_header_row,
                n_rows=len(plan_df),
                columns=list(plan_df.columns),
            )
            worksheet.freeze_panes = f"A{plan_header_row + 1}"
        else:
            for column_index in range(1, len(plan_df.columns) + 1):
                header_cell = worksheet.cell(row=plan_header_row, column=column_index)
                header_cell.font = HEADER_FONT
                header_cell.fill = HEADER_FILL
                header_cell.alignment = HEADER_ALIGNMENT
                header_cell.border = CELL_BORDER

        _apply_column_widths(worksheet, [method_info_df, plan_df])
        worksheet.sheet_view.showGridLines = False

    output.seek(0)
    return output

# -----------------------------
# Branding: card-style section styling
# -----------------------------
render_raw_html(
    """
    <style>
    /* Reserve space above any element the browser auto-scrolls to (e.g. a
    section revealed after saving), so it isn't hidden under the fixed
    progress stepper pinned at the top of the page. */
    section.stMain {
        scroll-padding-top: 110px;
    }

    /* Card-style sections for each expander */
    div[data-testid="stExpander"] {
        border: 1px solid #E0E0E0;
        border-left: 4px solid #1677A3;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
        margin-bottom: 1rem;
        overflow: hidden;
        scroll-margin-top: 110px;
    }

    /* Expander header row: tinted background, bolder label */
    div[data-testid="stExpander"] summary {
        background-color: #EAF2FA;
        font-weight: 600;
    }

    div[data-testid="stExpander"] summary:hover {
        background-color: #DCEBF7;
    }

    /* Accessible focus ring on inputs, using the brand primary color */
    input:focus, textarea:focus, select:focus {
        outline: 2px solid #1677A3 !important;
        outline-offset: 1px;
    }
    </style>
    """
)

# -----------------------------
# Progress stepper
# -----------------------------
# Rendered before the page title/logo (and pinned via position:fixed) so it
# acts as a persistent top bar: it stays visible while the branded header
# and everything else scrolls normally underneath it.
_stepper_stage_labels = [
    "Method Information",
    "Step 1: General Information",
    "Step 2: Characteristics",
    "Step 3: Validation Plan",
]
_stepper_selected_characteristics = st.session_state.selected_characteristics
_stepper_stage_complete = [
    bool(st.session_state.method_info),
    st.session_state.step1_completed,
    bool(_stepper_selected_characteristics),
    bool(_stepper_selected_characteristics)
    and all(is_characteristic_saved(c) for c in _stepper_selected_characteristics),
]
_stepper_current_index = next(
    (i for i, done in enumerate(_stepper_stage_complete) if not done),
    len(_stepper_stage_labels) - 1,
)

_stepper_n = len(_stepper_stage_labels)
_stepper_done_count = sum(1 for done in _stepper_stage_complete if done)
_stepper_progress_percent = (_stepper_done_count / (_stepper_n - 1)) * 75 if _stepper_n > 1 else 0

_stepper_cols_html = ""
for _i, _label in enumerate(_stepper_stage_labels):
    if _stepper_stage_complete[_i]:
        _circle, _state_class = "✓", "mvp-step-done"
    elif _i == _stepper_current_index:
        _circle, _state_class = str(_i + 1), "mvp-step-current"
    else:
        _circle, _state_class = str(_i + 1), "mvp-step-upcoming"

    _stepper_cols_html += f"""
    <div class="mvp-step {_state_class}">
        <div class="mvp-step-circle">{_circle}</div>
        <div class="mvp-step-label">{_label}</div>
    </div>
    """

render_raw_html(
    f"""
    <style>
    /* Fixed (not sticky) because Streamlit wraps each element in a
    tightly-fitted container, which gives position:sticky no room to
    "stick" within as the page scrolls. Fixed is anchored to the
    viewport instead, with a matching spacer below reserving its space
    in the normal document flow. */
    .mvp-stepper-spacer {{
        height: 98px;
    }}
    .mvp-stepper-wrap {{
        position: fixed;
        top: 60px;
        left: 0;
        right: 0;
        z-index: 999;
        background-color: #FFFFFF;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.06);
    }}
    .mvp-stepper-inner {{
        max-width: 900px;
        margin: 0 auto;
        padding: 0.75rem 1rem 0.6rem 1rem;
    }}
    /* Streamlit's default centered layout caps content at ~730px, which
    leaves large blank gutters either side when the app is embedded in a
    wide iframe (as on the iso4lab.com site). Widen it to 900px so the
    form fills the embedded space better, while still staying centered
    and readable rather than stretching edge-to-edge. */
    div[data-testid="stMainBlockContainer"], .block-container {{
        max-width: 900px !important;
    }}
    .mvp-stepper-track {{
        position: relative;
    }}
    .mvp-stepper-line-bg, .mvp-stepper-line-fg {{
        position: absolute;
        top: 14px;
        height: 2px;
        left: 12.5%;
    }}
    .mvp-stepper-line-bg {{
        right: 12.5%;
        background-color: #C7CDD3;
    }}
    .mvp-stepper-line-fg {{
        width: {_stepper_progress_percent}%;
        background-color: #1677A3;
        transition: width 0.2s ease;
    }}
    .mvp-stepper-cols {{
        display: grid;
        grid-template-columns: repeat({_stepper_n}, 1fr);
        position: relative;
    }}
    .mvp-step {{
        display: flex;
        flex-direction: column;
        align-items: center;
        text-align: center;
    }}
    .mvp-step-circle {{
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 600;
        font-size: 0.85rem;
        border: 2px solid #C7CDD3;
        color: #8A8F98;
        background-color: #FFFFFF;
        flex-shrink: 0;
    }}
    .mvp-step-label {{
        margin-top: 0.4rem;
        font-size: 0.78rem;
        color: #8A8F98;
        max-width: 85%;
        line-height: 1.2;
    }}
    .mvp-step-done .mvp-step-circle {{
        background-color: #1677A3;
        border-color: #1677A3;
        color: #FFFFFF;
    }}
    .mvp-step-done .mvp-step-label {{
        color: #1B3042;
        font-weight: 600;
    }}
    .mvp-step-current .mvp-step-circle {{
        border-color: #1677A3;
        color: #1677A3;
    }}
    .mvp-step-current .mvp-step-label {{
        color: #1B3042;
        font-weight: 600;
    }}
    </style>
    <div class="mvp-stepper-wrap">
        <div class="mvp-stepper-inner">
            <div class="mvp-stepper-track">
                <div class="mvp-stepper-line-bg"></div>
                <div class="mvp-stepper-line-fg"></div>
                <div class="mvp-stepper-cols">
                    {_stepper_cols_html}
                </div>
            </div>
        </div>
    </div>
    <div class="mvp-stepper-spacer"></div>
    """
)

# One-shot scroll-to-top, requested by request_scroll_to_top() right before
# a major-stage-transition rerun. Runs via a tiny same-origin component
# (st.markdown strips <script> tags, so it can't do this on its own).
if st.session_state.get("_scroll_to_top_pending"):
    st.session_state["_scroll_to_top_pending"] = False
    components.html(
        """
        <script>
        const mainSection = window.parent.document.querySelector('section.stMain');
        if (mainSection) { mainSection.scrollTop = 0; }
        </script>
        """,
        height=0,
    )

# -----------------------------
# Page title
# -----------------------------
header_logo_col, header_title_col = st.columns([1, 6], vertical_alignment="center")
with header_logo_col:
    st.image("iso4lab_logo.png")
with header_title_col:
    st.title("ISO4LAB Method Validation Planning Tool")

st.info(
    "This tool generates an initial method validation planning proposal based on the method "
    "information provided. The output is intended as a planning aid and should be reviewed by "
    "technically competent personnel before implementation."
)

# -----------------------------
# Step 1: Method General Information
# -----------------------------
# -----------------------------
# Method Information
# -----------------------------
method_info_saved = bool(st.session_state.method_info)
method_info_title = "Method Information"

with st.expander(method_info_title, expanded=not method_info_saved, icon="✅" if method_info_saved else None):
    with st.form("method_information_form"):
        method_info = {}

        method_info["method_name"] = st.text_input(
            "Method name",
            value=st.session_state.method_info.get("method_name", ""),
        )

        method_info["method_code"] = st.text_input(
            "Method code",
            value=st.session_state.method_info.get("method_code", ""),
        )

        method_info["sample_types"] = st.text_area(
            "Type of sample(s)",
            value=st.session_state.method_info.get("sample_types", ""),
        )

        method_info["sample_preparation"] = st.text_area(
            "Sample preparation / digestion process",
            value=st.session_state.method_info.get("sample_preparation", ""),
        )

        method_info["analytes"] = st.text_area(
            "Analyte(s) / measurand(s)",
            value=st.session_state.method_info.get("analytes", ""),
        )

        method_info["measurement_technique"] = st.text_input(
            "Analytical finish / reading technique",
            value=st.session_state.method_info.get("measurement_technique", ""),
        )

        save_method_info = st.form_submit_button("Save Method Information", type="primary")

    if save_method_info:
        required_method_fields = [
            "method_name",
            "method_code",
            "sample_types",
            "sample_preparation",
            "analytes",
            "measurement_technique",
        ]

        missing_method_fields = [
            field for field in required_method_fields
            if not method_info[field].strip()
        ]

        if missing_method_fields:
            st.error("Please complete the required method information fields before continuing.")
        else:
            st.session_state.method_info = method_info
            request_scroll_to_top()
            st.rerun()
        st.success("Method information saved.")

# -----------------------------
# Step 1: Method General Information
# -----------------------------
step1_saved = st.session_state.step1_completed
step1_title = "Step 1: Method General Information"

with st.expander(step1_title, expanded=not step1_saved, icon="✅" if step1_saved else None):

    selected_method_type = st.radio(
        "Type of method",
        ["Qualitative", "Quantitative"],
        index=None,
    )

    st.caption(
        "Select the type of method first. Questions that apply only to quantitative "
        "methods are hidden for qualitative methods."
    )

    with st.form("step1_form"):
        mv_plan_inputs = {}

        mv_plan_inputs["method_category"] = st.selectbox(
            "Method category",
            [
                "Standard method without change",
                "Standard method with minor modification",
                "Standard method with major modification",
                "In-house / developed method",
            ],
            index=None,
            placeholder="Select an option",
        )

        mv_plan_inputs["method_type"] = selected_method_type

        if selected_method_type != "Qualitative":
            mv_plan_inputs["reference_method_available"] = st.radio(
                "Is a suitable reference method or comparative method available?",
                ["Yes", "No"],
                index=None,
            )

        if selected_method_type == "Qualitative":
            interference_question = (
                "Could non-target compounds or organisms (closely related, cross-reacting or "
                "commonly co-occurring) or matrix components cause false positive or false "
                "negative results?"
            )
        else:
            interference_question = "Are there known matrix interferences or selectivity risks?"

        mv_plan_inputs["matrix_interference"] = st.radio(
            interference_question,
            ["Yes", "No"],
            index=None,
        )

        if selected_method_type != "Qualitative":
            mv_plan_inputs["reporting_range"] = st.selectbox(
                "Claimed reporting range",
                ["Near LOQ / trace level", "Mid to high level", "Wide range (low to high)"],
                index=None,
                placeholder="Select an option",
            )

            mv_plan_inputs["calibration_model"] = st.radio(
                "Is the method based on a calibration model?",
                ["Yes", "No"],
                index=None,
            )

        mv_plan_inputs["mu_estimation"] = st.radio(
            "Is measurement uncertainty to be estimated using validation data?",
            ["Yes", "No"],
            index=None,
        )

        if selected_method_type != "Qualitative":
            mv_plan_inputs["crm_available"] = st.radio(
                "Real-matrix CRM(s) available?",
                ["Yes", "No"],
                index=None,
            )

            mv_plan_inputs["pt_available"] = st.radio(
                "Are proficiency testing (PT) results available for this method and matrix?",
                ["Yes", "No"],
                index=None,
            )

            mv_plan_inputs["stable_samples"] = st.radio(
                "Availability of stable real samples?",
                ["Yes", "No"],
                index=None,
            )

            mv_plan_inputs["spiking_possible"] = st.radio(
                "Is spiking possible using standards or suitable reference solutions?",
                ["Yes", "No"],
                index=None,
            )

        mv_plan_inputs["method_modified"] = st.radio(
            "Has the laboratory applied any modifications?",
            ["Yes", "No"],
            index=None,
        )

        if selected_method_type != "Qualitative":
            mv_plan_inputs["can_run_standard"] = st.radio(
                "Is it possible to run the standard method without changes?",
                ["Yes", "No"],
                index=None,
            )

        mv_plan_inputs["multiple_changes"] = st.radio(
            "Has more than one factor in the method changed?",
            ["Yes", "No"],
            index=None,
        )

        submitted = st.form_submit_button("Generate Proposed Characteristics", type="primary")


    # -----------------------------
    # Step 1 validation and proposal generation
    # -----------------------------
    if submitted:
        if mv_plan_inputs["method_type"] == "Qualitative":
            for hidden_key in [
                "reference_method_available",
                "reporting_range",
                "calibration_model",
                "crm_available",
                "pt_available",
                "stable_samples",
                "spiking_possible",
                "can_run_standard",
            ]:
                mv_plan_inputs[hidden_key] = "Not applicable"

        unanswered_fields = [key for key, value in mv_plan_inputs.items() if value is None]

        if unanswered_fields:
            st.error("All questions are mandatory. Please answer all questions before continuing.")

        elif (
            mv_plan_inputs["method_category"] == "Standard method without change"
            and mv_plan_inputs["method_modified"] == "Yes"
        ):
            st.error(
                "Inconsistent answers: 'Standard method without change' cannot be combined with "
                "'Has the laboratory applied any modifications? = Yes'."
            )

        elif (
            mv_plan_inputs["method_modified"] == "No"
            and mv_plan_inputs["multiple_changes"] == "Yes"
        ):
            st.error(
                "Inconsistent answers: 'Has more than one factor in the method changed? = Yes' "
                "cannot be combined with 'Has the laboratory applied any modifications? = No'."
            )

        elif (
            mv_plan_inputs["method_category"] in [
                "Standard method with minor modification",
                "Standard method with major modification",
            ]
            and mv_plan_inputs["method_modified"] == "No"
        ):
            st.error(
                f"Inconsistent answers: '{mv_plan_inputs['method_category']}' cannot be combined "
                "with 'Has the laboratory applied any modifications? = No'."
            )

        elif (
            mv_plan_inputs["method_category"] == "Standard method without change"
            and mv_plan_inputs["can_run_standard"] == "No"
        ):
            st.error(
                "Inconsistent answers: 'Standard method without change' cannot be combined with "
                "'Is it possible to run the standard method without changes? = No'."
            )

        else:
            st.session_state.step1_completed = True
            st.session_state.mv_plan_inputs = mv_plan_inputs
            st.session_state.proposed_characteristics = generate_proposed_characteristics(mv_plan_inputs)
            st.session_state.proposal_warnings = get_proposal_warnings(mv_plan_inputs)
            st.session_state.selected_characteristics = []
            st.session_state.plan_table = []
            st.success("Proposed performance characteristics generated.")
            request_scroll_to_top()
            st.rerun()


# -----------------------------
# Step 2: Proposed Characteristics + User Override
# -----------------------------
if st.session_state.step1_completed:
    step2_saved = bool(st.session_state.selected_characteristics)
    step2_title = "Step 2: Proposed Performance Characteristics"

    with st.expander(step2_title, expanded=not step2_saved, icon="✅" if step2_saved else None):
        st.info(
            "Based on the information provided, the following performance characteristics "
            "are proposed. These are recommendations only. You may override this list by "
            "unticking any item that is not relevant to your validation study."
        )

        for warning_text in st.session_state.proposal_warnings:
            st.warning(warning_text)

        if st.session_state.proposed_characteristics:
            with st.form("characteristics_selection_form"):
                selected_characteristics = []

                for category_label, category_characteristics in group_characteristics_by_category(
                    st.session_state.proposed_characteristics
                ):
                    with st.container(border=True):
                        st.markdown(f"**{category_label}**")

                        for characteristic in category_characteristics:
                            include_item = st.checkbox(
                                characteristic,
                                value=True,
                                key=f"select_{safe_key(characteristic)}",
                            )

                            if include_item:
                                selected_characteristics.append(characteristic)

                continue_to_plan = st.form_submit_button("Continue to Initial Validation Plan", type="primary")

            if continue_to_plan:
                st.session_state.selected_characteristics = selected_characteristics

                # Keep already-saved rows for characteristics that are still selected;
                # only drop rows for ones that have since been unticked.
                st.session_state.plan_table = [
                    row
                    for row in st.session_state.plan_table
                    if row["Performance Characteristic"] in selected_characteristics
                ]

                if st.session_state.method_info:
                    st.subheader(
                        f"Initial Validation Plan for: {st.session_state.method_info.get('method_name', '')}"
                    )

                    st.markdown(
                        f"**Method code:** {st.session_state.method_info.get('method_code', '')}"
                    )

                if selected_characteristics:
                    st.success("Performance characteristics confirmed. Proceeding to initial validation plan.")
                    request_scroll_to_top()
                    st.rerun()
                else:
                    st.warning("No performance characteristics have been selected.")

        else:
            st.warning("No performance characteristics were proposed based on the current inputs.")


# -----------------------------
# Step 3: Initial Validation Plan Builder
# -----------------------------
if st.session_state.selected_characteristics:
    st.header("Step 3: Initial Validation Plan")

    if st.session_state.method_info:
        st.subheader(
            f"Initial Validation Plan for: {st.session_state.method_info.get('method_name', '')}"
        )
        st.markdown(
            f"**Method code:** {st.session_state.method_info.get('method_code', '')}"
        )

    st.info(
        "Complete the fields below to build an initial validation plan. "
        "The table will be updated after saving each performance characteristic."
    )

    render_coverage_strip(st.session_state.selected_characteristics)

    for characteristic in st.session_state.selected_characteristics:
        characteristic_saved = is_characteristic_saved(characteristic)
        expander_title = characteristic

        # Keep the section open once it has been saved, and also while the user is
        # actively choosing a design for it — otherwise picking "Multi-point evaluation"
        # (which sits outside the form, so the page can react to it) triggers a page
        # refresh that would otherwise snap the section shut before Save is clicked.
        # Stay open while the design is being chosen (a rerun happens as soon as it's
        # selected, since it sits outside the form), but collapse once saved.
        design_started = st.session_state.get(f"{safe_key(characteristic)}_design") is not None
        section_expanded = design_started and not characteristic_saved

        with st.expander(expander_title, expanded=section_expanded, icon="✅" if characteristic_saved else None):
            fixed_design = get_fixed_design(characteristic)

            if fixed_design:
                design = fixed_design
                st.markdown(f"**Evaluation design:** {fixed_design}")
            else:
                # Placed outside the form so the page updates immediately when the
                # design changes, which is needed to show/hide the "levels" field below.
                design = st.radio(
                    "One-point / multi-point evaluation",
                    [
                        "One-point evaluation",
                        "Multi-point evaluation",
                        "Not applicable / overall evaluation",
                    ],
                    index=None,
                    key=f"{safe_key(characteristic)}_design",
                )

            form_key = f"plan_form_{safe_key(characteristic)}"

            with st.form(form_key):
                criteria = st.radio(
                    "Performance criteria available?",
                    [
                        "Yes – criteria available from regulatory/reference/customer/technical source",
                        "No – user-defined criterion required",
                    ],
                    index=None,
                    key=f"{safe_key(characteristic)}_criteria",
                )

                criterion_guide = get_criterion_guide(
                    characteristic,
                    st.session_state.mv_plan_inputs.get("method_type"),
                )

                if criterion_guide:
                    criterion_placeholder = f"e.g. {criterion_guide[0]}"
                    criterion_help = (
                        f"**Example:** {criterion_guide[0]}\n\n"
                        f"**Example sources:** {criterion_guide[1]}\n\n"
                        "Illustrative only. Set the criterion to suit your method, "
                        "customer and regulatory requirements."
                    )
                else:
                    criterion_placeholder = "e.g. RSDr ≤ 5 %  or  recovery 90–110 %"
                    criterion_help = None

                criterion = st.text_input(
                    "Acceptance criterion",
                    placeholder=criterion_placeholder,
                    help=criterion_help,
                    key=f"{safe_key(characteristic)}_criterion",
                )

                criterion_source = st.text_input(
                    "Source / justification (optional)",
                    key=f"{safe_key(characteristic)}_criterion_source",
                )

                approaches = st.multiselect(
                    "How to evaluate",
                    get_evaluation_options(
                        characteristic,
                        st.session_state.mv_plan_inputs.get("method_type"),
                    ),
                    key=f"{safe_key(characteristic)}_approach",
                )

                design_fields = DESIGN_FIELDS.get(characteristic, {})
                design_values = {"levels": None, "replicates": None, "days_runs": None}

                # "Number of levels / materials / items" only applies once a multi-point
                # (or multi-level/multi-material) design has actually been chosen.
                visible_design_fields = {
                    field_key: help_text
                    for field_key, help_text in design_fields.items()
                    if field_key != "levels" or design == "Multi-point evaluation"
                }

                if visible_design_fields:
                    design_columns = st.columns(len(visible_design_fields))

                    for column, (field_key, help_text) in zip(design_columns, visible_design_fields.items()):
                        with column:
                            design_values[field_key] = st.number_input(
                                DESIGN_FIELD_LABELS[field_key],
                                min_value=1,
                                step=1,
                                value=None,
                                help=help_text + " Leave blank if not applicable.",
                                key=f"{safe_key(characteristic)}_{field_key}",
                            )

                notes = st.text_area(
                    "Notes",
                    key=f"{safe_key(characteristic)}_notes",
                )

                save_characteristic = st.form_submit_button(f"Save {characteristic}", type="primary")

            if save_characteristic:
                missing_fields = []

                if not design:
                    missing_fields.append("One-point / multi-point evaluation")

                if criteria is None:
                    missing_fields.append("Performance criteria availability")

                if not criterion.strip():
                    missing_fields.append("Acceptance criterion")

                if not approaches:
                    missing_fields.append("Evaluation approach")

                if missing_fields:
                    st.error(
                        "Please complete the following fields before saving: "
                        + ", ".join(missing_fields)
                    )
                else:
                    save_plan_row(
                        characteristic,
                        design,
                        criteria,
                        criterion.strip(),
                        criterion_source.strip(),
                        approaches,
                        design_values["levels"],
                        design_values["replicates"],
                        design_values["days_runs"],
                        notes,
                    )
                    st.success(f"{characteristic} added to the initial validation plan.")
                    st.rerun()


# -----------------------------
# Live validation plan table
# -----------------------------
if st.session_state.plan_table:
    st.header("Initial Validation Plan Table")

    plan_df = build_plan_dataframe(st.session_state.plan_table)

    st.dataframe(
        plan_df,
        width="stretch",
        hide_index=True,
    )

    render_workload_estimate(st.session_state.plan_table)

    excel_file = create_excel_output(
        st.session_state.method_info,
        st.session_state.plan_table
    )

    st.download_button(
        label="Download Initial Validation Plan as Excel",
        data=excel_file,
        file_name="initial_validation_plan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
