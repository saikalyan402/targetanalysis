from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Run Rate Analysis",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CONSTANTS
# ============================================================

GOLD = "#D4AF37"
TEXT = "#F3F0E7"
MUTED = "#A8A397"
CARD = "#111111"

# Revenue = Net Sales × 0.60 / 100
REVENUE_RATE = 0.006

MONTHS = {
    1: "July",
    2: "August",
    3: "September",
    4: "October",
    5: "November",
    6: "December",
    7: "January",
    8: "February",
    9: "March",
}

REQUIRED_COLUMNS = {
    "Emp Code",
    "Employee Name",
    "FY 26 TGT EQ NS",
    "Equity NS Ach YTD June",
}


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
<style>

.stApp {
    background: #070707;
    color: #F3F0E7;
}

[data-testid="stSidebar"] {
    background: #0B0B0B;
    border-right: 1px solid rgba(212, 175, 55, 0.24);
}

[data-testid="stSidebar"] * {
    color: #F3F0E7;
}

.hero {
    border: 1px solid rgba(212, 175, 55, 0.25);
    border-radius: 22px;
    padding: 24px 26px;
    margin-bottom: 18px;
    background: linear-gradient(
        110deg,
        rgba(212, 175, 55, 0.10),
        rgba(255, 255, 255, 0.015)
    );
}

.eyebrow {
    color: #D4AF37;
    font-size: 0.76rem;
    letter-spacing: 0.17em;
    font-weight: 750;
    text-transform: uppercase;
}

.hero-title {
    font-size: clamp(2rem, 3vw, 3.1rem);
    line-height: 1.05;
    font-weight: 780;
    margin-top: 8px;
}

.hero-sub {
    color: #A8A397;
    margin-top: 10px;
    font-size: 0.96rem;
    line-height: 1.6;
    max-width: 1100px;
}

.section-title {
    font-size: 1.28rem;
    font-weight: 750;
    margin-top: 18px;
}

.section-note {
    color: #A8A397;
    font-size: 0.86rem;
    margin: 4px 0 13px;
    line-height: 1.5;
}

.kpi {
    border: 1px solid rgba(212, 175, 55, 0.24);
    border-radius: 17px;
    padding: 16px;
    min-height: 120px;
    background: linear-gradient(
        145deg,
        #121212,
        #0D0D0D
    );
}

.kpi-label {
    color: #A8A397;
    font-size: 0.72rem;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    font-weight: 650;
}

.kpi-value {
    font-size: 1.62rem;
    font-weight: 780;
    margin-top: 7px;
    color: #D4AF37;
}

.kpi-foot {
    color: #A8A397;
    font-size: 0.72rem;
    margin-top: 7px;
    line-height: 1.4;
}

.callout {
    background: linear-gradient(
        145deg,
        rgba(212, 175, 55, 0.08),
        rgba(255, 255, 255, 0.01)
    );
    border: 1px solid rgba(212, 175, 55, 0.24);
    border-left: 3px solid #D4AF37;
    border-radius: 14px;
    padding: 15px 17px;
    color: #CBC6BA;
    font-size: 0.88rem;
    margin: 8px 0 18px;
    line-height: 1.65;
}

.management {
    border: 1px solid rgba(212, 175, 55, 0.24);
    border-radius: 16px;
    padding: 17px 18px;
    background: rgba(212, 175, 55, 0.055);
    color: #D7D2C7;
    line-height: 1.7;
    min-height: 175px;
}

.management b {
    color: #D4AF37;
}

[data-testid="stDataFrame"] {
    border: 1px solid rgba(212, 175, 55, 0.20);
    border-radius: 14px;
    overflow: hidden;
}

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# DATA CLEANING HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    if isinstance(value, float) and np.isnan(value):
        return ""

    return " ".join(
        str(value).strip().split()
    )


def unique_headers(values):
    seen = {}
    output = []

    for index, value in enumerate(values):
        base = clean_text(value) or f"Unnamed_{index}"
        count = seen.get(base, 0)

        if count == 0:
            output.append(base)
        else:
            output.append(f"{base}.{count}")

        seen[base] = count + 1

    return output


# ============================================================
# EXCEL READER
# ============================================================

@st.cache_data(show_spinner=False)
def load_workbook(file_bytes):
    raw = pd.read_excel(
        BytesIO(file_bytes),
        header=None,
        engine="openpyxl",
    )

    header_row = None

    for index in range(min(30, len(raw))):
        values = {
            clean_text(value)
            for value in raw.iloc[index].tolist()
        }

        if (
            "FY 26 TGT EQ NS" in values
            and "Equity NS Ach YTD June" in values
        ):
            header_row = index
            break

    if header_row is None:
        raise ValueError(
            "Could not locate the RM header row in the workbook."
        )

    dataframe = raw.iloc[
        header_row + 1:
    ].copy()

    dataframe.columns = unique_headers(
        raw.iloc[header_row].tolist()
    )

    dataframe = dataframe.dropna(
        how="all"
    ).copy()

    missing_columns = sorted(
        REQUIRED_COLUMNS - set(dataframe.columns)
    )

    if missing_columns:
        raise ValueError(
            "Missing required columns: "
            + ", ".join(missing_columns)
        )

    for column in [
        "FY 26 TGT EQ NS",
        "Equity NS Ach YTD June",
    ]:
        dataframe[column] = pd.to_numeric(
            dataframe[column],
            errors="coerce",
        )

    text_columns = [
        "Emp Code",
        "Employee Name",
        "Status",
        "Type",
        "ZONE",
        "REGION",
    ]

    for column in text_columns:
        if column in dataframe.columns:
            dataframe[column] = (
                dataframe[column]
                .fillna("Unknown")
                .astype(str)
                .str.strip()
            )

    dataframe["Emp Code"] = (
        dataframe["Emp Code"]
        .str.replace(
            r"\.0$",
            "",
            regex=True,
        )
    )

    # Use the second market-type column where available.
    if "MKT TYPE.1" in dataframe.columns:
        market_source = "MKT TYPE.1"

    elif "MKT TYPE" in dataframe.columns:
        market_source = "MKT TYPE"

    else:
        market_source = None

    if market_source:
        dataframe["Market Type"] = (
            dataframe[market_source]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
        )
    else:
        dataframe["Market Type"] = "Unknown"

    # Combine B30 variants and T30 variants.
    market_normalization = {
        "B30-SELECT": "B30",
        "B30 SELECT": "B30",
        "B30 SELECTED": "B30",
        "B30_SELECT": "B30",
        "T30-EXT": "T30",
        "T30 EXT": "T30",
        "T30 EXTENDED": "T30",
        "T30_EXT": "T30",
    }

    normalized_market = (
        dataframe["Market Type"]
        .str.upper()
        .str.strip()
    )

    dataframe["Market Type"] = (
        normalized_market.replace(
            market_normalization
        )
    )

    dataframe.loc[
        dataframe["Market Type"].eq(""),
        "Market Type",
    ] = "Unknown"

    return dataframe


# ============================================================
# RUN-RATE PROJECTION
# ============================================================

def project_run_rate(
    dataframe,
    projection_months,
    uplift_percentage,
):
    result = dataframe.copy()

    # April–June represents three completed months.
    result["Current Monthly RR"] = (
        result["Equity NS Ach YTD June"] / 3.0
    )

    result["Scenario Monthly RR"] = (
        result["Current Monthly RR"]
        * (
            1
            + uplift_percentage / 100
        )
    )

    result["Projected Future NS"] = (
        result["Scenario Monthly RR"]
        * projection_months
    )

    result["Projected Final NS"] = (
        result["Equity NS Ach YTD June"]
        + result["Projected Future NS"]
    )

    result["Projected Achievement %"] = np.where(
        result["FY 26 TGT EQ NS"] > 0,
        (
            result["Projected Final NS"]
            / result["FY 26 TGT EQ NS"]
            * 100
        ),
        np.nan,
    )

    result["Qualifies"] = (
        result["Projected Achievement %"] >= 100
    )

    return result


# ============================================================
# SCENARIO SUMMARY
# ============================================================

def summarize_scenario(
    scenario_dataframe,
    scenario_name,
    uplift_percentage,
    baseline_dataframe,
):
    total_target = (
        scenario_dataframe[
            "FY 26 TGT EQ NS"
        ].sum()
    )

    total_projected_ns = (
        scenario_dataframe[
            "Projected Final NS"
        ].sum()
    )

    qualifying_mask = (
        scenario_dataframe[
            "Projected Achievement %"
        ] >= 100
    )

    baseline_qualifying_mask = (
        baseline_dataframe[
            "Projected Achievement %"
        ] >= 100
    )

    qualifying_ns = (
        scenario_dataframe.loc[
            qualifying_mask,
            "Projected Final NS",
        ].sum()
    )

    baseline_projected_ns = (
        baseline_dataframe[
            "Projected Final NS"
        ].sum()
    )

    incremental_ns = (
        total_projected_ns
        - baseline_projected_ns
    )

    return {
        "Scenario": scenario_name,
        "RR Uplift %": uplift_percentage,
        "Total RMs": len(scenario_dataframe),
        "Target": total_target,
        "Projected NS": total_projected_ns,
        "Portfolio Achievement %": (
            total_projected_ns
            / total_target
            * 100
            if total_target
            else 0
        ),
        "RMs ≥100%": int(
            qualifying_mask.sum()
        ),
        "Qualification Rate %": (
            qualifying_mask.mean() * 100
            if len(scenario_dataframe)
            else 0
        ),
        "New Qualifiers": int(
            (
                qualifying_mask
                & ~baseline_qualifying_mask
            ).sum()
        ),
        "Qualifying RM NS": qualifying_ns,
        "Qualifying NS Contribution %": (
            qualifying_ns
            / total_projected_ns
            * 100
            if total_projected_ns
            else 0
        ),
        "Revenue": (
            total_projected_ns
            * REVENUE_RATE
        ),
        "Incremental NS": incremental_ns,
        "Incremental Revenue": (
            incremental_ns
            * REVENUE_RATE
        ),
    }


# ============================================================
# DISPLAY HELPERS
# ============================================================

def format_number(value):
    if value is None or pd.isna(value):
        return "—"

    return f"{float(value):,.2f}"


def format_percentage(value):
    if value is None or pd.isna(value):
        return "—"

    return f"{float(value):,.1f}%"


def section(title, note=""):
    st.html(
        f"""
<div class="section-title">
    {title}
</div>

<div class="section-note">
    {note}
</div>
"""
    )


def kpi(label, value, foot=""):
    st.html(
        f"""
<div class="kpi">
    <div class="kpi-label">
        {label}
    </div>

    <div class="kpi-value">
        {value}
    </div>

    <div class="kpi-foot">
        {foot}
    </div>
</div>
"""
    )


def show_table(
    dataframe,
    height=None,
):
    display = dataframe.copy()

    numeric_columns = (
        display
        .select_dtypes(include=np.number)
        .columns
    )

    display[numeric_columns] = (
        display[numeric_columns].round(2)
    )

    arguments = {
        "data": display,
        "width": "stretch",
        "hide_index": True,
    }

    if height is not None:
        arguments["height"] = height

    st.dataframe(**arguments)


def apply_chart_style(
    figure,
    height=390,
):
    figure.update_layout(
        template=None,
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor=CARD,
        font=dict(
            color=TEXT,
        ),
        margin=dict(
            l=45,
            r=30,
            t=65,
            b=45,
        ),
        xaxis=dict(
            gridcolor="rgba(212,175,55,.10)",
            zeroline=False,
        ),
        yaxis=dict(
            gridcolor="rgba(212,175,55,.10)",
            zeroline=False,
        ),
        hoverlabel=dict(
            bgcolor="#111111",
            font_color=TEXT,
        ),
    )

    return figure


# ============================================================
# PAGE HEADER
# ============================================================

st.html(
    """
<div class="hero">
    <div class="eyebrow">
        Executive Performance Review
    </div>

    <div class="hero-title">
        Run Rate Analysis
    </div>

    <div class="hero-sub">
        Upload the RM workbook to compare the current
        monthly run rate with +5%, +10%, +15% and a
        custom uplift. All results respond to the
        sidebar filters.
    </div>
</div>
"""
)


# ============================================================
# FILE UPLOAD
# ============================================================

with st.sidebar:
    st.markdown("### Upload Data")

    uploaded_file = st.file_uploader(
        "Upload RM Workbook",
        type=["xlsx"],
        help=(
            "Upload the workbook containing the "
            "RM Equity Net Sales data."
        ),
    )


if uploaded_file is None:
    st.info(
        "Upload the RM Excel workbook from the sidebar "
        "to begin the analysis."
    )
    st.stop()


try:
    data = load_workbook(
        uploaded_file.getvalue()
    )

except Exception as error:
    st.error(
        f"Could not read workbook: {error}"
    )
    st.stop()


# ============================================================
# VALIDATE ROWS
# ============================================================

valid_identity = (
    data["Employee Name"].ne("")
    | data["Emp Code"].ne("")
)

valid_numeric = (
    data["FY 26 TGT EQ NS"].notna()
    & data["Equity NS Ach YTD June"].notna()
)

positive_target = (
    data["FY 26 TGT EQ NS"] > 0
)

data = data[
    valid_identity
    & valid_numeric
    & positive_target
].copy()


# ============================================================
# SIDEBAR FILTERS
# ============================================================

with st.sidebar:
    st.divider()

    st.markdown("### Global Filters")

    market_options = sorted(
        data["Market Type"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    selected_market = st.selectbox(
        "MKT TYPE",
        ["All"] + market_options,
    )

    if selected_market != "All":
        data = data[
            data["Market Type"]
            == selected_market
        ].copy()

    filter_definitions = [
        ("Status", "Status"),
        ("Type", "Type"),
        ("ZONE", "Zone"),
        ("REGION", "Region"),
    ]

    for column, label in filter_definitions:
        if column not in data.columns:
            continue

        available_options = sorted(
            data[column]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if (
            column == "Status"
            and "Active" in available_options
        ):
            default_options = ["Active"]
        else:
            default_options = available_options

        selected_options = st.multiselect(
            label,
            available_options,
            default=default_options,
        )

        data = data[
            data[column].isin(
                selected_options
            )
        ].copy()

    st.divider()

    projection_months = st.select_slider(
        "Projection Months After June",
        options=list(MONTHS),
        value=9,
        format_func=lambda value: (
            f"{value} months · "
            f"through {MONTHS[value]}"
        ),
    )

    custom_uplift = st.number_input(
        "Custom Run Rate Increase (%)",
        min_value=0.0,
        max_value=500.0,
        value=20.0,
        step=1.0,
    )


if data.empty:
    st.warning(
        "No valid RMs remain after applying "
        "the selected filters."
    )
    st.stop()


# ============================================================
# CREATE SCENARIOS
# ============================================================

scenario_definitions = [
    ("Current", 0.0),
    ("+5%", 5.0),
    ("+10%", 10.0),
    ("+15%", 15.0),
    (
        f"Custom +{custom_uplift:.1f}%",
        custom_uplift,
    ),
]

scenario_frames = {
    scenario_name: project_run_rate(
        data,
        projection_months,
        uplift_percentage,
    )
    for (
        scenario_name,
        uplift_percentage,
    ) in scenario_definitions
}

baseline = scenario_frames["Current"]

custom_scenario_name = (
    f"Custom +{custom_uplift:.1f}%"
)

custom_frame = scenario_frames[
    custom_scenario_name
]

comparison = pd.DataFrame(
    [
        summarize_scenario(
            scenario_frames[scenario_name],
            scenario_name,
            uplift_percentage,
            baseline,
        )
        for (
            scenario_name,
            uplift_percentage,
        ) in scenario_definitions
    ]
)

current_summary = comparison.iloc[0]
custom_summary = comparison.iloc[-1]


# ============================================================
# EXECUTIVE CALCULATIONS
# ============================================================

current_target_gap = max(
    current_summary["Target"]
    - current_summary["Projected NS"],
    0,
)

gap_closed = max(
    custom_summary["Projected NS"]
    - current_summary["Projected NS"],
    0,
)

if current_target_gap:
    gap_closed_percentage = min(
        gap_closed
        / current_target_gap
        * 100,
        100,
    )
else:
    gap_closed_percentage = 100

additional_qualifying_rms = int(
    custom_summary["RMs ≥100%"]
    - current_summary["RMs ≥100%"]
)


# ============================================================
# SCOPE
# ============================================================

st.html(
    f"""
<div class="callout">
    <b style="color:#D4AF37">
        Scope
    </b>

    <br><br>

    {selected_market} market ·
    {len(data):,} valid RMs ·
    projection through {MONTHS[projection_months]}.

    <br><br>

    Revenue is calculated as:

    <b style="color:#F3F0E7">
        Net Sales × 0.60 / 100
    </b>
</div>
"""
)


# ============================================================
# EXECUTIVE KPI CARDS
# ============================================================

kpi_columns = st.columns(5)

kpi_cards = [
    (
        "Total Target",
        format_number(
            current_summary["Target"]
        ),
        f"{len(data):,} RMs in scope",
    ),
    (
        "Current Achievement",
        format_percentage(
            current_summary[
                "Portfolio Achievement %"
            ]
        ),
        (
            "Projected NS "
            + format_number(
                current_summary[
                    "Projected NS"
                ]
            )
        ),
    ),
    (
        "Custom Achievement",
        format_percentage(
            custom_summary[
                "Portfolio Achievement %"
            ]
        ),
        (
            f"At +{custom_uplift:.1f}% "
            "run rate"
        ),
    ),
    (
        "Additional RMs ≥100%",
        f"+{additional_qualifying_rms:,}",
        (
            f"{int(current_summary['RMs ≥100%']):,}"
            " → "
            f"{int(custom_summary['RMs ≥100%']):,}"
        ),
    ),
    (
        "Incremental Net Sales",
        format_number(
            custom_summary["Incremental NS"]
        ),
        (
            "Revenue impact "
            + format_number(
                custom_summary[
                    "Incremental Revenue"
                ]
            )
        ),
    ),
]

for column, card in zip(
    kpi_columns,
    kpi_cards,
):
    with column:
        kpi(*card)


# ============================================================
# EXECUTIVE CONCLUSION
# ============================================================

if (
    custom_summary[
        "Portfolio Achievement %"
    ] >= 100
):
    conclusion = (
        "The custom scenario takes the filtered "
        "portfolio above its total target."
    )

elif gap_closed_percentage >= 50:
    conclusion = (
        "The custom uplift closes a meaningful "
        "portion of the target gap, but additional "
        "action is still required."
    )

else:
    conclusion = (
        "Run-rate uplift alone is insufficient. "
        "The closest-to-target RMs should become "
        "the immediate action pool."
    )

st.html(
    f"""
<div class="callout">
    <b style="color:#D4AF37">
        Executive conclusion
    </b>

    <br><br>

    {conclusion}

    <br><br>

    A +{custom_uplift:.1f}% run-rate increase adds

    <b style="color:#F3F0E7">
        {format_number(custom_summary["Incremental NS"])}
    </b>

    in projected NS, creates

    <b style="color:#F3F0E7">
        {additional_qualifying_rms:,} new qualifiers
    </b>

    and closes

    <b style="color:#F3F0E7">
        {format_percentage(gap_closed_percentage)}
    </b>

    of the current target gap.
</div>
"""
)


# ============================================================
# 1. SCENARIO SCORECARD
# ============================================================

section(
    "1. Scenario Scorecard",
    (
        "Management comparison of achievement, "
        "qualification, Net Sales and Revenue."
    ),
)

scorecard_columns = [
    "Scenario",
    "RR Uplift %",
    "Projected NS",
    "Portfolio Achievement %",
    "RMs ≥100%",
    "Qualification Rate %",
    "New Qualifiers",
    "Incremental NS",
    "Revenue",
    "Incremental Revenue",
]

show_table(
    comparison[scorecard_columns]
)


# ============================================================
# 2. PORTFOLIO MOVEMENT
# ============================================================

section(
    "2. Portfolio Movement",
    (
        "The left chart shows portfolio delivery. "
        "The right chart shows people conversion."
    ),
)

left_chart, right_chart = st.columns(
    2,
    gap="large",
)

with left_chart:
    achievement_chart = go.Figure(
        go.Bar(
            x=comparison["Scenario"],
            y=comparison[
                "Portfolio Achievement %"
            ],
            marker_color=(
                ["#555555"]
                + [GOLD]
                * (len(comparison) - 1)
            ),
            text=[
                format_percentage(value)
                for value in comparison[
                    "Portfolio Achievement %"
                ]
            ],
            textposition="outside",
        )
    )

    achievement_chart.add_hline(
        y=100,
        line_dash="dash",
        line_color=TEXT,
        annotation_text="100% target",
        annotation_position="top left",
    )

    achievement_chart.update_layout(
        title="Portfolio Achievement by Scenario",
        xaxis_title="",
        yaxis_title="Achievement (%)",
    )

    st.plotly_chart(
        apply_chart_style(
            achievement_chart
        ),
        width="stretch",
        config={
            "displayModeBar": False,
        },
    )

with right_chart:
    qualification_chart = go.Figure(
        go.Scatter(
            x=comparison["Scenario"],
            y=comparison["RMs ≥100%"],
            mode="lines+markers+text",
            text=(
                comparison["RMs ≥100%"]
                .astype(int)
            ),
            textposition="top center",
            line=dict(
                color=GOLD,
                width=3,
            ),
            marker=dict(
                color=GOLD,
                size=10,
            ),
        )
    )

    qualification_chart.update_layout(
        title="RMs Crossing 100%",
        xaxis_title="",
        yaxis_title="Number of RMs",
    )

    st.plotly_chart(
        apply_chart_style(
            qualification_chart
        ),
        width="stretch",
        config={
            "displayModeBar": False,
        },
    )


# ============================================================
# 3. CONVERSION OPPORTUNITY
# ============================================================

section(
    "3. Conversion Opportunity",
    (
        "RMs nearest to 100% offer the fastest "
        "qualification opportunity."
    ),
)

opportunity = baseline.copy()

opportunity["Achievement Band"] = pd.cut(
    opportunity["Projected Achievement %"],
    bins=[
        0,
        80,
        90,
        95,
        100,
        np.inf,
    ],
    labels=[
        "Below 80%",
        "80–90%",
        "90–95%",
        "95–100%",
        "100%+",
    ],
    right=False,
)

band_table = (
    opportunity
    .groupby(
        "Achievement Band",
        observed=False,
    )
    .agg(
        **{
            "#RMs": (
                "Employee Name",
                "size",
            ),
            "Target": (
                "FY 26 TGT EQ NS",
                "sum",
            ),
            "Projected NS": (
                "Projected Final NS",
                "sum",
            ),
        }
    )
    .reset_index()
)

band_table["Share of RMs %"] = (
    band_table["#RMs"]
    / len(data)
    * 100
)

show_table(
    band_table
)

near_target_mask = (
    baseline["Projected Achievement %"]
    .between(
        90,
        100,
        inclusive="left",
    )
)

near_target_count = int(
    near_target_mask.sum()
)

near_target_gap = (
    baseline.loc[
        near_target_mask,
        "FY 26 TGT EQ NS",
    ]
    -
    baseline.loc[
        near_target_mask,
        "Projected Final NS",
    ]
).clip(lower=0).sum()


# ============================================================
# 4. MARKET-TYPE DRIVERS
# ============================================================

section(
    "4. Market-Type Drivers",
    (
        f"Custom +{custom_uplift:.1f}% scenario "
        "contribution and qualification."
    ),
)

market_summary = (
    custom_frame
    .groupby(
        "Market Type",
        dropna=False,
    )
    .agg(
        **{
            "#RMs": (
                "Employee Name",
                "size",
            ),
            "Target": (
                "FY 26 TGT EQ NS",
                "sum",
            ),
            "Projected NS": (
                "Projected Final NS",
                "sum",
            ),
            "RMs ≥100%": (
                "Qualifies",
                "sum",
            ),
        }
    )
    .reset_index()
)

market_summary[
    "Portfolio Achievement %"
] = np.where(
    market_summary["Target"] > 0,
    (
        market_summary["Projected NS"]
        / market_summary["Target"]
        * 100
    ),
    0,
)

market_summary[
    "Qualification Rate %"
] = np.where(
    market_summary["#RMs"] > 0,
    (
        market_summary["RMs ≥100%"]
        / market_summary["#RMs"]
        * 100
    ),
    0,
)

total_market_ns = (
    market_summary["Projected NS"].sum()
)

market_summary[
    "NS Contribution %"
] = np.where(
    total_market_ns != 0,
    (
        market_summary["Projected NS"]
        / total_market_ns
        * 100
    ),
    0,
)

new_qualifiers_by_market = (
    custom_frame.loc[
        custom_frame["Qualifies"]
        & ~baseline["Qualifies"]
    ]
    .groupby("Market Type")
    .size()
    .rename("New Qualifiers")
    .reset_index()
)

market_summary = market_summary.merge(
    new_qualifiers_by_market,
    on="Market Type",
    how="left",
)

market_summary["New Qualifiers"] = (
    market_summary["New Qualifiers"]
    .fillna(0)
    .astype(int)
)

market_chart_column, market_table_column = (
    st.columns(
        [1, 1.2],
        gap="large",
    )
)

with market_chart_column:
    market_chart_data = (
        market_summary
        .sort_values("Projected NS")
    )

    market_chart = go.Figure(
        go.Bar(
            x=market_chart_data[
                "Projected NS"
            ],
            y=market_chart_data[
                "Market Type"
            ],
            orientation="h",
            marker_color=GOLD,
            text=(
                market_chart_data[
                    "Projected NS"
                ]
                .map(format_number)
            ),
            textposition="auto",
        )
    )

    market_chart.update_layout(
        title="Projected NS by Market Type",
        xaxis_title="Projected NS",
        yaxis_title="",
    )

    st.plotly_chart(
        apply_chart_style(
            market_chart,
            420,
        ),
        width="stretch",
        config={
            "displayModeBar": False,
        },
    )

with market_table_column:
    show_table(
        market_summary.sort_values(
            "Projected NS",
            ascending=False,
        ),
        420,
    )


# ============================================================
# 5. MANAGEMENT TAKEAWAYS
# ============================================================

section(
    "5. Management Takeaways",
    (
        "Ready-to-present summary for the "
        "performance review."
    ),
)

remaining_target_gap = max(
    custom_summary["Target"]
    - custom_summary["Projected NS"],
    0,
)

portfolio_column, people_column = st.columns(
    2,
    gap="large",
)

with portfolio_column:
    st.html(
        f"""
<div class="management">
    <b>Portfolio story</b>

    <br><br>

    • Current achievement is
    {format_percentage(
        current_summary["Portfolio Achievement %"]
    )}.

    <br>

    • +{custom_uplift:.1f}% run rate raises
    achievement to
    {format_percentage(
        custom_summary["Portfolio Achievement %"]
    )}.

    <br>

    • Incremental Net Sales is
    {format_number(
        custom_summary["Incremental NS"]
    )}.

    <br>

    • Incremental Revenue is
    {format_number(
        custom_summary["Incremental Revenue"]
    )}.

    <br>

    • Remaining target gap is
    {format_number(remaining_target_gap)}.
</div>
"""
    )

with people_column:
    st.html(
        f"""
<div class="management">
    <b>People story</b>

    <br><br>

    • Qualified RMs increase from
    {int(current_summary["RMs ≥100%"]):,}
    to
    {int(custom_summary["RMs ≥100%"]):,}.

    <br>

    • {additional_qualifying_rms:,}
    RMs newly cross 100%.

    <br>

    • {near_target_count:,}
    RMs are currently in the
    90–100% conversion zone.

    <br>

    • Their combined gap to target is
    {format_number(near_target_gap)}.

    <br>

    • This group should receive the
    first focused intervention.
</div>
"""
    )


# ============================================================
# RM-LEVEL DRILL-DOWN
# ============================================================

with st.expander(
    "RM-Level Drill-Down"
):
    detail = custom_frame.copy()

    detail["Current Achievement %"] = (
        baseline["Projected Achievement %"]
    )

    detail["Newly Qualifies"] = (
        detail["Qualifies"]
        & ~baseline["Qualifies"]
    )

    detail_columns = [
        column
        for column in [
            "Emp Code",
            "Employee Name",
            "Market Type",
            "ZONE",
            "REGION",
            "FY 26 TGT EQ NS",
            "Equity NS Ach YTD June",
            "Current Monthly RR",
            "Current Achievement %",
            "Scenario Monthly RR",
            "Projected Final NS",
            "Projected Achievement %",
            "Newly Qualifies",
        ]
        if column in detail.columns
    ]

    detail = (
        detail[detail_columns]
        .sort_values(
            [
                "Newly Qualifies",
                "Projected Achievement %",
            ],
            ascending=[
                False,
                False,
            ],
        )
    )

    show_table(
        detail,
        550,
    )
