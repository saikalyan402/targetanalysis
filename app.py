from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="Run Rate Analysis",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)

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

REQUIRED = {
    "Emp Code",
    "Employee Name",
    "FY 26 TGT EQ NS",
    "Equity NS Ach YTD June",
}


st.markdown(
    """
<style>
.stApp {
    background: #070707;
    color: #F3F0E7;
}

[data-testid="stSidebar"] {
    background: #0B0B0B;
    border-right: 1px solid rgba(212,175,55,.24);
}

[data-testid="stSidebar"] * {
    color: #F3F0E7;
}

.hero {
    border: 1px solid rgba(212,175,55,.25);
    border-radius: 22px;
    padding: 24px 26px;
    margin-bottom: 18px;
    background: linear-gradient(
        110deg,
        rgba(212,175,55,.10),
        rgba(255,255,255,.015)
    );
}

.eyebrow {
    color: #D4AF37;
    font-size: .76rem;
    letter-spacing: .17em;
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
    font-size: .96rem;
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
    font-size: .86rem;
    margin: 4px 0 13px;
    line-height: 1.5;
}

.kpi {
    border: 1px solid rgba(212,175,55,.24);
    border-radius: 17px;
    padding: 16px;
    min-height: 120px;
    background: linear-gradient(145deg, #121212, #0D0D0D);
}

.kpi-label {
    color: #A8A397;
    font-size: .72rem;
    letter-spacing: .04em;
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
    font-size: .72rem;
    margin-top: 7px;
    line-height: 1.4;
}

.callout {
    background: linear-gradient(
        145deg,
        rgba(212,175,55,.08),
        rgba(255,255,255,.01)
    );
    border: 1px solid rgba(212,175,55,.24);
    border-left: 3px solid #D4AF37;
    border-radius: 14px;
    padding: 15px 17px;
    color: #CBC6BA;
    font-size: .88rem;
    margin: 8px 0 18px;
    line-height: 1.65;
}

.management {
    border: 1px solid rgba(212,175,55,.24);
    border-radius: 16px;
    padding: 17px 18px;
    background: rgba(212,175,55,.055);
    color: #D7D2C7;
    line-height: 1.7;
    min-height: 175px;
}

.management b {
    color: #D4AF37;
}

[data-testid="stDataFrame"] {
    border: 1px solid rgba(212,175,55,.20);
    border-radius: 14px;
    overflow: hidden;
}
</style>
""",
    unsafe_allow_html=True,
)


def clean_text(value):
    if value is None:
        return ""

    if isinstance(value, float) and np.isnan(value):
        return ""

    return " ".join(str(value).strip().split())


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
            "Could not locate the RM header row in the uploaded workbook."
        )

    df = raw.iloc[header_row + 1:].copy()

    df.columns = unique_headers(
        raw.iloc[header_row].tolist()
    )

    df = df.dropna(how="all").copy()

    missing = sorted(REQUIRED - set(df.columns))

    if missing:
        raise ValueError(
            "Missing required columns: " + ", ".join(missing)
        )

    for column in [
        "FY 26 TGT EQ NS",
        "Equity NS Ach YTD June",
    ]:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        )

    for column in [
        "Emp Code",
        "Employee Name",
        "Status",
        "Type",
        "ZONE",
        "REGION",
    ]:
        if column in df.columns:
            df[column] = (
                df[column]
                .fillna("Unknown")
                .astype(str)
                .str.strip()
            )

    df["Emp Code"] = (
        df["Emp Code"]
        .str.replace(r"\.0$", "", regex=True)
    )

    if "MKT TYPE.1" in df.columns:
        market_source = "MKT TYPE.1"
    elif "MKT TYPE" in df.columns:
        market_source = "MKT TYPE"
    else:
        market_source = None

    if market_source:
        df["Market Type"] = (
            df[market_source]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
        )
    else:
        df["Market Type"] = "Unknown"

    market_normalization = {
        "B30-SELECT": "B30",
        "B30 SELECT": "B30",
        "B30 SELECTED": "B30",
        "T30-EXT": "T30",
        "T30 EXT": "T30",
        "T30 EXTENDED": "T30",
    }

    normalized_market = (
        df["Market Type"]
        .str.upper()
        .str.strip()
    )

    df["Market Type"] = normalized_market.replace(
        market_normalization
    )

    df.loc[
        df["Market Type"].eq(""),
        "Market Type",
    ] = "Unknown"

    return df


def project(df, months, uplift):
    result = df.copy()

    # Current achieved value represents April–June = 3 months
    result["Current Monthly RR"] = (
        result["Equity NS Ach YTD June"] / 3.0
    )

    result["Scenario Monthly RR"] = (
        result["Current Monthly RR"]
        * (1 + uplift / 100)
    )

    result["Projected Future NS"] = (
        result["Scenario Monthly RR"] * months
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


def summarize(
    frame,
    label,
    uplift,
    baseline,
):
    target = frame["FY 26 TGT EQ NS"].sum()
    projected = frame["Projected Final NS"].sum()

    qualifying = (
        frame["Projected Achievement %"] >= 100
    )

    baseline_qualifying = (
        baseline["Projected Achievement %"] >= 100
    )

    qualifying_ns = frame.loc[
        qualifying,
        "Projected Final NS",
    ].sum()

    baseline_projected = (
        baseline["Projected Final NS"].sum()
    )

    return {
        "Scenario": label,
        "RR Uplift %": uplift,
        "Total RMs": len(frame),
        "Target": target,
        "Projected NS": projected,
        "Portfolio Achievement %": (
            projected / target * 100
            if target
            else 0
        ),
        "RMs ≥100%": int(qualifying.sum()),
        "Qualification Rate %": (
            qualifying.mean() * 100
            if len(frame)
            else 0
        ),
        "New Qualifiers": int(
            (
                qualifying
                & ~baseline_qualifying
            ).sum()
        ),
        "Qualifying RM NS": qualifying_ns,
        "Qualifying NS Contribution %": (
            qualifying_ns / projected * 100
            if projected
            else 0
        ),
        "Revenue": projected * REVENUE_RATE,
        "Incremental NS": (
            projected - baseline_projected
        ),
        "Incremental Revenue": (
            projected - baseline_projected
        ) * REVENUE_RATE,
    }


def fmt(value):
    if value is None or pd.isna(value):
        return "—"

    return f"{float(value):,.2f}"


def pct(value):
    if value is None or pd.isna(value):
        return "—"

    return f"{float(value):,.1f}%"


def section(title, note=""):
    st.markdown(
        f"""
<div class="section-title">
    {title}
</div>
<div class="section-note">
    {note}
</div>
""",
        unsafe_allow_html=True,
    )


def kpi(label, value, foot=""):
    st.markdown(
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
""",
        unsafe_allow_html=True,
    )


def show_table(df, height=None):
    display = df.copy()

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


def chart_style(fig, height=390):
    fig.update_layout(
        template=None,
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor=CARD,
        font=dict(color=TEXT),
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

    return fig


st.markdown(
    """
<div class="hero">
    <div class="eyebrow">
        Executive Performance Review
    </div>

    <div class="hero-title">
        Run Rate Analysis
    </div>

    <div class="hero-sub">
        Upload the RM workbook to compare the current monthly
        run rate with +5%, +10%, +15% and a custom uplift.
        All results respond to the sidebar filters.
    </div>
</div>
""",
    unsafe_allow_html=True,
)


with st.sidebar:
    st.markdown("### Upload Data")

    uploaded = st.file_uploader(
        "Upload RM Workbook",
        type=["xlsx"],
    )


if uploaded is None:
    st.info(
        "Upload the RM Excel workbook from the sidebar to begin."
    )
    st.stop()


try:
    data = load_workbook(
        uploaded.getvalue()
    )

except Exception as error:
    st.error(
        f"Could not read workbook: {error}"
    )
    st.stop()


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


with st.sidebar:
    st.divider()

    st.markdown("### Global Filters")

    market_options = sorted(
        data["Market Type"]
        .dropna()
        .unique()
        .tolist()
    )

    selected_market = st.selectbox(
        "MKT TYPE",
        ["All"] + market_options,
    )

    if selected_market != "All":
        data = data[
            data["Market Type"] == selected_market
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

        options = sorted(
            data[column]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if (
            column == "Status"
            and "Active" in options
        ):
            default = ["Active"]
        else:
            default = options

        selected = st.multiselect(
            label,
            options,
            default=default,
        )

        data = data[
            data[column].isin(selected)
        ].copy()

    st.divider()

    months = st.select_slider(
        "Projection Months After June",
        options=list(MONTHS),
        value=9,
        format_func=lambda value: (
            f"{value} months · through {MONTHS[value]}"
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
        "No valid RMs remain after applying the selected filters."
    )
    st.stop()


scenario_specs = [
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
    label: project(
        data,
        months,
        uplift,
    )
    for label, uplift in scenario_specs
}


baseline = scenario_frames["Current"]


comparison = pd.DataFrame(
    [
        summarize(
            scenario_frames[label],
            label,
            uplift,
            baseline,
        )
        for label, uplift in scenario_specs
    ]
)


current = comparison.iloc[0]
custom = comparison.iloc[-1]

custom_frame = scenario_frames[
    f"Custom +{custom_uplift:.1f}%"
]


target_gap = max(
    current["Target"]
    - current["Projected NS"],
    0,
)

gap_closed = max(
    custom["Projected NS"]
    - current["Projected NS"],
    0,
)

if target_gap:
    gap_closed_pct = min(
        gap_closed / target_gap * 100,
        100,
    )
else:
    gap_closed_pct = 100


extra_rms = int(
    custom["RMs ≥100%"]
    - current["RMs ≥100%"]
)


st.markdown(
    f"""
<div class="callout">
    <b style="color:#D4AF37">
        Scope:
    </b>

    {selected_market} market ·
    {len(data):,} valid RMs ·
    projection through {MONTHS[months]}.

    Revenue is calculated as
    <b style="color:#F3F0E7">
        Net Sales × 0.60 / 100
    </b>.
</div>
""",
    unsafe_allow_html=True,
)


columns = st.columns(5)

cards = [
    (
        "Total Target",
        fmt(current["Target"]),
        f"{len(data):,} RMs in scope",
    ),
    (
        "Current Achievement",
        pct(current["Portfolio Achievement %"]),
        f"Projected NS {fmt(current['Projected NS'])}",
    ),
    (
        "Custom Achievement",
        pct(custom["Portfolio Achievement %"]),
        f"At +{custom_uplift:.1f}% run rate",
    ),
    (
        "Additional RMs ≥100%",
        f"+{extra_rms:,}",
        (
            f"{int(current['RMs ≥100%']):,} → "
            f"{int(custom['RMs ≥100%']):,}"
        ),
    ),
    (
        "Incremental Net Sales",
        fmt(custom["Incremental NS"]),
        (
            "Revenue impact "
            f"{fmt(custom['Incremental Revenue'])}"
        ),
    ),
]


for column, card in zip(
    columns,
    cards,
):
    with column:
        kpi(*card)


if custom["Portfolio Achievement %"] >= 100:
    conclusion = (
        "The custom scenario takes the filtered "
        "portfolio above its total target."
    )

elif gap_closed_pct >= 50:
    conclusion = (
        "The custom uplift closes a meaningful portion "
        "of the target gap, but additional action is "
        "still required."
    )

else:
    conclusion = (
        "Run-rate uplift alone is insufficient; "
        "the closest-to-target RMs should become "
        "the immediate action pool."
    )


st.markdown(
    f"""
<div class="callout">
    <b style="color:#D4AF37">
        Executive conclusion
    </b>

    <br><br>

    {conclusion}

    A +{custom_uplift:.1f}% run-rate increase adds
    <b>{fmt(custom["Incremental NS"])}</b>
    in projected NS, creates
    <b>{extra_rms:,} new qualifiers</b>,
    and closes
    <b>{pct(gap_closed_pct)}</b>
    of the current target gap.
</div>
""",
    unsafe_allow_html=True,
)


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


section(
    "2. Portfolio Movement",
    (
        "The left chart shows portfolio delivery; "
        "the right chart shows people conversion."
    ),
)


left, right = st.columns(
    2,
    gap="large",
)


with left:
    fig = go.Figure(
        go.Bar(
            x=comparison["Scenario"],
            y=comparison[
                "Portfolio Achievement %"
            ],
            marker_color=(
                ["#555555"]
                + [GOLD] * (len(comparison) - 1)
            ),
            text=[
                pct(value)
                for value in comparison[
                    "Portfolio Achievement %"
                ]
            ],
            textposition="outside",
        )
    )

    fig.add_hline(
        y=100,
        line_dash="dash",
        line_color=TEXT,
        annotation_text="100% target",
    )

    fig.update_layout(
        title="Portfolio Achievement by Scenario",
        xaxis_title="",
        yaxis_title="Achievement (%)",
    )

    st.plotly_chart(
        chart_style(fig),
        width="stretch",
        config={
            "displayModeBar": False,
        },
    )


with right:
    fig = go.Figure(
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

    fig.update_layout(
        title="RMs Crossing 100%",
        xaxis_title="",
        yaxis_title="Number of RMs",
    )

    st.plotly_chart(
        chart_style(fig),
        width="stretch",
        config={
            "displayModeBar": False,
        },
    )


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


near_mask = (
    baseline["Projected Achievement %"]
    .between(
        90,
        100,
        inclusive="left",
    )
)


near_count = int(
    near_mask.sum()
)


near_gap = (
    baseline.loc[
        near_mask,
        "FY 26 TGT EQ NS",
    ]
    -
    baseline.loc[
        near_mask,
        "Projected Final NS",
    ]
).clip(lower=0).sum()


section(
    "4. Market-Type Drivers",
    (
        f"Custom +{custom_uplift:.1f}% scenario "
        "contribution and qualification."
    ),
)


market_group = (
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


market_group["Portfolio Achievement %"] = np.where(
    market_group["Target"] > 0,
    (
        market_group["Projected NS"]
        / market_group["Target"]
        * 100
    ),
    0,
)


market_group["Qualification Rate %"] = (
    market_group["RMs ≥100%"]
    / market_group["#RMs"]
    * 100
)


market_group["NS Contribution %"] = (
    market_group["Projected NS"]
    / market_group["Projected NS"].sum()
    * 100
)


new_market = (
    custom_frame.loc[
        custom_frame["Qualifies"]
        & ~baseline["Qualifies"]
    ]
    .groupby("Market Type")
    .size()
    .rename("New Qualifiers")
    .reset_index()
)


market_group = market_group.merge(
    new_market,
    on="Market Type",
    how="left",
)


market_group["New Qualifiers"] = (
    market_group["New Qualifiers"]
    .fillna(0)
    .astype(int)
)


chart_column, table_column = st.columns(
    [1, 1.2],
    gap="large",
)


with chart_column:
    plot_data = market_group.sort_values(
        "Projected NS"
    )

    fig = go.Figure(
        go.Bar(
            x=plot_data["Projected NS"],
            y=plot_data["Market Type"],
            orientation="h",
            marker_color=GOLD,
            text=(
                plot_data["Projected NS"]
                .map(fmt)
            ),
            textposition="auto",
        )
    )

    fig.update_layout(
        title="Projected NS by Market Type",
        xaxis_title="Projected NS",
        yaxis_title="",
    )

    st.plotly_chart(
        chart_style(fig, 420),
        width="stretch",
        config={
            "displayModeBar": False,
        },
    )


with table_column:
    show_table(
        market_group.sort_values(
            "Projected NS",
            ascending=False,
        ),
        420,
    )


section(
    "5. Management Takeaways",
    (
        "Ready-to-present summary for the "
        "performance review."
    ),
)


left, right = st.columns(
    2,
    gap="large",
)


remaining_gap = max(
    custom["Target"]
    - custom["Projected NS"],
    0,
)


with left:
    st.markdown(
        f"""
<div class="management">
    <b>Portfolio story</b>

    <br><br>

    • Current achievement is
    {pct(current["Portfolio Achievement %"])}.

    <br>

    • +{custom_uplift:.1f}% run rate raises achievement to
    {pct(custom["Portfolio Achievement %"])}.

    <br>

    • Incremental NS is
    {fmt(custom["Incremental NS"])}.

    <br>

    • Incremental Revenue is
    {fmt(custom["Incremental Revenue"])}.

    <br>

    • Remaining target gap is
    {fmt(remaining_gap)}.
</div>
""",
        unsafe_allow_html=True,
    )


with right:
    st.markdown(
        f"""
<div class="management">
    <b>People story</b>

    <br><br>

    • Qualified RMs increase from
    {int(current["RMs ≥100%"]):,}
    to
    {int(custom["RMs ≥100%"]):,}.

    <br>

    • {extra_rms:,} RMs newly cross 100%.

    <br>

    • {near_count:,} RMs are currently in the
    90–100% conversion zone.

    <br>

    • Their combined gap to target is
    {fmt(near_gap)}.

    <br>

    • This group should receive the first
    focused intervention.
</div>
""",
        unsafe_allow_html=True,
    )


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

    detail = detail[
        detail_columns
    ].sort_values(
        [
            "Newly Qualifies",
            "Projected Achievement %",
        ],
        ascending=[
            False,
            False,
        ],
    )

    show_table(
        detail,
        550,
    )
