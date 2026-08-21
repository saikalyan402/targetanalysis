from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# PAGE CONFIG + THEME
# ============================================================

st.set_page_config(
    page_title="RM Equity NS Strategy Lab",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)

GOLD = "#D4AF37"
GOLD2 = "#BFA24A"
TEXT = "#F3F0E7"
MUTED = "#9F9B90"
CARD = "#111111"
GRID = "rgba(212,175,55,.12)"
BORDER = "rgba(212,175,55,.24)"

st.markdown(
    """
<style>
.stApp {background:#070707;color:#F3F0E7;}
[data-testid="stSidebar"] {background:#0B0B0B;border-right:1px solid rgba(212,175,55,.24);}
[data-testid="stSidebar"] * {color:#F3F0E7;}
.hero {border:1px solid rgba(212,175,55,.24);border-radius:22px;padding:24px 26px;margin-bottom:18px;background:linear-gradient(110deg,rgba(212,175,55,.10),rgba(255,255,255,.015));}
.eyebrow {color:#D4AF37;font-size:.76rem;letter-spacing:.17em;font-weight:750;text-transform:uppercase;margin-bottom:8px;}
.hero-title {font-size:clamp(1.9rem,3vw,3.05rem);line-height:1.04;font-weight:760;}
.hero-sub {color:#9F9B90;margin-top:10px;font-size:.96rem;line-height:1.6;max-width:1100px;}
.section-title {font-size:1.22rem;font-weight:730;margin-top:12px;}
.section-note {color:#9F9B90;font-size:.87rem;margin-bottom:13px;line-height:1.55;}
.kpi {border:1px solid rgba(212,175,55,.24);border-radius:17px;padding:15px 16px;min-height:112px;background:linear-gradient(145deg,#121212,#0D0D0D);}
.kpi-label {color:#9F9B90;font-size:.72rem;letter-spacing:.04em;text-transform:uppercase;font-weight:650;}
.kpi-value {font-size:1.55rem;font-weight:760;margin-top:7px;}
.gold {color:#D4AF37;}
.kpi-foot {color:#9F9B90;font-size:.71rem;margin-top:7px;line-height:1.35;}
.info,.callout {background:#0D0D0D;border:1px solid rgba(212,175,55,.24);border-left:3px solid #D4AF37;border-radius:12px;padding:13px 15px;color:#B9B5A9;font-size:.87rem;margin:8px 0 16px;line-height:1.65;}
.callout {background:linear-gradient(145deg,rgba(212,175,55,.08),rgba(255,255,255,.01));border-left:1px solid rgba(212,175,55,.24);}
[data-testid="stDataFrame"] {border:1px solid rgba(212,175,55,.20);border-radius:14px;overflow:hidden;}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# CONSTANTS
# ============================================================

REQUIRED = {
    "Emp Code",
    "Employee Name",
    "FY 26 TGT EQ NS",
    "Equity NS Ach YTD June",
}

FILTER_COLS = ["Status", "Type", "ZONE", "REGION"]

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


# ============================================================
# BASIC HELPERS
# ============================================================

def norm(value):
    if value is None:
        return ""

    if isinstance(value, float) and np.isnan(value):
        return ""

    return " ".join(str(value).strip().split())


def unique_headers(values):
    seen = {}
    output = []

    for i, value in enumerate(values):

        base = norm(value) or "Unnamed_{}".format(i)

        if base not in seen:
            seen[base] = 0
            output.append(base)

        else:
            seen[base] += 1
            output.append(
                "{}.{}".format(
                    base,
                    seen[base]
                )
            )

    return output


# ============================================================
# EXCEL LOADER
# ============================================================

@st.cache_data(show_spinner=False)
def load_uploaded(file_bytes):

    raw = pd.read_excel(
        BytesIO(file_bytes),
        header=None,
        engine="openpyxl",
    )

    header = None

    for i in range(
        min(
            25,
            len(raw)
        )
    ):

        values = {
            norm(v)
            for v
            in raw.iloc[i].tolist()
        }

        if (
            "FY 26 TGT EQ NS" in values
            and
            "Equity NS Ach YTD June" in values
        ):

            header = i
            break


    if header is None:

        raise ValueError(
            "Could not locate the RM header row."
        )


    df = raw.iloc[
        header + 1:
    ].copy()


    df.columns = unique_headers(
        raw.iloc[
            header
        ].tolist()
    )


    df = df.dropna(
        how="all"
    ).copy()


    missing = [
        c
        for c
        in REQUIRED
        if c not in df.columns
    ]


    if missing:

        raise ValueError(
            "Missing columns: {}".format(
                ", ".join(missing)
            )
        )


    for c in [
        "FY 26 TGT EQ NS",
        "Equity NS Ach YTD June",
    ]:

        df[c] = pd.to_numeric(
            df[c],
            errors="coerce",
        )


    df[
        "Employee Name"
    ] = (
        df[
            "Employee Name"
        ]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    df[
        "Emp Code"
    ] = (
        df[
            "Emp Code"
        ]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
    )


    for c in FILTER_COLS:

        if c in df.columns:

            df[c] = (
                df[c]
                .fillna("Unknown")
                .astype(str)
                .str.strip()
            )

            df.loc[
                df[c].eq(""),
                c
            ] = "Unknown"


    # ========================================================
    # USE SECOND MKT TYPE COLUMN
    # ========================================================

    if "MKT TYPE.1" in df.columns:

        market_source = (
            "MKT TYPE.1"
        )

    elif "MKT TYPE" in df.columns:

        market_source = (
            "MKT TYPE"
        )

    else:

        market_source = None


    if market_source:

        df[
            "Market Type"
        ] = (
            df[
                market_source
            ]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
        )

        df.loc[
            df[
                "Market Type"
            ].eq(""),
            "Market Type"
        ] = "Unknown"

    else:

        df[
            "Market Type"
        ] = "Unknown"


    return (
        df,
        header + 1,
        market_source,
    )


# ============================================================
# FORMATTING
# ============================================================

def fmt(value):

    if (
        value is None
        or
        pd.isna(value)
    ):
        return "—"


    value = float(value)


    if abs(value) >= 1_000_000:

        return "{:,.2f}M".format(
            value
            /
            1_000_000
        )


    if abs(value) >= 1_000:

        return "{:,.2f}K".format(
            value
            /
            1_000
        )


    return "{:,.2f}".format(
        value
    )


def pct(value):

    if (
        value is None
        or
        pd.isna(value)
    ):

        return "—"


    return "{:,.1f}%".format(
        float(value)
    )


# ============================================================
# UI HELPERS
# ============================================================

def kpi(
    label,
    value,
    foot="",
    accent=False,
):

    value_class = (
        "kpi-value gold"
        if accent
        else
        "kpi-value"
    )


    st.html(
        """
<div class="kpi">

    <div class="kpi-label">
        {}
    </div>

    <div class="{}">
        {}
    </div>

    <div class="kpi-foot">
        {}
    </div>

</div>
""".format(
            label,
            value_class,
            value,
            foot,
        )
    )


def section(
    title,
    note="",
):

    st.html(
        """
<div class="section-title">
    {}
</div>

<div class="section-note">
    {}
</div>
""".format(
            title,
            note,
        )
    )


def showdf(
    dataframe,
    height=None,
):

    display = (
        dataframe.copy()
    )


    numeric_cols = (
        display
        .select_dtypes(
            include=[
                np.number
            ]
        )
        .columns
    )


    display[
        numeric_cols
    ] = (
        display[
            numeric_cols
        ]
        .round(2)
    )


    kwargs = {

        "data":
            display,

        "width":
            "stretch",

        "hide_index":
            True,

    }


    if height is not None:

        kwargs[
            "height"
        ] = height


    st.dataframe(
        **kwargs
    )


def style(
    fig,
    height=400,
):

    fig.update_layout(

        template=None,

        height=height,

        paper_bgcolor=
            "rgba(0,0,0,0)",

        plot_bgcolor=
            CARD,

        font=dict(
            color=TEXT
        ),

        margin=dict(
            l=28,
            r=24,
            t=58,
            b=48,
        ),

        hoverlabel=dict(
            bgcolor="#171717",
            font_color=TEXT,
            bordercolor=GOLD2,
        ),

        legend=dict(
            bgcolor=
                "rgba(0,0,0,0)"
        ),

    )


    fig.update_xaxes(

        showgrid=True,

        gridcolor=
            GRID,

        zeroline=False,

        linecolor=
            BORDER,

        tickfont=dict(
            color=MUTED
        ),

        title_font=dict(
            color=MUTED
        ),

    )


    fig.update_yaxes(

        showgrid=True,

        gridcolor=
            GRID,

        zeroline=False,

        linecolor=
            BORDER,

        tickfont=dict(
            color=MUTED
        ),

        title_font=dict(
            color=MUTED
        ),

    )


    return fig


# ============================================================
# CORE SCENARIO ENGINE
# ============================================================

def scenario(
    dataframe,
    months,
    uplift=0.0,
    threshold=100.0,
):

    df = (
        dataframe.copy()
    )


    # ========================================================
    # APR + MAY + JUN = 3 ACTUAL MONTHS
    # ========================================================

    df[
        "Current Monthly RR"
    ] = (
        df[
            "Equity NS Ach YTD June"
        ]
        /
        3.0
    )


    # ========================================================
    # ADJUSTED RUN RATE
    # ========================================================

    df[
        "Scenario Monthly RR"
    ] = (

        df[
            "Current Monthly RR"
        ]

        *

        (
            1.0
            +
            uplift
            /
            100.0
        )

    )


    # ========================================================
    # FUTURE NS
    # ========================================================

    df[
        "Projected Future NS"
    ] = (

        df[
            "Scenario Monthly RR"
        ]

        *

        months

    )


    # ========================================================
    # FINAL PROJECTED NS
    # ========================================================

    df[
        "Projected Final NS"
    ] = (

        df[
            "Equity NS Ach YTD June"
        ]

        +

        df[
            "Projected Future NS"
        ]

    )


    # ========================================================
    # PROJECTED ACHIEVEMENT
    # ========================================================

    df[
        "Projected Achievement %"
    ] = np.where(

        df[
            "FY 26 TGT EQ NS"
        ]
        > 0,

        (

            df[
                "Projected Final NS"
            ]

            /

            df[
                "FY 26 TGT EQ NS"
            ]

            *

            100.0

        ),

        np.nan,

    )


    df[
        "Crosses Threshold"
    ] = (

        df[
            "Projected Achievement %"
        ]

        >=

        threshold

    )


    qualifying_total = (

        df.loc[

            df[
                "Crosses Threshold"
            ],

            "Projected Final NS",

        ]

        .sum()

    )


    df[
        "Contribution %"
    ] = 0.0


    if qualifying_total != 0:

        mask = (
            df[
                "Crosses Threshold"
            ]
        )


        df.loc[
            mask,
            "Contribution %",
        ] = (

            df.loc[
                mask,
                "Projected Final NS",
            ]

            /

            qualifying_total

            *

            100.0

        )


    return df


# ============================================================
# BELL CURVE ANALYSIS
# ============================================================

def bell_curve_analysis(
    scenario_df
):

    df = (
        scenario_df.copy()
    )


    values = (

        df[
            "Projected Achievement %"
        ]

        .replace(
            [
                np.inf,
                -np.inf
            ],
            np.nan
        )

        .dropna()

    )


    if values.empty:

        return (
            pd.DataFrame(),
            np.nan,
            np.nan,
            np.nan,
        )


    mean_value = (
        values.mean()
    )


    median_value = (
        values.median()
    )


    std_value = (
        values.std()
    )


    if (
        pd.isna(
            std_value
        )
        or
        std_value == 0
    ):

        std_value = (
            0.000001
        )


    edges = [

        -np.inf,

        mean_value
        -
        2
        *
        std_value,

        mean_value
        -
        std_value,

        mean_value,

        mean_value
        +
        std_value,

        mean_value
        +
        2
        *
        std_value,

        np.inf,

    ]


    labels = [

        "Below μ - 2σ",

        "μ - 2σ to μ - 1σ",

        "μ - 1σ to Mean",

        "Mean to μ + 1σ",

        "μ + 1σ to μ + 2σ",

        "Above μ + 2σ",

    ]


    df[
        "Bell Curve Band"
    ] = pd.cut(

        df[
            "Projected Achievement %"
        ],

        bins=
            edges,

        labels=
            labels,

        include_lowest=
            True,

    )


    normal_reference = {

        "Below μ - 2σ":
            2.3,

        "μ - 2σ to μ - 1σ":
            13.6,

        "μ - 1σ to Mean":
            34.1,

        "Mean to μ + 1σ":
            34.1,

        "μ + 1σ to μ + 2σ":
            13.6,

        "Above μ + 2σ":
            2.3,

    }


    result = (

        df

        .groupby(
            "Bell Curve Band",
            observed=False,
        )

        .agg(

            **{

                "No. of RMs":
                    (
                        "Employee Name",
                        "size"
                    ),

                "Eq Target":
                    (
                        "FY 26 TGT EQ NS",
                        "sum"
                    ),

                "Projected NS":
                    (
                        "Projected Final NS",
                        "sum"
                    ),

                "Avg Achievement %":
                    (
                        "Projected Achievement %",
                        "mean"
                    ),

                "Median Achievement %":
                    (
                        "Projected Achievement %",
                        "median"
                    ),

                "Currently Qualifying":
                    (
                        "Crosses Threshold",
                        "sum"
                    ),

            }

        )

        .reset_index()

    )


    result[
        "Bell Curve Band"
    ] = (
        result[
            "Bell Curve Band"
        ]
        .astype(str)
    )


    total_rms = (
        len(df)
    )


    total_ns = (
        df[
            "Projected Final NS"
        ]
        .sum()
    )


    result[
        "Actual RM Distribution %"
    ] = np.where(

        total_rms > 0,

        result[
            "No. of RMs"
        ]

        /

        total_rms

        *

        100,

        0,

    )


    result[
        "Normal Curve Reference %"
    ] = (

        result[
            "Bell Curve Band"
        ]

        .map(
            normal_reference
        )

        .fillna(
            0
        )

    )


    result[
        "Projected NS Contribution %"
    ] = np.where(

        total_ns != 0,

        result[
            "Projected NS"
        ]

        /

        total_ns

        *

        100,

        0,

    )


    total_row = pd.DataFrame(

        [

            {

                "Bell Curve Band":
                    "Total",

                "No. of RMs":
                    result[
                        "No. of RMs"
                    ].sum(),

                "Eq Target":
                    result[
                        "Eq Target"
                    ].sum(),

                "Projected NS":
                    result[
                        "Projected NS"
                    ].sum(),

                "Avg Achievement %":
                    df[
                        "Projected Achievement %"
                    ].mean(),

                "Median Achievement %":
                    df[
                        "Projected Achievement %"
                    ].median(),

                "Currently Qualifying":
                    int(
                        df[
                            "Crosses Threshold"
                        ].sum()
                    ),

                "Actual RM Distribution %":
                    100.0,

                "Normal Curve Reference %":
                    100.0,

                "Projected NS Contribution %":
                    100.0,

            }

        ]

    )


    return (

        pd.concat(

            [
                result,
                total_row,
            ],

            ignore_index=True,

        ),

        mean_value,

        median_value,

        std_value,

    )


# ============================================================
# MARKET TYPE CONTRIBUTION
# ============================================================

def market_type_contribution_analysis(
    scenario_df
):

    df = (
        scenario_df.copy()
    )


    if (
        "Market Type"
        not in df.columns
    ):

        return (
            pd.DataFrame()
        )


    total_ns = (
        df[
            "Projected Final NS"
        ]
        .sum()
    )


    total_target = (
        df[
            "FY 26 TGT EQ NS"
        ]
        .sum()
    )


    total_current = (
        df[
            "Equity NS Ach YTD June"
        ]
        .sum()
    )


    total_qualifying_ns = (

        df.loc[

            df[
                "Crosses Threshold"
            ],

            "Projected Final NS",

        ]

        .sum()

    )


    result = (

        df

        .groupby(
            "Market Type",
            dropna=False,
        )

        .agg(

            **{

                "#RMs":
                    (
                        "Employee Name",
                        "size"
                    ),

                "Eq Target":
                    (
                        "FY 26 TGT EQ NS",
                        "sum"
                    ),

                "Current Achievement":
                    (
                        "Equity NS Ach YTD June",
                        "sum"
                    ),

                "Projected NS":
                    (
                        "Projected Final NS",
                        "sum"
                    ),

                "Currently Qualifying":
                    (
                        "Crosses Threshold",
                        "sum"
                    ),

                "Average Achievement %":
                    (
                        "Projected Achievement %",
                        "mean"
                    ),

            }

        )

        .reset_index()

    )


    qualifying = (

        df.loc[
            df[
                "Crosses Threshold"
            ]
        ]

        .groupby(
            "Market Type"
        )[
            "Projected Final NS"
        ]

        .sum()

        .rename(
            "Qualifying Projected NS"
        )

        .reset_index()

    )


    result = result.merge(

        qualifying,

        on=
            "Market Type",

        how=
            "left",

    )


    result[
        "Qualifying Projected NS"
    ] = (
        result[
            "Qualifying Projected NS"
        ]
        .fillna(0)
    )


    result[
        "Qualification Rate %"
    ] = np.where(

        result[
            "#RMs"
        ]
        > 0,

        result[
            "Currently Qualifying"
        ]

        /

        result[
            "#RMs"
        ]

        *

        100,

        0,

    )


    result[
        "Target Contribution %"
    ] = np.where(

        total_target
        !=
        0,

        result[
            "Eq Target"
        ]

        /

        total_target

        *

        100,

        0,

    )


    result[
        "Current Achievement Contribution %"
    ] = np.where(

        total_current
        !=
        0,

        result[
            "Current Achievement"
        ]

        /

        total_current

        *

        100,

        0,

    )


    result[
        "Projected NS Contribution %"
    ] = np.where(

        total_ns
        !=
        0,

        result[
            "Projected NS"
        ]

        /

        total_ns

        *

        100,

        0,

    )


    result[
        "Qualifying NS Contribution %"
    ] = np.where(

        total_qualifying_ns
        !=
        0,

        result[
            "Qualifying Projected NS"
        ]

        /

        total_qualifying_ns

        *

        100,

        0,

    )


    result = result.sort_values(

        "Projected NS Contribution %",

        ascending=False,

    )


    total_row = pd.DataFrame(

        [

            {

                "Market Type":
                    "Total",

                "#RMs":
                    result[
                        "#RMs"
                    ].sum(),

                "Eq Target":
                    result[
                        "Eq Target"
                    ].sum(),

                "Current Achievement":
                    result[
                        "Current Achievement"
                    ].sum(),

                "Projected NS":
                    result[
                        "Projected NS"
                    ].sum(),

                "Currently Qualifying":
                    result[
                        "Currently Qualifying"
                    ].sum(),

                "Average Achievement %":
                    df[
                        "Projected Achievement %"
                    ].mean(),

                "Qualifying Projected NS":
                    result[
                        "Qualifying Projected NS"
                    ].sum(),

                "Qualification Rate %":
                    df[
                        "Crosses Threshold"
                    ].mean()
                    *
                    100,

                "Target Contribution %":
                    100.0,

                "Current Achievement Contribution %":
                    100.0,

                "Projected NS Contribution %":
                    100.0,

                "Qualifying NS Contribution %":
                    100.0,

            }

        ]

    )


    return pd.concat(

        [
            result,
            total_row,
        ],

        ignore_index=True,

    )


# ============================================================
# INDIVIDUAL CONTRIBUTION BINS
# ============================================================

def contribution_bins(
    dataframe
):

    q = (
        dataframe[
            dataframe[
                "Crosses Threshold"
            ]
        ]
        .copy()
    )


    if q.empty:

        return (
            pd.DataFrame()
        )


    q[
        "Contribution Bin"
    ] = pd.cut(

        q[
            "Contribution %"
        ],

        bins=[

            -np.inf,

            0.10,

            0.25,

            0.50,

            1,

            2,

            5,

            10,

            np.inf,

        ],

        labels=[

            "≤0.10%",

            "0.10–0.25%",

            "0.25–0.50%",

            "0.50–1.00%",

            "1.00–2.00%",

            "2.00–5.00%",

            "5.00–10.00%",

            ">10.00%",

        ],

        include_lowest=True,

    )


    result = (

        q

        .groupby(
            "Contribution Bin",
            observed=False,
        )

        .agg(

            **{

                "No. of RMs":
                    (
                        "Employee Name",
                        "size"
                    ),

                "Projected NS":
                    (
                        "Projected Final NS",
                        "sum"
                    ),

                "Contribution %":
                    (
                        "Contribution %",
                        "sum"
                    ),

            }

        )

        .reset_index()

    )


    total = pd.DataFrame(

        [

            {

                "Contribution Bin":
                    "Total",

                "No. of RMs":
                    result[
                        "No. of RMs"
                    ].sum(),

                "Projected NS":
                    result[
                        "Projected NS"
                    ].sum(),

                "Contribution %":
                    result[
                        "Contribution %"
                    ].sum(),

            }

        ]

    )


    return pd.concat(

        [
            result,
            total,
        ],

        ignore_index=True,

    )


# ============================================================
# PAGE 1 SCENARIO PANEL
# ============================================================

def scenario_panel(
    df,
    name,
    threshold,
    months,
    uplift,
    market,
):

    qualifying = (
        df[
            df[
                "Crosses Threshold"
            ]
        ]
        .copy()
    )


    non_qualifying = (
        df[
            ~df[
                "Crosses Threshold"
            ]
        ]
        .copy()
    )


    hit_rate = (

        len(
            qualifying
        )

        /

        len(
            df
        )

        *

        100

        if len(
            df
        )

        else
        0

    )


    section(

        name,

        "{} · through {} · uplift {:.1f}% · threshold {:.1f}%".format(

            market,

            MONTHS[
                months
            ],

            uplift,

            threshold,

        ),

    )


    cols = st.columns(
        6
    )


    cards = [

        (
            "RMs Analysed",
            len(
                df
            ),
            market,
            False,
        ),

        (
            "RMs Crossing",
            len(
                qualifying
            ),
            "{:.1f}% of RMs".format(
                hit_rate
            ),
            True,
        ),

        (
            "Qualifying RM NS",
            fmt(
                qualifying[
                    "Projected Final NS"
                ]
                .sum()
            ),
            "Crossing threshold",
            True,
        ),

        (
            "Non-Qualifying NS",
            fmt(
                non_qualifying[
                    "Projected Final NS"
                ]
                .sum()
            ),
            "Below threshold",
            False,
        ),

        (
            "All RM Projected NS",
            fmt(
                df[
                    "Projected Final NS"
                ]
                .sum()
            ),
            "Target "
            +
            fmt(
                df[
                    "FY 26 TGT EQ NS"
                ]
                .sum()
            ),
            False,
        ),

        (
            "Median Achievement",
            pct(
                df[
                    "Projected Achievement %"
                ]
                .median()
            ),
            "Projected FY achievement",
            False,
        ),

    ]


    for i, card in enumerate(
        cards
    ):

        with cols[i]:

            kpi(
                *card
            )


    st.html(
        """
<div class="info">

<b style="color:#F3F0E7">
Calculation
</b>

<br>

Monthly RR =
Apr-Jun achieved ÷ 3.

<br>

Projected Final NS =
Apr-Jun actual +
scenario monthly RR × selected future months.

<br>

Achievement % =
projected final NS ÷ FY target × 100.

<br>

Contribution % =
qualifying RM projected NS ÷
total qualifying projected NS.

</div>
"""
    )


    # ========================================================
    # BELL CURVE TABLE
    # ========================================================

    section(

        "Bell Curve Analysis",

        (
            "Actual RM achievement is grouped around "
            "the population mean and standard deviation. "
            "The normal-curve percentage is only a reference benchmark."
        ),

    )


    (
        bell_table,
        bell_mean,
        bell_median,
        bell_std,
    ) = bell_curve_analysis(
        df
    )


    bell_cols = st.columns(
        4
    )


    with bell_cols[0]:

        kpi(
            "Mean Achievement",
            pct(
                bell_mean
            ),
            "Average projected achievement",
        )


    with bell_cols[1]:

        kpi(
            "Median Achievement",
            pct(
                bell_median
            ),
            "50th percentile RM",
        )


    with bell_cols[2]:

        kpi(
            "Standard Deviation",
            pct(
                bell_std
            ),
            "Achievement spread",
        )


    with bell_cols[3]:

        above = int(

            (
                df[
                    "Projected Achievement %"
                ]

                >=
                bell_mean

            )

            .sum()

        )


        share = (

            above

            /

            len(
                df
            )

            *

            100

            if len(
                df
            )

            else
            0

        )


        kpi(
            "RMs Above Mean",
            above,
            "{:.1f}% of RMs".format(
                share
            ),
            True,
        )


    if not bell_table.empty:

        showdf(
            bell_table
        )


    # ========================================================
    # MARKET TYPE
    # ========================================================

    section(

        "Market Type Contribution Analysis",

        (
            "Contribution of T2, T6, B30, EM, T30 "
            "and other market types to target, projected NS "
            "and qualifying NS."
        ),

    )


    market_table = (
        market_type_contribution_analysis(
            df
        )
    )


    if market_table.empty:

        st.info(
            "Market Type data is unavailable."
        )

    else:

        showdf(
            market_table
        )


    # ========================================================
    # CONTRIBUTION BINS
    # ========================================================

    section(

        "Individual RM Contribution Bins",

        (
            "Among qualifying RMs only, this shows "
            "the concentration of contribution to "
            "qualifying projected NS."
        ),

    )


    bins = (
        contribution_bins(
            df
        )
    )


    if bins.empty:

        st.info(
            "No qualifying RMs."
        )

    else:

        showdf(
            bins
        )


    # ========================================================
    # RM DETAIL
    # ========================================================

    section(

        "Qualifying RM Detail",

        "RM-level output for RMs crossing the scenario threshold.",

    )


    if qualifying.empty:

        st.info(
            "No qualifying RMs."
        )

    else:

        display_cols = [

            "Emp Code",

            "Employee Name",

            "Market Type",

        ]


        display_cols += [

            c

            for c
            in FILTER_COLS

            if c
            in qualifying.columns

        ]


        display_cols += [

            "FY 26 TGT EQ NS",

            "Equity NS Ach YTD June",

            "Current Monthly RR",

            "Scenario Monthly RR",

            "Projected Final NS",

            "Projected Achievement %",

            "Contribution %",

        ]


        showdf(

            qualifying[
                display_cols
            ]

            .sort_values(
                "Contribution %",
                ascending=False,
            ),

            520,

        )


# ============================================================
# BEFORE VS NOW BELL CURVE
# ============================================================

def before_after_bell_curve(
    before_df,
    after_df,
):

    before = (

        before_df[
            "Projected Achievement %"
        ]

        .replace(
            [
                np.inf,
                -np.inf
            ],
            np.nan
        )

        .dropna()

        .astype(float)

    )


    after = (

        after_df[
            "Projected Achievement %"
        ]

        .replace(
            [
                np.inf,
                -np.inf
            ],
            np.nan
        )

        .dropna()

        .astype(float)

    )


    fig = go.Figure()


    if (
        before.empty
        or
        after.empty
    ):

        return style(
            fig,
            470
        )


    b_mean = (
        before.mean()
    )

    a_mean = (
        after.mean()
    )


    b_std = (
        before.std()
    )

    a_std = (
        after.std()
    )


    b_std = (

        0.000001

        if (
            pd.isna(
                b_std
            )
            or
            b_std == 0
        )

        else

        float(
            b_std
        )

    )


    a_std = (

        0.000001

        if (
            pd.isna(
                a_std
            )
            or
            a_std == 0
        )

        else

        float(
            a_std
        )

    )


    min_x = min(

        before.min(),

        after.min(),

        b_mean
        -
        3
        *
        b_std,

        a_mean
        -
        3
        *
        a_std,

    )


    max_x = max(

        before.max(),

        after.max(),

        b_mean
        +
        3
        *
        b_std,

        a_mean
        +
        3
        *
        a_std,

    )


    x = np.linspace(
        min_x,
        max_x,
        500,
    )


    b_y = (

        1

        /

        (
            b_std
            *
            np.sqrt(
                2
                *
                np.pi
            )
        )

        *

        np.exp(

            -0.5

            *

            (
                (
                    x
                    -
                    b_mean
                )

                /

                b_std
            )

            ** 2

        )

    )


    a_y = (

        1

        /

        (
            a_std
            *
            np.sqrt(
                2
                *
                np.pi
            )
        )

        *

        np.exp(

            -0.5

            *

            (
                (
                    x
                    -
                    a_mean
                )

                /

                a_std
            )

            ** 2

        )

    )


    fig.add_trace(

        go.Scatter(

            x=x,

            y=b_y,

            mode=
                "lines",

            name=
                "Before · Current Run Rate",

            line=dict(
                color="#7D7D7D",
                width=3,
            ),

            fill=
                "tozeroy",

            fillcolor=
                "rgba(125,125,125,0.10)",

        )

    )


    fig.add_trace(

        go.Scatter(

            x=x,

            y=a_y,

            mode=
                "lines",

            name=
                "Now · Increased Run Rate",

            line=dict(
                color=GOLD,
                width=3,
            ),

            fill=
                "tozeroy",

            fillcolor=
                "rgba(212,175,55,0.12)",

        )

    )


    fig.add_vline(

        x=
            b_mean,

        line_dash=
            "dash",

        line_color=
            "#999999",

        annotation_text=
            "Before Mean {:.1f}%".format(
                b_mean
            ),

    )


    fig.add_vline(

        x=
            a_mean,

        line_dash=
            "dash",

        line_color=
            GOLD,

        annotation_text=
            "Now Mean {:.1f}%".format(
                a_mean
            ),

    )


    fig.add_vline(

        x=100,

        line_dash=
            "dot",

        line_color=
            TEXT,

        annotation_text=
            "100% Target",

    )


    fig.update_layout(

        title=
            "Bell Curve Shift · Before vs Now",

        xaxis_title=
            "Projected Achievement (%)",

        yaxis_title=
            "Distribution Density",

        hovermode=
            "x unified",

    )


    return style(
        fig,
        470
    )


def before_after_distribution_summary(
    before_df,
    after_df,
):

    before = (

        before_df[
            "Projected Achievement %"
        ]

        .replace(
            [
                np.inf,
                -np.inf
            ],
            np.nan
        )

        .dropna()

    )


    after = (

        after_df[
            "Projected Achievement %"
        ]

        .replace(
            [
                np.inf,
                -np.inf
            ],
            np.nan
        )

        .dropna()

    )


    return pd.DataFrame(

        {

            "Metric": [

                "Mean Achievement %",

                "Median Achievement %",

                "Standard Deviation",

                "RMs ≥ 100%",

                "RMs ≥ 120%",

            ],


            "Before": [

                before.mean(),

                before.median(),

                before.std(),

                int(
                    (
                        before
                        >=
                        100
                    )
                    .sum()
                ),

                int(
                    (
                        before
                        >=
                        120
                    )
                    .sum()
                ),

            ],


            "Now": [

                after.mean(),

                after.median(),

                after.std(),

                int(
                    (
                        after
                        >=
                        100
                    )
                    .sum()
                ),

                int(
                    (
                        after
                        >=
                        120
                    )
                    .sum()
                ),

            ],


            "Change": [

                after.mean()
                -
                before.mean(),

                after.median()
                -
                before.median(),

                after.std()
                -
                before.std(),

                int(
                    (
                        after
                        >=
                        100
                    )
                    .sum()
                )
                -
                int(
                    (
                        before
                        >=
                        100
                    )
                    .sum()
                ),

                int(
                    (
                        after
                        >=
                        120
                    )
                    .sum()
                )
                -
                int(
                    (
                        before
                        >=
                        120
                    )
                    .sum()
                ),

            ],

        }

    )


# ============================================================
# TARGET BUCKETS
# ============================================================

def bucket_option_1(
    base_df,
    step,
    top,
):

    target = (
        base_df[
            "FY 26 TGT EQ NS"
        ]
    )


    rows = []

    lower = 0.0

    upper = float(
        step
    )


    while upper < top:

        mask = (

            (
                target
                >
                lower
            )

            &

            (
                target
                <=
                upper
            )

        )


        rows.append(

            {

                "Current Target Bucketing":
                    "Upto {:g}".format(
                        upper
                    ),

                "#RMs":
                    int(
                        mask.sum()
                    ),

                "Eq Target":
                    target[
                        mask
                    ].sum(),

            }

        )


        lower = (
            upper
        )


        upper += (
            step
        )


    mask = (

        (
            target
            >
            lower
        )

        &

        (
            target
            <=
            top
        )

    )


    rows.append(

        {

            "Current Target Bucketing":
                "Upto {:g}".format(
                    top
                ),

            "#RMs":
                int(
                    mask.sum()
                ),

            "Eq Target":
                target[
                    mask
                ].sum(),

        }

    )


    mask = (
        target
        >
        top
    )


    rows.append(

        {

            "Current Target Bucketing":
                "Above {:g}".format(
                    top
                ),

            "#RMs":
                int(
                    mask.sum()
                ),

            "Eq Target":
                target[
                    mask
                ].sum(),

        }

    )


    rows.append(

        {

            "Current Target Bucketing":
                "Total",

            "#RMs":
                len(
                    base_df
                ),

            "Eq Target":
                target.sum(),

        }

    )


    return pd.DataFrame(
        rows
    )


def bucket_option_2(
    base_df,
    current_scenario,
    cutoff,
):

    target = (
        base_df[
            "FY 26 TGT EQ NS"
        ]
    )


    achieved = (
        base_df[
            "Equity NS Ach YTD June"
        ]
    )


    qualifies = (

        current_scenario[
            "Projected Achievement %"
        ]

        >=
        100

    )


    rows = []


    for label, mask in [

        (
            "Upto {:g} crores".format(
                cutoff
            ),
            target
            <=
            cutoff,
        ),

        (
            "{:g} and above".format(
                cutoff
            ),
            target
            >
            cutoff,
        ),

    ]:


        rows.append(

            {

                "Current Target":
                    label,

                "#RMs":
                    int(
                        mask.sum()
                    ),

                "Eq Target":
                    target[
                        mask
                    ].sum(),

                "Current Achievement":
                    achieved[
                        mask
                    ].sum(),

                "Currently Qualifying":
                    int(
                        (
                            qualifies
                            &
                            mask
                        )
                        .sum()
                    ),

            }

        )


    rows.append(

        {

            "Current Target":
                "Total",

            "#RMs":
                len(
                    base_df
                ),

            "Eq Target":
                target.sum(),

            "Current Achievement":
                achieved.sum(),

            "Currently Qualifying":
                int(
                    qualifies.sum()
                ),

        }

    )


    return pd.DataFrame(
        rows
    )


def bucket_option_3(
    base_df,
    cutoff_1,
    cutoff_2,
):

    target = (
        base_df[
            "FY 26 TGT EQ NS"
        ]
    )


    rows = []


    specs = [

        (
            "Upto {:g} crores".format(
                cutoff_1
            ),
            target
            <=
            cutoff_1,
        ),

        (
            "{:g} to {:g} crores".format(
                cutoff_1,
                cutoff_2
            ),
            (
                target
                >
                cutoff_1
            )
            &
            (
                target
                <=
                cutoff_2
            ),
        ),

        (
            "Above {:g} crores".format(
                cutoff_2
            ),
            target
            >
            cutoff_2,
        ),

    ]


    for label, mask in specs:

        rows.append(

            {

                "Current Target":
                    label,

                "#RMs":
                    int(
                        mask.sum()
                    ),

                "Eq Target":
                    target[
                        mask
                    ].sum(),

            }

        )


    rows.append(

        {

            "Current Target":
                "Total",

            "#RMs":
                len(
                    base_df
                ),

            "Eq Target":
                target.sum(),

        }

    )


    return pd.DataFrame(
        rows
    )


# ============================================================
# TRAVEL BUDGET
# ============================================================

def qualification_budget_section(
    base_df,
    current_scenario,
    uplift_scenario,
    months,
):

    section(

        "Qualification & Travel Budget Planner",

        (
            "Convert qualification count into an editable "
            "travel budget and compare different current-target cohort cuts."
        ),

    )


    qualified = int(

        (
            current_scenario[
                "Projected Achievement %"
            ]

            >=
            100

        )

        .sum()

    )


    incremental_ns = (

        uplift_scenario[
            "Projected Final NS"
        ]
        .sum()

        -

        current_scenario[
            "Projected Final NS"
        ]
        .sum()

    )


    c1, c2, c3, c4 = st.columns(
        4
    )


    with c1:

        travelers = st.number_input(

            "Number of People for Trip",

            min_value=0,

            max_value=
                max(
                    len(
                        base_df
                    ),
                    1
                ),

            value=
                min(
                    40,
                    len(
                        base_df
                    )
                ),

            step=1,

            key=
                "budget_people",

        )


    with c2:

        cost_lakh = st.number_input(

            "Trip Cost per Person (₹ lakh)",

            min_value=0.0,

            max_value=100.0,

            value=3.0,

            step=0.25,

            key=
                "budget_cost",

        )


    budget_cr = (

        travelers

        *

        cost_lakh

        /

        100.0

    )


    budget_pct = (

        budget_cr

        /

        incremental_ns

        *

        100

        if incremental_ns
        >
        0

        else

        np.nan

    )


    with c3:

        kpi(

            "Trip Budget",

            "₹{:,.2f} Cr".format(
                budget_cr
            ),

            "{} people × ₹{:.2f} lakh".format(
                travelers,
                cost_lakh
            ),

            True,

        )


    with c4:

        kpi(

            "Budget / Incremental NS",

            pct(
                budget_pct
            ),

            "Against Scenario 2 incremental NS",

        )


    st.html(
        """
<div class="callout">

<b style="color:#D4AF37">
Current qualification:
</b>

{} RMs are projected to cross 100%
through {}.

<br><br>

Example:

<b style="color:#F3F0E7">
40 people × ₹3 lakh = ₹1.20 Cr.
</b>

</div>
""".format(
            qualified,
            MONTHS[
                months
            ],
        )
    )


    section(

        "Current Target Bucketing Options",

        (
            "Editable ways to split RMs "
            "by their current Equity NS target."
        ),

    )


    col1, col2, col3 = st.columns(
        3
    )


    with col1:

        st.markdown(
            "#### Option 1 · Detailed Bands"
        )


        step = st.number_input(

            "Bucket Step (Cr)",

            1.0,

            25.0,

            5.0,

            1.0,

            key=
                "b1_step",

        )


        top = st.number_input(

            "Last Cutoff (Cr)",

            min_value=
                step,

            max_value=
                200.0,

            value=
                50.0,

            step=
                5.0,

            key=
                "b1_top",

        )


        option_1 = (
            bucket_option_1(
                base_df,
                step,
                top,
            )
        )


        showdf(
            option_1
        )


    with col2:

        st.markdown(
            "#### Option 2 · Two Cohorts"
        )


        cutoff = st.number_input(

            "Split Cutoff (Cr)",

            0.5,

            100.0,

            6.5,

            0.5,

            key=
                "b2_cut",

        )


        option_2 = (
            bucket_option_2(
                base_df,
                current_scenario,
                cutoff,
            )
        )


        showdf(
            option_2
        )


    with col3:

        st.markdown(
            "#### Option 3 · Three Cohorts"
        )


        cut1 = st.number_input(

            "First Cutoff (Cr)",

            0.5,

            100.0,

            5.0,

            0.5,

            key=
                "b3_cut1",

        )


        cut2 = st.number_input(

            "Second Cutoff (Cr)",

            min_value=
                cut1
                +
                0.5,

            max_value=
                200.0,

            value=
                max(
                    10.0,
                    cut1 + 0.5
                ),

            step=
                0.5,

            key=
                "b3_cut2",

        )


        option_3 = (
            bucket_option_3(
                base_df,
                cut1,
                cut2,
            )
        )


        showdf(
            option_3
        )


    section(

        "Editable Cohort Planning",

        (
            "Edit stretch % and number "
            "of travelers for each target cohort."
        ),

    )


    plan = (
        option_3
        .iloc[:-1]
        .copy()
    )


    plan[
        "Proposed Stretch %"
    ] = [

        35.0,

        25.0,

        15.0,

    ][
        :len(
            plan
        )
    ]


    plan[
        "Planned Travelers"
    ] = [

        min(
            10,
            int(v)
        )

        for v
        in plan[
            "#RMs"
        ]

    ]


    edited = st.data_editor(

        plan,

        width=
            "stretch",

        hide_index=
            True,

        key=
            "cohort_editor",

        column_config={

            "Proposed Stretch %":

                st.column_config.NumberColumn(

                    "Proposed Stretch %",

                    min_value=0.0,

                    max_value=200.0,

                    step=1.0,

                ),

            "Planned Travelers":

                st.column_config.NumberColumn(

                    "Planned Travelers",

                    min_value=0,

                    step=1,

                ),

        },

        disabled=[

            "Current Target",

            "#RMs",

            "Eq Target",

        ],

    )


    calc = (
        edited.copy()
    )


    calc[
        "Implied Extra Target"
    ] = (

        calc[
            "Eq Target"
        ]

        *

        calc[
            "Proposed Stretch %"
        ]

        /

        100.0

    )


    calc[
        "Trip Budget (Cr)"
    ] = (

        calc[
            "Planned Travelers"
        ]

        *

        cost_lakh

        /

        100.0

    )


    calc[
        "Extra Target / Trip Budget (x)"
    ] = np.where(

        calc[
            "Trip Budget (Cr)"
        ]
        > 0,

        calc[
            "Implied Extra Target"
        ]

        /

        calc[
            "Trip Budget (Cr)"
        ],

        np.nan,

    )


    showdf(
        calc
    )


# ============================================================
# PAGE 2 - NEW INSIGHTS
# ============================================================

def summary_by(
    df,
    dimension,
):

    if dimension not in df.columns:

        return (
            pd.DataFrame()
        )


    result = (

        df

        .groupby(
            dimension,
            dropna=False,
        )

        .agg(

            RMs=(
                "Employee Name",
                "size"
            ),

            Target=(
                "FY 26 TGT EQ NS",
                "sum"
            ),

            YTD=(
                "Equity NS Ach YTD June",
                "sum"
            ),

            Projected_NS=(
                "Projected Final NS",
                "sum"
            ),

            Median_Ach=(
                "Projected Achievement %",
                "median"
            ),

            Hit_Rate=(
                "Crosses Threshold",
                "mean"
            ),

        )

        .reset_index()

    )


    result[
        "Hit_Rate"
    ] = (

        result[
            "Hit_Rate"
        ]

        *

        100

    )


    return result.sort_values(

        "Projected_NS",

        ascending=False,

    )


def new_insights(
    base,
    months,
    market,
):

    df = scenario(
        base,
        months,
        0,
        100,
    )


    st.html(
        """
<div class="hero">

    <div class="eyebrow">
        Page 2 · New Insights
    </div>

    <div class="hero-title">
        Where is the Sales Opportunity Actually Sitting?
    </div>

    <div class="hero-sub">
        Management cuts across market type,
        geography, quick wins and high-value risk.
    </div>

</div>
"""
    )


    cols = st.columns(
        5
    )


    cards = [

        (
            "FY Target",
            fmt(
                df[
                    "FY 26 TGT EQ NS"
                ]
                .sum()
            ),
            market,
            False,
        ),

        (
            "Apr-Jun Achievement",
            fmt(
                df[
                    "Equity NS Ach YTD June"
                ]
                .sum()
            ),
            "3-month actual",
            False,
        ),

        (
            "Projected NS",
            fmt(
                df[
                    "Projected Final NS"
                ]
                .sum()
            ),
            MONTHS[
                months
            ],
            True,
        ),

        (
            "RMs ≥100%",
            int(
                (
                    df[
                        "Projected Achievement %"
                    ]
                    >=
                    100
                )
                .sum()
            ),
            pct(
                (
                    df[
                        "Projected Achievement %"
                    ]
                    >=
                    100
                )
                .mean()
                *
                100
            ),
            True,
        ),

        (
            "Median Achievement",
            pct(
                df[
                    "Projected Achievement %"
                ]
                .median()
            ),
            "Current-RR projection",
            False,
        ),

    ]


    for i, card in enumerate(
        cards
    ):

        with cols[i]:

            kpi(
                *card
            )


    for dim, title in [

        (
            "Market Type",
            "Market Type"
        ),

        (
            "ZONE",
            "Zone"
        ),

        (
            "REGION",
            "Region"
        ),

    ]:


        if dim in df.columns:

            section(

                "{} Cut".format(
                    title
                ),

                (
                    "Target, YTD, projection "
                    "and hit-rate by {}."
                    .format(
                        title.lower()
                    )
                ),

            )


            showdf(
                summary_by(
                    df,
                    dim,
                )
            )


    # ========================================================
    # TARGET VS PROJECTED
    # ========================================================

    section(

        "Target vs Projected NS",

        (
            "Dots above the diagonal are projected "
            "to beat the existing target."
        ),

    )


    fig = go.Figure(

        go.Scatter(

            x=
                df[
                    "FY 26 TGT EQ NS"
                ],

            y=
                df[
                    "Projected Final NS"
                ],

            mode=
                "markers",

            marker=dict(

                size=8,

                color=
                    df[
                        "Projected Achievement %"
                    ],

                colorscale=
                    "Cividis",

                showscale=
                    True,

            ),

            text=
                df[
                    "Employee Name"
                ],

            hovertemplate=(

                "<b>%{text}</b>"

                "<br>Target %{x:,.2f}"

                "<br>Projected %{y:,.2f}"

                "<extra></extra>"

            ),

        )

    )


    mx = max(

        df[
            "FY 26 TGT EQ NS"
        ]
        .max(),

        df[
            "Projected Final NS"
        ]
        .max(),

        1,

    )


    fig.add_scatter(

        x=[
            0,
            mx
        ],

        y=[
            0,
            mx
        ],

        mode=
            "lines",

        line=dict(
            color=GOLD2,
            dash="dash",
        ),

        name=
            "100% Line",

    )


    fig.update_layout(

        xaxis_title=
            "FY Target",

        yaxis_title=
            "Projected Final NS",

    )


    st.plotly_chart(

        style(
            fig,
            430
        ),

        config={
            "displayModeBar":
                False
        },

    )


    # ========================================================
    # REQUIRED UPLIFT
    # ========================================================

    df[
        "Required Future RR for 100%"
    ] = (

        df[
            "FY 26 TGT EQ NS"
        ]

        -

        df[
            "Equity NS Ach YTD June"
        ]

    ) / months


    df[
        "Required RR Uplift %"
    ] = np.where(

        df[
            "Current Monthly RR"
        ]
        > 0,

        (

            df[
                "Required Future RR for 100%"
            ]

            /

            df[
                "Current Monthly RR"
            ]

            -

            1

        )

        *

        100,

        np.nan,

    )


    # ========================================================
    # QUICK WINS
    # ========================================================

    quick = (

        df[

            (
                df[
                    "Projected Achievement %"
                ]
                >=
                80
            )

            &

            (
                df[
                    "Projected Achievement %"
                ]
                <
                100
            )

            &

            df[
                "Required RR Uplift %"
            ]
            .between(
                0,
                30
            )

        ]

        .copy()

    )


    section(

        "Quick Wins",

        (
            "RMs projected at 80–100% who need "
            "≤30% future run-rate uplift to reach target."
        ),

    )


    if quick.empty:

        st.info(
            "No quick wins under current filters."
        )

    else:

        cols2 = [

            "Emp Code",

            "Employee Name",

            "Market Type",

        ]


        cols2 += [

            c

            for c in [

                "ZONE",

                "REGION",

            ]

            if c
            in quick.columns

        ]


        cols2 += [

            "FY 26 TGT EQ NS",

            "Projected Final NS",

            "Projected Achievement %",

            "Required RR Uplift %",

        ]


        showdf(

            quick[
                cols2
            ]

            .sort_values(
                "FY 26 TGT EQ NS",
                ascending=False,
            ),

            420,

        )


    # ========================================================
    # HIGH VALUE RISK
    # ========================================================

    risk = (

        df[
            df[
                "Projected Achievement %"
            ]
            <
            80
        ]

        .copy()

    )


    risk[
        "Target at Risk"
    ] = (

        risk[
            "FY 26 TGT EQ NS"
        ]

        -

        risk[
            "Projected Final NS"
        ]

    ).clip(
        lower=0
    )


    section(

        "High-Value Target at Risk",

        (
            "Prioritize absolute sales value "
            "at risk rather than percentage miss alone."
        ),

    )


    if risk.empty:

        st.success(
            "No RMs below 80%."
        )

    else:

        cols3 = [

            "Emp Code",

            "Employee Name",

            "Market Type",

        ]


        cols3 += [

            c

            for c in [

                "ZONE",

                "REGION",

            ]

            if c
            in risk.columns

        ]


        cols3 += [

            "FY 26 TGT EQ NS",

            "Projected Final NS",

            "Projected Achievement %",

            "Target at Risk",

        ]


        showdf(

            risk

            .nlargest(
                30,
                "Target at Risk"
            )[
                cols3
            ],

            420,

        )


# ============================================================
# PAGE 3 - BONVOYAGE
# ============================================================

def bonvoyage_model(
    base,
    months,
    min_uplift,
    max_uplift,
    allocation,
    trip_lakh,
    max_feasible,
):

    df = scenario(
        base,
        months,
        0,
        100,
    )


    df[
        "No-Incentive Expected NS"
    ] = np.maximum(

        df[
            "FY 26 TGT EQ NS"
        ],

        df[
            "Projected Final NS"
        ],

    )


    df[
        "Target Size Percentile"
    ] = (

        df[
            "FY 26 TGT EQ NS"
        ]

        .rank(
            pct=True,
            method="average",
        )

    )


    df[
        "Peer Performance Percentile"
    ] = (

        df

        .groupby(
            "Market Type"
        )[
            "Projected Achievement %"
        ]

        .rank(
            pct=True,
            method="average",
        )

        .fillna(
            0.5
        )

    )


    momentum = (

        df[
            "Projected Achievement %"
        ]

        /

        120

    ).clip(
        0,
        1
    ).fillna(
        0
    )


    df[
        "Capacity Score"
    ] = (

        0.45

        *

        (
            1
            -
            df[
                "Target Size Percentile"
            ]
        )

        +

        0.35

        *

        df[
            "Peer Performance Percentile"
        ]

        +

        0.20

        *

        momentum

    ).clip(
        0,
        1
    )


    df[
        "Planned Future RR Uplift %"
    ] = (

        min_uplift

        +

        (
            max_uplift
            -
            min_uplift
        )

        *

        df[
            "Capacity Score"
        ]

    )


    df[
        "Capacity-Based Target"
    ] = (

        df[
            "Equity NS Ach YTD June"
        ]

        +

        df[
            "Current Monthly RR"
        ]

        *

        (
            1

            +

            df[
                "Planned Future RR Uplift %"
            ]

            /

            100
        )

        *

        months

    )


    df[
        "BonVoyage Stretch Target"
    ] = np.maximum(

        df[
            "No-Incentive Expected NS"
        ],

        df[
            "Capacity-Based Target"
        ],

    )


    df[
        "Incremental NS Required"
    ] = (

        df[
            "BonVoyage Stretch Target"
        ]

        -

        df[
            "No-Incentive Expected NS"
        ]

    ).clip(
        lower=0
    )


    df[
        "Required Future RR for BonVoyage"
    ] = (

        df[
            "BonVoyage Stretch Target"
        ]

        -

        df[
            "Equity NS Ach YTD June"
        ]

    ) / months


    df[
        "Required Future RR Uplift %"
    ] = np.where(

        df[
            "Current Monthly RR"
        ]
        > 0,

        (

            df[
                "Required Future RR for BonVoyage"
            ]

            /

            df[
                "Current Monthly RR"
            ]

            -

            1

        )

        *

        100,

        np.nan,

    )


    trip_cost_cr = (
        trip_lakh
        /
        100.0
    )


    df[
        "Trip Budget Ceiling"
    ] = (

        df[
            "Incremental NS Required"
        ]

        *

        allocation

        /

        100.0

    )


    df[
        "Trip Funding Coverage x"
    ] = np.where(

        trip_cost_cr
        > 0,

        df[
            "Trip Budget Ceiling"
        ]

        /

        trip_cost_cr,

        np.inf,

    )


    df[
        "Recommended Candidate"
    ] = (

        (
            df[
                "Trip Funding Coverage x"
            ]
            >=
            1
        )

        &

        df[
            "Required Future RR Uplift %"
        ]
        .between(
            0,
            max_feasible
        )

        &

        (
            df[
                "Incremental NS Required"
            ]
            >
            0
        )

        &

        (
            df[
                "Current Monthly RR"
            ]
            >
            0
        )

    )


    return df


def bonvoyage_page(
    base,
    months,
    market,
    min_uplift,
    max_uplift,
    allocation,
    trip_lakh,
    max_feasible,
):

    df = bonvoyage_model(

        base,

        months,

        min_uplift,

        max_uplift,

        allocation,

        trip_lakh,

        max_feasible,

    )


    candidates = (

        df[
            df[
                "Recommended Candidate"
            ]
        ]

    )


    st.html(
        """
<div class="hero">

    <div class="eyebrow">
        Page 3 · BonVoyage
    </div>

    <div class="hero-title">
        Fund the Foreign Trip with Genuinely Incremental Sales
    </div>

    <div class="hero-sub">
        Sales the RM was already expected to generate
        are not counted as BonVoyage incremental sales.
    </div>

</div>
"""
    )


    cols = st.columns(
        5
    )


    cards = [

        (
            "Official Target",

            fmt(
                df[
                    "FY 26 TGT EQ NS"
                ]
                .sum()
            ),

            market,

            False,

        ),

        (
            "No-Incentive Expected",

            fmt(
                df[
                    "No-Incentive Expected NS"
                ]
                .sum()
            ),

            "What we expect anyway",

            False,

        ),

        (
            "BonVoyage Target",

            fmt(
                df[
                    "BonVoyage Stretch Target"
                ]
                .sum()
            ),

            "Personalized stretch",

            True,

        ),

        (
            "Incremental NS",

            fmt(
                df[
                    "Incremental NS Required"
                ]
                .sum()
            ),

            "Above no-incentive baseline",

            True,

        ),

        (
            "Recommended Candidates",

            len(
                candidates
            ),

            "≤{}% required RR uplift".format(
                max_feasible
            ),

            False,

        ),

    ]


    for i, card in enumerate(
        cards
    ):

        with cols[i]:

            kpi(
                *card
            )


    section(

        "BonVoyage Target vs No-Incentive Baseline",

        (
            "Distance above the diagonal represents "
            "the incremental target being asked from the RM."
        ),

    )


    fig = go.Figure(

        go.Scatter(

            x=
                df[
                    "No-Incentive Expected NS"
                ],

            y=
                df[
                    "BonVoyage Stretch Target"
                ],

            mode=
                "markers",

            marker=dict(

                size=np.clip(

                    7

                    +

                    df[
                        "Incremental NS Required"
                    ]

                    .clip(
                        lower=0
                    )

                    *

                    2,

                    7,

                    28,

                ),

                color=
                    df[
                        "Required Future RR Uplift %"
                    ],

                colorscale=
                    "Cividis",

                showscale=
                    True,

                colorbar=dict(
                    title=
                        "RR uplift %"
                ),

            ),

            text=
                df[
                    "Employee Name"
                ],

            hovertemplate=(

                "<b>%{text}</b>"

                "<br>No-incentive %{x:,.2f}"

                "<br>BonVoyage %{y:,.2f}"

                "<extra></extra>"

            ),

        )

    )


    mx = max(

        df[
            "BonVoyage Stretch Target"
        ]
        .max(),

        df[
            "No-Incentive Expected NS"
        ]
        .max(),

        1,

    )


    fig.add_scatter(

        x=[
            0,
            mx
        ],

        y=[
            0,
            mx
        ],

        mode=
            "lines",

        line=dict(
            color=GOLD2,
            dash="dash",
        ),

        name=
            "No stretch",

    )


    fig.update_layout(

        xaxis_title=
            "No-Incentive Expected NS",

        yaxis_title=
            "BonVoyage Stretch Target",

    )


    st.plotly_chart(

        style(
            fig,
            460
        ),

        config={
            "displayModeBar":
                False
        },

    )


    section(

        "RM-Level BonVoyage Target Book",

        (
            "Personalized target, incremental ask, "
            "trip-funding coverage and eligibility."
        ),

    )


    display_cols = [

        "Emp Code",

        "Employee Name",

        "Market Type",

    ]


    display_cols += [

        c

        for c in [

            "ZONE",

            "REGION",

        ]

        if c
        in df.columns

    ]


    display_cols += [

        "FY 26 TGT EQ NS",

        "Projected Final NS",

        "No-Incentive Expected NS",

        "Capacity Score",

        "Planned Future RR Uplift %",

        "BonVoyage Stretch Target",

        "Incremental NS Required",

        "Required Future RR Uplift %",

        "Trip Budget Ceiling",

        "Trip Funding Coverage x",

        "Recommended Candidate",

    ]


    showdf(

        df[
            display_cols
        ]

        .sort_values(
            "Incremental NS Required",
            ascending=False,
        ),

        620,

    )


    for dim in [

        "Market Type",

        "ZONE",

        "REGION",

    ]:


        if dim not in df.columns:

            continue


        section(

            "{} BonVoyage Economics".format(
                dim
            ),

            "Incremental opportunity and candidate count by {}.".format(
                dim
            ),

        )


        summary = (

            df

            .groupby(
                dim
            )

            .agg(

                RMs=(
                    "Employee Name",
                    "size"
                ),

                Original_Target=(
                    "FY 26 TGT EQ NS",
                    "sum"
                ),

                BonVoyage_Target=(
                    "BonVoyage Stretch Target",
                    "sum"
                ),

                Incremental_NS=(
                    "Incremental NS Required",
                    "sum"
                ),

                Trip_Budget=(
                    "Trip Budget Ceiling",
                    "sum"
                ),

                Candidates=(
                    "Recommended Candidate",
                    "sum"
                ),

            )

            .reset_index()

            .sort_values(
                "Incremental_NS",
                ascending=False,
            )

        )


        showdf(
            summary
        )


# ============================================================
# PAGE 4 - RUN RATE ANALYSIS
# ============================================================

def run_rate_summary_row(
    scenario_df,
    label,
    uplift_pct,
    baseline_df,
):

    total_rms = (
        len(
            scenario_df
        )
    )


    total_target = (
        scenario_df[
            "FY 26 TGT EQ NS"
        ]
        .sum()
    )


    total_ns = (
        scenario_df[
            "Projected Final NS"
        ]
        .sum()
    )


    qmask = (

        scenario_df[
            "Projected Achievement %"
        ]

        >=
        100

    )


    baseline_qmask = (

        baseline_df[
            "Projected Achievement %"
        ]

        >=
        100

    )


    q_count = int(
        qmask.sum()
    )


    q_ns = (

        scenario_df.loc[

            qmask,

            "Projected Final NS",

        ]

        .sum()

    )


    non_q_ns = (
        total_ns
        -
        q_ns
    )


    return {

        "Scenario":
            label,

        "Run Rate Uplift %":
            uplift_pct,

        "Total RMs":
            total_rms,

        "Total Target":
            total_target,

        "Total Projected NS":
            total_ns,

        "Portfolio Achievement %":

            (
                total_ns
                /
                total_target
                *
                100

                if total_target

                else 0
            ),

        "RMs Crossing 100%":
            q_count,

        "Qualification Rate %":

            (
                q_count
                /
                total_rms
                *
                100

                if total_rms

                else 0
            ),

        "Qualifying RM NS":
            q_ns,

        "Qualifying NS Contribution %":

            (
                q_ns
                /
                total_ns
                *
                100

                if total_ns

                else 0
            ),

        "Non-Qualifying RM NS":
            non_q_ns,

        "Non-Qualifying NS Contribution %":

            (
                non_q_ns
                /
                total_ns
                *
                100

                if total_ns

                else 0
            ),

        "Newly Crossing 100% vs Current":

            int(

                (
                    qmask
                    &
                    ~baseline_qmask
                )

                .sum()

            ),

        "Incremental Total NS vs Current":

            total_ns

            -

            baseline_df[
                "Projected Final NS"
            ]
            .sum(),

    }


# ============================================================
# PAGE 4 - MARKET TYPE TABLE
# ============================================================

def run_rate_market_type_table(
    scenario_df,
    scenario_label,
):

    if (
        "Market Type"
        not in scenario_df.columns
    ):

        return (
            pd.DataFrame()
        )


    df = (
        scenario_df.copy()
    )


    all_ns = (
        df[
            "Projected Final NS"
        ]
        .sum()
    )


    qmask = (

        df[
            "Projected Achievement %"
        ]

        >=
        100

    )


    grouped = (

        df

        .groupby(
            "Market Type",
            dropna=False,
        )

        .agg(

            **{

                "#RMs":
                    (
                        "Employee Name",
                        "size"
                    ),

                "Eq Target":
                    (
                        "FY 26 TGT EQ NS",
                        "sum"
                    ),

                "Total Projected NS":
                    (
                        "Projected Final NS",
                        "sum"
                    ),

            }

        )

        .reset_index()

    )


    q_count = (

        df

        .assign(
            _qual=qmask
        )

        .groupby(
            "Market Type"
        )[
            "_qual"
        ]

        .sum()

        .rename(
            "RMs Crossing 100%"
        )

        .reset_index()

    )


    q_ns = (

        df.loc[
            qmask
        ]

        .groupby(
            "Market Type"
        )[
            "Projected Final NS"
        ]

        .sum()

        .rename(
            "Qualifying RM NS"
        )

        .reset_index()

    )


    grouped = grouped.merge(

        q_count,

        on=
            "Market Type",

        how=
            "left",

    )


    grouped = grouped.merge(

        q_ns,

        on=
            "Market Type",

        how=
            "left",

    )


    grouped[
        "RMs Crossing 100%"
    ] = (

        grouped[
            "RMs Crossing 100%"
        ]

        .fillna(
            0
        )

    )


    grouped[
        "Qualifying RM NS"
    ] = (

        grouped[
            "Qualifying RM NS"
        ]

        .fillna(
            0
        )

    )


    grouped[
        "Qualification Rate %"
    ] = np.where(

        grouped[
            "#RMs"
        ]
        > 0,

        grouped[
            "RMs Crossing 100%"
        ]

        /

        grouped[
            "#RMs"
        ]

        *

        100,

        0,

    )


    grouped[
        "Market Type NS Contribution %"
    ] = np.where(

        all_ns
        !=
        0,

        grouped[
            "Total Projected NS"
        ]

        /

        all_ns

        *

        100,

        0,

    )


    grouped[
        "Qualifying Amount Contribution to Total NS"
    ] = (

        grouped[
            "Qualifying RM NS"
        ]

    )


    grouped[
        "Qualifying % Contribution to Total NS"
    ] = np.where(

        all_ns
        !=
        0,

        grouped[
            "Qualifying RM NS"
        ]

        /

        all_ns

        *

        100,

        0,

    )


    grouped.insert(

        0,

        "Scenario",

        scenario_label,

    )


    total_row = pd.DataFrame(

        [

            {

                "Scenario":
                    scenario_label,

                "Market Type":
                    "Total",

                "#RMs":
                    grouped[
                        "#RMs"
                    ].sum(),

                "Eq Target":
                    grouped[
                        "Eq Target"
                    ].sum(),

                "Total Projected NS":
                    grouped[
                        "Total Projected NS"
                    ].sum(),

                "RMs Crossing 100%":
                    grouped[
                        "RMs Crossing 100%"
                    ].sum(),

                "Qualifying RM NS":
                    grouped[
                        "Qualifying RM NS"
                    ].sum(),

                "Qualification Rate %":

                    (
                        grouped[
                            "RMs Crossing 100%"
                        ]
                        .sum()

                        /

                        grouped[
                            "#RMs"
                        ]
                        .sum()

                        *

                        100

                        if grouped[
                            "#RMs"
                        ]
                        .sum()

                        else 0
                    ),

                "Market Type NS Contribution %":
                    100.0,

                "Qualifying Amount Contribution to Total NS":

                    grouped[
                        "Qualifying RM NS"
                    ]
                    .sum(),

                "Qualifying % Contribution to Total NS":

                    (
                        grouped[
                            "Qualifying RM NS"
                        ]
                        .sum()

                        /

                        all_ns

                        *

                        100

                        if all_ns

                        else 0
                    ),

            }

        ]

    )


    return pd.concat(

        [

            grouped

            .sort_values(

                "Qualifying % Contribution to Total NS",

                ascending=False,

            ),

            total_row,

        ],

        ignore_index=True,

    )


# ============================================================
# PAGE 4
# ============================================================

def run_rate_analysis_page(
    base,
    months,
    market,
    custom_uplift,
):

    # ========================================================
    # BUILD ALL SCENARIOS
    # ========================================================

    baseline = scenario(
        base,
        months,
        0,
        100,
    )


    scenario_5 = scenario(
        base,
        months,
        5,
        100,
    )


    scenario_10 = scenario(
        base,
        months,
        10,
        100,
    )


    scenario_15 = scenario(
        base,
        months,
        15,
        100,
    )


    scenario_custom = scenario(
        base,
        months,
        custom_uplift,
        100,
    )


    scenario_map = {

        "Current Run Rate":
            (
                baseline,
                0.0
            ),

        "+5% Run Rate":
            (
                scenario_5,
                5.0
            ),

        "+10% Run Rate":
            (
                scenario_10,
                10.0
            ),

        "+15% Run Rate":
            (
                scenario_15,
                15.0
            ),

        "Custom +{:.1f}%".format(
            custom_uplift
        ):
            (
                scenario_custom,
                float(
                    custom_uplift
                )
            ),

    }


    # ========================================================
    # HEADER
    # ========================================================

    st.html(
        """
<div class="hero">

    <div class="eyebrow">
        Page 4 · Run Rate Qualification Analysis
    </div>

    <div class="hero-title">
        How much value do the 100%+ RMs create at each run-rate level?
    </div>

    <div class="hero-sub">
        Compare Current, +5%, +10%, +15% and Custom run-rate scenarios.
        See total target, total projected NS, how many RMs cross 100%,
        how much NS those RMs generate, and their amount and percentage
        contribution to total NS.
    </div>

</div>
"""
    )


    st.html(
        """
<div class="callout">

<b style="color:#D4AF37">
Important definition
</b>

<br><br>

On this page,
<b style="color:#F3F0E7">
achieving target
</b>
and
<b style="color:#F3F0E7">
crossing 100%
</b>
mean projected achievement ≥ 100%.

<br>

The page separately shows
<b style="color:#F3F0E7">
Newly Crossing 100%
</b>,
meaning RMs who were below 100% at the current run rate
but cross it after the uplift.

</div>
"""
    )


    # ========================================================
    # MASTER COMPARISON
    # ========================================================

    rows = [

        run_rate_summary_row(
            df,
            label,
            uplift,
            baseline,
        )

        for label, (
            df,
            uplift
        )

        in scenario_map.items()

    ]


    comparison = pd.DataFrame(
        rows
    )


    section(

        "Executive Comparison",

        (
            "Main comparison of total target, projected NS, "
            "100%+ RM count, qualifying NS amount, contribution "
            "percentage and incremental improvement."
        ),

    )


    showdf(
        comparison
    )


    # ========================================================
    # CURRENT VS CUSTOM KPI
    # ========================================================

    current = (
        comparison.iloc[
            0
        ]
    )


    custom = (
        comparison.iloc[
            -1
        ]
    )


    cols = st.columns(
        6
    )


    cards = [

        (
            "Total Target",

            fmt(
                current[
                    "Total Target"
                ]
            ),

            market,

            False,

        ),

        (
            "Current Projected NS",

            fmt(
                current[
                    "Total Projected NS"
                ]
            ),

            pct(
                current[
                    "Portfolio Achievement %"
                ]
            ),

            False,

        ),

        (
            "Current RMs ≥100%",

            int(
                current[
                    "RMs Crossing 100%"
                ]
            ),

            pct(
                current[
                    "Qualification Rate %"
                ]
            ),

            True,

        ),

        (
            "Current Qualifying RM NS",

            fmt(
                current[
                    "Qualifying RM NS"
                ]
            ),

            pct(
                current[
                    "Qualifying NS Contribution %"
                ]
            )
            +
            " of total NS",

            True,

        ),

        (
            "Custom RMs ≥100%",

            int(
                custom[
                    "RMs Crossing 100%"
                ]
            ),

            "+{} newly crossing".format(

                int(
                    custom[
                        "Newly Crossing 100% vs Current"
                    ]
                )

            ),

            True,

        ),

        (
            "Custom Qualifying RM NS",

            fmt(
                custom[
                    "Qualifying RM NS"
                ]
            ),

            pct(
                custom[
                    "Qualifying NS Contribution %"
                ]
            )
            +
            " of total NS",

            True,

        ),

    ]


    for i, card in enumerate(
        cards
    ):

        with cols[i]:

            kpi(
                *card
            )


    # ========================================================
    # COMPARISON CHARTS
    # ========================================================

    left, right = st.columns(

        2,

        gap=
            "large",

    )


    with left:

        fig = go.Figure()


        fig.add_bar(

            x=
                comparison[
                    "Scenario"
                ],

            y=
                comparison[
                    "Total Projected NS"
                ],

            name=
                "Total Projected NS",

            marker_color=
                "#595959",

        )


        fig.add_bar(

            x=
                comparison[
                    "Scenario"
                ],

            y=
                comparison[
                    "Qualifying RM NS"
                ],

            name=
                "NS from RMs ≥100%",

            marker_color=
                GOLD,

        )


        fig.update_layout(

            title=
                "Total Net Sales vs Net Sales from 100%+ RMs",

            barmode=
                "group",

            xaxis_title=
                "Run Rate Scenario",

            yaxis_title=
                "Projected Net Sales",

        )


        st.plotly_chart(

            style(
                fig,
                430
            ),

            config={
                "displayModeBar":
                    False
            },

        )


    with right:

        fig = go.Figure()


        fig.add_bar(

            x=
                comparison[
                    "Scenario"
                ],

            y=
                comparison[
                    "RMs Crossing 100%"
                ],

            name=
                "RMs ≥100%",

            marker_color=
                GOLD,

        )


        fig.add_scatter(

            x=
                comparison[
                    "Scenario"
                ],

            y=
                comparison[
                    "Qualifying NS Contribution %"
                ],

            mode=
                "lines+markers",

            name=
                "Qualifying NS Contribution %",

            yaxis=
                "y2",

            line=dict(
                color=TEXT,
                width=2,
            ),

        )


        fig.update_layout(

            title=
                "Qualification Count and Contribution to Total NS",

            xaxis_title=
                "Run Rate Scenario",

            yaxis_title=
                "RMs Crossing 100%",

            yaxis2=dict(

                title=
                    "Contribution to Total NS (%)",

                overlaying=
                    "y",

                side=
                    "right",

                gridcolor=
                    "rgba(0,0,0,0)",

                tickfont=dict(
                    color=MUTED
                ),

                title_font=dict(
                    color=MUTED
                ),

            ),

        )


        st.plotly_chart(

            style(
                fig,
                430
            ),

            config={
                "displayModeBar":
                    False
            },

        )


    # ========================================================
    # SCENARIO DETAIL
    # ========================================================

    section(

        "Scenario-by-Scenario Detail",

        (
            "Each tab shows how much the 100%+ RMs contribute "
            "in amount and percentage terms, plus Market Type "
            "and RM-level detail."
        ),

    )


    tabs = st.tabs(
        list(
            scenario_map.keys()
        )
    )


    baseline_qmask = (

        baseline[
            "Projected Achievement %"
        ]

        >=
        100

    )


    for tab, (
        label,
        (
            df,
            uplift
        )
    ) in zip(

        tabs,

        scenario_map.items(),

    ):


        with tab:

            total_target = (
                df[
                    "FY 26 TGT EQ NS"
                ]
                .sum()
            )


            total_ns = (
                df[
                    "Projected Final NS"
                ]
                .sum()
            )


            qmask = (

                df[
                    "Projected Achievement %"
                ]

                >=
                100

            )


            q_count = int(
                qmask.sum()
            )


            q_ns = (

                df.loc[

                    qmask,

                    "Projected Final NS",

                ]

                .sum()

            )


            q_share = (

                q_ns
                /
                total_ns
                *
                100

                if total_ns

                else 0

            )


            q_rate = (

                q_count
                /
                len(
                    df
                )
                *
                100

                if len(
                    df
                )

                else 0

            )


            newly = int(

                (
                    qmask
                    &
                    ~baseline_qmask
                )

                .sum()

            )


            cols2 = st.columns(
                6
            )


            cards2 = [

                (
                    "Total Target",

                    fmt(
                        total_target
                    ),

                    "Same target base",

                    False,

                ),

                (
                    "Total Projected NS",

                    fmt(
                        total_ns
                    ),

                    pct(

                        total_ns
                        /
                        total_target
                        *
                        100

                        if total_target

                        else 0

                    ),

                    False,

                ),

                (
                    "RMs Crossing 100%",

                    q_count,

                    pct(
                        q_rate
                    ),

                    True,

                ),

                (
                    "Qualifying RM NS",

                    fmt(
                        q_ns
                    ),

                    "Amount contribution",

                    True,

                ),

                (
                    "Qualifying NS Contribution",

                    pct(
                        q_share
                    ),

                    "Share of total projected NS",

                    True,

                ),

                (
                    "Newly Crossing vs Current",

                    newly,

                    "Run-rate uplift {:.1f}%".format(
                        uplift
                    ),

                    False,

                ),

            ]


            for i, card in enumerate(
                cards2
            ):

                with cols2[i]:

                    kpi(
                        *card
                    )


            # =================================================
            # MARKET TYPE
            # =================================================

            section(

                "Market Type Contribution",

                (
                    "For each MKT TYPE: RM count, target, "
                    "total projected NS, number crossing 100%, "
                    "qualifying NS amount and qualifying contribution "
                    "to total NS."
                ),

            )


            mt = run_rate_market_type_table(

                df,

                label,

            )


            if mt.empty:

                st.info(
                    "Market Type data is unavailable."
                )

            else:

                showdf(
                    mt
                )


            # =================================================
            # RM DETAIL
            # =================================================

            section(

                "Qualifying RM Detail",

                (
                    "Every RM at or above 100%, "
                    "sorted by amount contribution "
                    "to total projected NS."
                ),

            )


            detail = (

                df.loc[
                    qmask
                ]

                .copy()

            )


            if detail.empty:

                st.info(
                    "No RMs cross 100% in this scenario."
                )

            else:

                detail[
                    "Amount Contribution to Total NS"
                ] = (

                    detail[
                        "Projected Final NS"
                    ]

                )


                detail[
                    "% Contribution to Total NS"
                ] = np.where(

                    total_ns
                    !=
                    0,

                    detail[
                        "Projected Final NS"
                    ]

                    /

                    total_ns

                    *

                    100,

                    0,

                )


                detail_cols = [

                    "Emp Code",

                    "Employee Name",

                    "Market Type",

                ]


                detail_cols += [

                    c

                    for c in [

                        "ZONE",

                        "REGION",

                    ]

                    if c
                    in detail.columns

                ]


                detail_cols += [

                    "FY 26 TGT EQ NS",

                    "Projected Final NS",

                    "Projected Achievement %",

                    "Amount Contribution to Total NS",

                    "% Contribution to Total NS",

                ]


                showdf(

                    detail[
                        detail_cols
                    ]

                    .sort_values(

                        "Amount Contribution to Total NS",

                        ascending=False,

                    ),

                    520,

                )


    # ========================================================
    # UPLIFT IMPACT
    # ========================================================

    section(

        "What Do We Gain by Increasing Run Rate?",

        (
            "Incremental 100%+ RMs and incremental "
            "total projected NS compared with the "
            "current-run-rate baseline."
        ),

    )


    impact_cols = [

        "Scenario",

        "Run Rate Uplift %",

        "RMs Crossing 100%",

        "Newly Crossing 100% vs Current",

        "Total Projected NS",

        "Incremental Total NS vs Current",

        "Qualifying RM NS",

        "Qualifying NS Contribution %",

    ]


    showdf(

        comparison[
            impact_cols
        ]

    )


# ============================================================
# MAIN APP HEADER
# ============================================================

st.html(
    """
<div class="hero">

    <div class="eyebrow">
        RM Equity Net Sales · Strategy Lab
    </div>

    <div class="hero-title">
        Scenario Lab, New Insights, BonVoyage
        and Run-Rate Qualification Analysis
    </div>

    <div class="hero-sub">
        Upload the RM workbook once.
        The app reads only the uploaded workbook.
    </div>

</div>
"""
)


# ============================================================
# SIDEBAR NAVIGATION + UPLOAD
# ============================================================

with st.sidebar:

    st.markdown(
        "### Navigation"
    )


    page = st.radio(

        "Page",

        [

            "1 · Scenario Lab",

            "2 · New Insights",

            "3 · BonVoyage",

            "4 · Run Rate Analysis",

        ],

    )


    st.divider()


    st.markdown(
        "### Upload Data"
    )


    uploaded = st.file_uploader(

        "Upload RM Workbook",

        type=[
            "xlsx"
        ],

        help=
            "The app reads only the workbook uploaded here.",

    )


# ============================================================
# REQUIRE UPLOAD
# ============================================================

if uploaded is None:

    st.info(

        "Upload the RM Excel workbook from the sidebar. "
        "No stored/local workbook is read."

    )

    st.stop()


try:

    (
        raw,
        header_row,
        market_type_source,
    ) = load_uploaded(

        uploaded.getvalue()

    )


except Exception as error:

    st.error(

        "Could not read workbook: {}".format(
            error
        )

    )

    st.stop()


# ============================================================
# VALID DATA
# ============================================================

identity = (

    raw[
        "Employee Name"
    ].ne("")

    |

    raw[
        "Emp Code"
    ].ne("")

)


numeric = (

    raw[
        "FY 26 TGT EQ NS"
    ].notna()

    &

    raw[
        "Equity NS Ach YTD June"
    ].notna()

)


positive_target = (

    raw[
        "FY 26 TGT EQ NS"
    ]

    >
    0

)


excluded_missing = int(

    (
        identity
        &
        ~numeric
    )

    .sum()

)


excluded_target = int(

    (
        identity
        &
        numeric
        &
        ~positive_target
    )

    .sum()

)


data = (

    raw[

        identity

        &

        numeric

        &

        positive_target

    ]

    .copy()

)


# ============================================================
# GLOBAL FILTERS + PAGE CONTROLS
# ============================================================

with st.sidebar:

    st.divider()


    st.markdown(
        "### Global Filters"
    )


    market_values = sorted(

        [

            v

            for v
            in data[
                "Market Type"
            ]

            .dropna()

            .astype(str)

            .unique()

            if v != "Unknown"

        ]

    )


    selected_market_type = st.selectbox(

        "MKT TYPE",

        [
            "All"
        ]

        +

        market_values,

    )


    if selected_market_type != "All":

        data = (

            data[

                data[
                    "Market Type"
                ]

                ==

                selected_market_type

            ]

            .copy()

        )


    # ========================================================
    # OTHER FILTERS
    # ========================================================

    filter_specs = [

        (
            "Status",
            "Status",
            True
        ),

        (
            "Type",
            "Type",
            False
        ),

        (
            "ZONE",
            "Zone",
            False
        ),

        (
            "REGION",
            "Region",
            False
        ),

    ]


    for (
        column,
        label,
        active_default,
    ) in filter_specs:


        if column not in data.columns:

            continue


        values = sorted(

            data[
                column
            ]

            .dropna()

            .astype(str)

            .unique()

            .tolist()

        )


        if (

            active_default

            and

            "Active" in values

        ):

            default = [
                "Active"
            ]

        else:

            default = (
                values
            )


        selected = st.multiselect(

            label,

            values,

            default=
                default,

        )


        data = (

            data[

                data[
                    column
                ]

                .isin(
                    selected
                )

            ]

            .copy()

        )


    # ========================================================
    # PROJECTION MONTH
    # ========================================================

    st.divider()


    months = st.select_slider(

        "Projection Months After June",

        options=
            list(
                MONTHS
            ),

        value=9,

        format_func=

            lambda m:

            "{} months · through {}".format(

                m,

                MONTHS[
                    m
                ],

            ),

    )


    # ========================================================
    # PAGE 1 CONTROLS
    # ========================================================

    uplift = (
        10.0
    )


    threshold = (
        100.0
    )


    if page == "1 · Scenario Lab":


        st.divider()


        st.markdown(
            "### Scenario 2"
        )


        uplift_choice = st.selectbox(

            "Increase Current Run Rate By",

            [

                "5%",

                "10%",

                "15%",

                "Custom",

            ],

            index=1,

        )


        if uplift_choice == "Custom":

            uplift = st.number_input(

                "Custom Run Rate Increase (%)",

                min_value=0.0,

                max_value=500.0,

                value=10.0,

                step=1.0,

            )

        else:

            uplift = float(

                uplift_choice.replace(
                    "%",
                    ""
                )

            )


        threshold = st.number_input(

            "Scenario 2 Achievement Threshold (%)",

            min_value=1.0,

            max_value=500.0,

            value=100.0,

            step=5.0,

        )


    # ========================================================
    # PAGE 3 CONTROLS
    # ========================================================

    min_bv_uplift = (
        15
    )


    max_bv_uplift = (
        60
    )


    allocation = (
        10
    )


    trip_lakh = (
        3.0
    )


    max_feasible = (
        50
    )


    if page == "3 · BonVoyage":


        st.divider()


        st.markdown(
            "### BonVoyage Economics"
        )


        min_bv_uplift = st.slider(

            "Minimum Planned Future RR Uplift (%)",

            0,

            50,

            15,

        )


        max_bv_uplift = st.slider(

            "Maximum Planned Future RR Uplift (%)",

            min_value=
                max(
                    min_bv_uplift,
                    10
                ),

            max_value=
                100,

            value=
                max(
                    60,
                    min_bv_uplift
                ),

        )


        allocation = st.slider(

            "% of Incremental NS for Trip Budget",

            1,

            30,

            10,

        )


        trip_lakh = st.number_input(

            "Assumed Trip Cost per RM (₹ lakh)",

            min_value=0.1,

            max_value=25.0,

            value=3.0,

            step=0.5,

        )


        max_feasible = st.slider(

            "Maximum Feasible Future RR Uplift (%)",

            5,

            100,

            50,

            5,

        )


    # ========================================================
    # PAGE 4 CONTROL
    # ========================================================

    custom_analysis_uplift = (
        20.0
    )


    if page == "4 · Run Rate Analysis":


        st.divider()


        st.markdown(
            "### Run Rate Comparison"
        )


        custom_analysis_uplift = st.number_input(

            "Custom Run Rate Increase (%)",

            min_value=0.0,

            max_value=500.0,

            value=20.0,

            step=1.0,

            help=(

                "Page 4 always compares "
                "Current, +5%, +10%, +15% "
                "and this custom uplift."

            ),

        )


    st.divider()


    st.caption(

        (
            "Uploaded {} | "
            "Header row {} | "
            "MKT source {} | "
            "Excluded missing {} | "
            "Excluded target≤0 {}"
        )

        .format(

            uploaded.name,

            header_row,

            market_type_source,

            excluded_missing,

            excluded_target,

        )

    )


# ============================================================
# EMPTY CHECK
# ============================================================

if data.empty:

    st.warning(
        "No RMs remain after filters."
    )

    st.stop()


market = (

    "All Market Types"

    if selected_market_type
    ==
    "All"

    else

    selected_market_type

)


# ============================================================
# PAGE ROUTING
# ============================================================

if page == "1 · Scenario Lab":


    scenario_1 = scenario(

        data,

        months,

        0,

        100,

    )


    scenario_2 = scenario(

        data,

        months,

        uplift,

        threshold,

    )


    tab1, tab2, tab3 = st.tabs(

        [

            "Scenario 1 · Current Run Rate",

            "Scenario 2 · Increased Run Rate",

            "Scenario Comparison",

        ]

    )


    # ========================================================
    # TAB 1
    # ========================================================

    with tab1:

        scenario_panel(

            scenario_1,

            "Scenario 1 · Current Run Rate",

            100,

            months,

            0,

            market,

        )


    # ========================================================
    # TAB 2
    # ========================================================

    with tab2:

        scenario_panel(

            scenario_2,

            "Scenario 2 · Increased Run Rate",

            threshold,

            months,

            uplift,

            market,

        )


    # ========================================================
    # TAB 3
    # ========================================================

    with tab3:

        section(

            "Scenario Comparison",

            "Current monthly run rate vs {:.1f}% higher run rate."
            .format(
                uplift
            ),

        )


        comparison = pd.DataFrame(

            {

                "Metric": [

                    "RMs Crossing",

                    "Hit-rate %",

                    "Qualifying Projected NS",

                    "All Projected NS",

                    "Median Achievement %",

                ],


                "Scenario 1": [

                    int(

                        (
                            scenario_1[
                                "Projected Achievement %"
                            ]

                            >=
                            100

                        )

                        .sum()

                    ),

                    (
                        scenario_1[
                            "Projected Achievement %"
                        ]

                        >=
                        100

                    )

                    .mean()

                    *

                    100,

                    scenario_1.loc[

                        scenario_1[
                            "Projected Achievement %"
                        ]
                        >=
                        100,

                        "Projected Final NS",

                    ]

                    .sum(),

                    scenario_1[
                        "Projected Final NS"
                    ]

                    .sum(),

                    scenario_1[
                        "Projected Achievement %"
                    ]

                    .median(),

                ],


                "Scenario 2": [

                    int(

                        (
                            scenario_2[
                                "Projected Achievement %"
                            ]

                            >=
                            threshold

                        )

                        .sum()

                    ),

                    (
                        scenario_2[
                            "Projected Achievement %"
                        ]

                        >=
                        threshold

                    )

                    .mean()

                    *

                    100,

                    scenario_2.loc[

                        scenario_2[
                            "Projected Achievement %"
                        ]
                        >=
                        threshold,

                        "Projected Final NS",

                    ]

                    .sum(),

                    scenario_2[
                        "Projected Final NS"
                    ]

                    .sum(),

                    scenario_2[
                        "Projected Achievement %"
                    ]

                    .median(),

                ],

            }

        )


        showdf(
            comparison
        )


    # ========================================================
    # BELL CURVE SHIFT
    # ========================================================

    st.divider()


    section(

        "Bell Curve Shift · Before vs Now",

        (
            "Before is the current-run-rate projection. "
            "Now is the {:.1f}% uplift scenario."
        )
        .format(
            uplift
        ),

    )


    before = (

        scenario_1[
            "Projected Achievement %"
        ]

        .replace(
            [
                np.inf,
                -np.inf
            ],
            np.nan
        )

        .dropna()

    )


    after = (

        scenario_2[
            "Projected Achievement %"
        ]

        .replace(
            [
                np.inf,
                -np.inf
            ],
            np.nan
        )

        .dropna()

    )


    before_mean = (
        before.mean()
    )


    after_mean = (
        after.mean()
    )


    before_std = (
        before.std()
    )


    after_std = (
        after.std()
    )


    before_q = int(

        (
            before
            >=
            100
        )

        .sum()

    )


    after_q = int(

        (
            after
            >=
            100
        )

        .sum()

    )


    cols = st.columns(
        4
    )


    with cols[0]:

        kpi(

            "Before Mean",

            pct(
                before_mean
            ),

            "Current-run-rate distribution",

        )


    with cols[1]:

        kpi(

            "Now Mean",

            pct(
                after_mean
            ),

            "{:+.1f} pp movement".format(

                after_mean
                -
                before_mean

            ),

            True,

        )


    with cols[2]:

        spread_change = (

            after_std

            -

            before_std

        )


        spread_label = (

            "Wider spread"

            if spread_change
            >
            0

            else

            "Narrower spread"

            if spread_change
            <
            0

            else

            "No spread change"

        )


        kpi(

            "Spread Change",

            "{:+.1f} pp".format(
                spread_change
            ),

            "{} · {:.1f} → {:.1f}".format(

                spread_label,

                before_std,

                after_std,

            ),

        )


    with cols[3]:

        kpi(

            "Additional RMs ≥100%",

            "{:+,}".format(

                after_q
                -
                before_q

            ),

            "{} → {} RMs".format(

                before_q,

                after_q,

            ),

            True,

        )


    st.plotly_chart(

        before_after_bell_curve(

            scenario_1,

            scenario_2,

        ),

        config={
            "displayModeBar":
                False
        },

    )


    showdf(

        before_after_distribution_summary(

            scenario_1,

            scenario_2,

        )

    )


    # ========================================================
    # TRAVEL PLANNER
    # ========================================================

    st.divider()


    qualification_budget_section(

        data,

        scenario_1,

        scenario_2,

        months,

    )


# ============================================================
# PAGE 2
# ============================================================

elif page == "2 · New Insights":


    new_insights(

        data,

        months,

        market,

    )


# ============================================================
# PAGE 3
# ============================================================

elif page == "3 · BonVoyage":


    bonvoyage_page(

        data,

        months,

        market,

        min_bv_uplift,

        max_bv_uplift,

        allocation,

        trip_lakh,

        max_feasible,

    )


# ============================================================
# PAGE 4
# ============================================================

else:


    run_rate_analysis_page(

        data,

        months,

        market,

        custom_analysis_uplift,

    )
