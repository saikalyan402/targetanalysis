from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="RM Equity NS Strategy Lab",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# THEME
# ============================================================

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

.stApp {
    background:#070707;
    color:#F3F0E7;
}

[data-testid="stSidebar"] {
    background:#0B0B0B;
    border-right:1px solid rgba(212,175,55,.24);
}

[data-testid="stSidebar"] * {
    color:#F3F0E7;
}

.hero {
    border:1px solid rgba(212,175,55,.24);
    border-radius:22px;
    padding:24px 26px;
    margin-bottom:18px;
    background:linear-gradient(
        110deg,
        rgba(212,175,55,.10),
        rgba(255,255,255,.015)
    );
}

.eyebrow {
    color:#D4AF37;
    font-size:.76rem;
    letter-spacing:.17em;
    font-weight:750;
    text-transform:uppercase;
    margin-bottom:8px;
}

.hero-title {
    font-size:clamp(1.9rem, 3vw, 3.05rem);
    line-height:1.04;
    font-weight:760;
}

.hero-sub {
    color:#9F9B90;
    margin-top:10px;
    font-size:.96rem;
    line-height:1.6;
    max-width:1100px;
}

.section-title {
    font-size:1.22rem;
    font-weight:730;
    margin-top:12px;
}

.section-note {
    color:#9F9B90;
    font-size:.87rem;
    margin-bottom:13px;
    line-height:1.55;
}

.kpi {
    border:1px solid rgba(212,175,55,.24);
    border-radius:17px;
    padding:15px 16px;
    min-height:112px;
    background:linear-gradient(
        145deg,
        #121212,
        #0D0D0D
    );
}

.kpi-label {
    color:#9F9B90;
    font-size:.72rem;
    letter-spacing:.04em;
    text-transform:uppercase;
    font-weight:650;
}

.kpi-value {
    font-size:1.55rem;
    font-weight:760;
    margin-top:7px;
}

.gold {
    color:#D4AF37;
}

.kpi-foot {
    color:#9F9B90;
    font-size:.71rem;
    margin-top:7px;
    line-height:1.35;
}

.info,
.callout {
    background:#0D0D0D;
    border:1px solid rgba(212,175,55,.24);
    border-left:3px solid #D4AF37;
    border-radius:12px;
    padding:13px 15px;
    color:#B9B5A9;
    font-size:.87rem;
    margin:8px 0 16px;
    line-height:1.65;
}

.callout {
    background:linear-gradient(
        145deg,
        rgba(212,175,55,.08),
        rgba(255,255,255,.01)
    );
    border-left:1px solid rgba(212,175,55,.24);
}

.small-gold {
    color:#D4AF37;
    font-weight:700;
}

[data-testid="stDataFrame"] {
    border:1px solid rgba(212,175,55,.20);
    border-radius:14px;
    overflow:hidden;
}

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# DATA DEFINITIONS
# ============================================================

REQUIRED = {
    "Emp Code",
    "Employee Name",
    "FY 26 TGT EQ NS",
    "Equity NS Ach YTD June",
}


FILTER_COLS = [
    "Status",
    "Type",
    "ZONE",
    "REGION",
]


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
# GENERAL HELPERS
# ============================================================

def norm(value):

    if value is None:
        return ""

    if (
        isinstance(value, float)
        and np.isnan(value)
    ):
        return ""

    return " ".join(
        str(value)
        .strip()
        .split()
    )


def unique_headers(values):

    seen = {}
    output = []

    for index, value in enumerate(values):

        base = (
            norm(value)
            or
            f"Unnamed_{index}"
        )

        if base not in seen:

            seen[base] = 0
            output.append(base)

        else:

            seen[base] += 1

            output.append(
                f"{base}.{seen[base]}"
            )

    return output


# ============================================================
# UPLOAD-ONLY EXCEL READER
# ============================================================

@st.cache_data(show_spinner=False)
def load_uploaded(file_bytes):

    raw = pd.read_excel(
        BytesIO(file_bytes),
        header=None,
        engine="openpyxl",
    )

    header = None

    for index in range(
        min(
            25,
            len(raw)
        )
    ):

        values = {
            norm(value)
            for value
            in raw.iloc[index].tolist()
        }

        if (
            "FY 26 TGT EQ NS" in values
            and
            "Equity NS Ach YTD June" in values
        ):

            header = index
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
        column
        for column
        in REQUIRED
        if column not in df.columns
    ]


    if missing:

        raise ValueError(
            "Missing columns: "
            +
            ", ".join(missing)
        )


    for column in [
        "FY 26 TGT EQ NS",
        "Equity NS Ach YTD June",
    ]:

        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        )


    df["Employee Name"] = (
        df["Employee Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    df["Emp Code"] = (
        df["Emp Code"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True,
        )
    )


    for column in FILTER_COLS:

        if column in df.columns:

            df[column] = (
                df[column]
                .fillna("Unknown")
                .astype(str)
                .str.strip()
            )

            df.loc[
                df[column].eq(""),
                column
            ] = "Unknown"


    # ========================================================
    # SECOND MKT TYPE COLUMN
    # ========================================================

    if "MKT TYPE.1" in df.columns:

        market_type_source = "MKT TYPE.1"

    elif "MKT TYPE" in df.columns:

        market_type_source = "MKT TYPE"

    else:

        market_type_source = None


    if market_type_source:

        df["Market Type"] = (
            df[
                market_type_source
            ]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
        )

        df.loc[
            df["Market Type"].eq(""),
            "Market Type"
        ] = "Unknown"

    else:

        df["Market Type"] = "Unknown"


    return (
        df,
        header + 1,
        market_type_source,
    )


# ============================================================
# NUMBER FORMATTING
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

        return (
            f"{value / 1_000_000:,.2f}M"
        )

    if abs(value) >= 1_000:

        return (
            f"{value / 1_000:,.2f}K"
        )

    return f"{value:,.2f}"


def pct(value):

    if (
        value is None
        or
        pd.isna(value)
    ):
        return "—"

    return (
        f"{float(value):,.1f}%"
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

    css_class = (
        "kpi-value gold"
        if accent
        else
        "kpi-value"
    )

    st.html(
        f"""
<div class="kpi">

    <div class="kpi-label">
        {label}
    </div>

    <div class="{css_class}">
        {value}
    </div>

    <div class="kpi-foot">
        {foot}
    </div>

</div>
"""
    )


def section(
    title,
    note="",
):

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


# ============================================================
# SAFE DATAFRAME
# ============================================================

def showdf(
    dataframe,
    height=None,
):

    display = dataframe.copy()

    numeric_columns = (
        display
        .select_dtypes(
            include=[
                np.number
            ]
        )
        .columns
    )

    display[
        numeric_columns
    ] = (
        display[
            numeric_columns
        ]
        .round(2)
    )

    kwargs = {
        "data": display,
        "width": "stretch",
        "hide_index": True,
    }

    if height is not None:

        kwargs[
            "height"
        ] = height

    st.dataframe(
        **kwargs
    )


# ============================================================
# PLOTLY STYLE
# ============================================================

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
        gridcolor=GRID,
        zeroline=False,
        linecolor=BORDER,
        tickfont=dict(
            color=MUTED
        ),
        title_font=dict(
            color=MUTED
        ),
    )


    fig.update_yaxes(
        showgrid=True,
        gridcolor=GRID,
        zeroline=False,
        linecolor=BORDER,
        tickfont=dict(
            color=MUTED
        ),
        title_font=dict(
            color=MUTED
        ),
    )


    return fig


# ============================================================
# SCENARIO CALCULATION
# ============================================================

def scenario(
    dataframe,
    months,
    uplift=0.0,
    threshold=100.0,
):

    df = dataframe.copy()


    # ========================================================
    # CURRENT MONTHLY RUN RATE
    # APR + MAY + JUN = 3 MONTHS
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
    # SCENARIO MONTHLY RUN RATE
    # ========================================================

    df[
        "Scenario Monthly RR"
    ] = (

        df[
            "Current Monthly RR"
        ]

        *

        (
            1
            +
            uplift
            /
            100.0
        )

    )


    # ========================================================
    # PROJECTED FUTURE NS
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


    # ========================================================
    # QUALIFICATION CHECK
    # ========================================================

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


    # ========================================================
    # RM CONTRIBUTION
    # ========================================================

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

            100

        )


    return df


# ============================================================
# BELL CURVE TABLE
# ============================================================

def bell_curve_analysis(
    scenario_df
):

    df = scenario_df.copy()

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


    lower_2 = (
        mean_value
        -
        2
        *
        std_value
    )


    lower_1 = (
        mean_value
        -
        std_value
    )


    upper_1 = (
        mean_value
        +
        std_value
    )


    upper_2 = (
        mean_value
        +
        2
        *
        std_value
    )


    bins = [
        -np.inf,
        lower_2,
        lower_1,
        mean_value,
        upper_1,
        upper_2,
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
        bins=bins,
        labels=labels,
        include_lowest=True,
    )


    total_rms = (
        len(df)
    )


    total_projected_ns = (
        df[
            "Projected Final NS"
        ]
        .sum()
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


    if total_rms > 0:

        result[
            "Actual RM Distribution %"
        ] = (

            result[
                "No. of RMs"
            ]

            /

            total_rms

            *

            100

        )

    else:

        result[
            "Actual RM Distribution %"
        ] = 0


    result[
        "Normal Curve Reference %"
    ] = (

        result[
            "Bell Curve Band"
        ]

        .map(
            normal_reference
        )

        .fillna(0)

    )


    if total_projected_ns != 0:

        result[
            "Projected NS Contribution %"
        ] = (

            result[
                "Projected NS"
            ]

            /

            total_projected_ns

            *

            100

        )

    else:

        result[
            "Projected NS Contribution %"
        ] = 0


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


    result = pd.concat(
        [
            result,
            total_row,
        ],
        ignore_index=True,
    )


    return (
        result,
        mean_value,
        median_value,
        std_value,
    )


# ============================================================
# BEFORE VS NOW BELL CURVE GRAPH
# ============================================================

def before_after_bell_curve(
    before_df,
    after_df,
):

    before_values = (

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


    after_values = (

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
        before_values.empty
        or
        after_values.empty
    ):

        fig.update_layout(
            title=
                "Bell Curve Shift · Before vs Now"
        )

        return style(
            fig,
            470
        )


    # ========================================================
    # STATISTICS
    # ========================================================

    before_mean = (
        before_values.mean()
    )


    before_std = (
        before_values.std()
    )


    after_mean = (
        after_values.mean()
    )


    after_std = (
        after_values.std()
    )


    if (
        pd.isna(
            before_std
        )
        or
        before_std == 0
    ):

        before_std = (
            0.000001
        )


    if (
        pd.isna(
            after_std
        )
        or
        after_std == 0
    ):

        after_std = (
            0.000001
        )


    before_std = float(
        before_std
    )


    after_std = float(
        after_std
    )


    # ========================================================
    # COMMON X RANGE
    # ========================================================

    minimum_x = min(

        before_values.min(),

        after_values.min(),

        before_mean
        -
        3
        *
        before_std,

        after_mean
        -
        3
        *
        after_std,

    )


    maximum_x = max(

        before_values.max(),

        after_values.max(),

        before_mean
        +
        3
        *
        before_std,

        after_mean
        +
        3
        *
        after_std,

    )


    x = np.linspace(
        minimum_x,
        maximum_x,
        500,
    )


    # ========================================================
    # FITTED NORMAL CURVE - BEFORE
    # ========================================================

    before_y = (

        1

        /

        (
            before_std
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
                    before_mean
                )

                /

                before_std

            )

            ** 2

        )

    )


    # ========================================================
    # FITTED NORMAL CURVE - AFTER
    # ========================================================

    after_y = (

        1

        /

        (
            after_std
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
                    after_mean
                )

                /

                after_std

            )

            ** 2

        )

    )


    # ========================================================
    # BEFORE CURVE
    # ========================================================

    fig.add_trace(

        go.Scatter(

            x=x,

            y=before_y,

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

            hovertemplate=(

                "Projected Achievement: %{x:.1f}%"

                "<br>Density: %{y:.4f}"

                "<extra>Before</extra>"

            ),

        )

    )


    # ========================================================
    # AFTER CURVE
    # ========================================================

    fig.add_trace(

        go.Scatter(

            x=x,

            y=after_y,

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

            hovertemplate=(

                "Projected Achievement: %{x:.1f}%"

                "<br>Density: %{y:.4f}"

                "<extra>Now</extra>"

            ),

        )

    )


    # ========================================================
    # MEAN LINES
    # ========================================================

    fig.add_vline(

        x=
            before_mean,

        line_width=2,

        line_dash=
            "dash",

        line_color=
            "#999999",

        annotation_text=(

            "Before Mean "

            f"{before_mean:.1f}%"

        ),

        annotation_position=
            "top left",

    )


    fig.add_vline(

        x=
            after_mean,

        line_width=2,

        line_dash=
            "dash",

        line_color=
            GOLD,

        annotation_text=(

            "Now Mean "

            f"{after_mean:.1f}%"

        ),

        annotation_position=
            "top right",

    )


    # ========================================================
    # TARGET LINE
    # ========================================================

    fig.add_vline(

        x=100,

        line_width=1.5,

        line_dash=
            "dot",

        line_color=
            TEXT,

        annotation_text=
            "100% Target",

        annotation_position=
            "bottom right",

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

        legend=dict(

            orientation=
                "h",

            yanchor=
                "bottom",

            y=
                1.02,

            xanchor=
                "left",

            x=
                0,

        ),

    )


    return style(
        fig,
        470
    )


# ============================================================
# BEFORE VS NOW SUMMARY
# ============================================================

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


    before_mean = (
        before.mean()
    )


    after_mean = (
        after.mean()
    )


    before_median = (
        before.median()
    )


    after_median = (
        after.median()
    )


    before_std = (
        before.std()
    )


    after_std = (
        after.std()
    )


    before_100 = int(
        (
            before
            >=
            100
        )
        .sum()
    )


    after_100 = int(
        (
            after
            >=
            100
        )
        .sum()
    )


    before_120 = int(
        (
            before
            >=
            120
        )
        .sum()
    )


    after_120 = int(
        (
            after
            >=
            120
        )
        .sum()
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

                before_mean,

                before_median,

                before_std,

                before_100,

                before_120,

            ],

            "Now": [

                after_mean,

                after_median,

                after_std,

                after_100,

                after_120,

            ],

            "Change": [

                after_mean
                -
                before_mean,

                after_median
                -
                before_median,

                after_std
                -
                before_std,

                after_100
                -
                before_100,

                after_120
                -
                before_120,

            ],

        }
    )


# ============================================================
# MARKET TYPE CONTRIBUTION
# ============================================================

def market_type_contribution_analysis(
    scenario_df
):

    df = scenario_df.copy()


    if (
        "Market Type"
        not in df.columns
    ):

        return pd.DataFrame()


    total_projected_ns = (
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


    total_current_achievement = (
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


    qualifying_summary = (

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
        qualifying_summary,
        on="Market Type",
        how="left",
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

        (

            result[
                "Currently Qualifying"
            ]

            /

            result[
                "#RMs"
            ]

            *

            100

        ),

        0,

    )


    if total_target != 0:

        result[
            "Target Contribution %"
        ] = (

            result[
                "Eq Target"
            ]

            /

            total_target

            *

            100

        )

    else:

        result[
            "Target Contribution %"
        ] = 0


    if total_current_achievement != 0:

        result[
            "Current Achievement Contribution %"
        ] = (

            result[
                "Current Achievement"
            ]

            /

            total_current_achievement

            *

            100

        )

    else:

        result[
            "Current Achievement Contribution %"
        ] = 0


    if total_projected_ns != 0:

        result[
            "Projected NS Contribution %"
        ] = (

            result[
                "Projected NS"
            ]

            /

            total_projected_ns

            *

            100

        )

    else:

        result[
            "Projected NS Contribution %"
        ] = 0


    if total_qualifying_ns != 0:

        result[
            "Qualifying NS Contribution %"
        ] = (

            result[
                "Qualifying Projected NS"
            ]

            /

            total_qualifying_ns

            *

            100

        )

    else:

        result[
            "Qualifying NS Contribution %"
        ] = 0


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
                    (
                        df[
                            "Crosses Threshold"
                        ].mean()
                        *
                        100
                    ),

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

    qualifying = (

        dataframe[
            dataframe[
                "Crosses Threshold"
            ]
        ]

        .copy()

    )


    if qualifying.empty:

        return pd.DataFrame()


    bins = [
        -np.inf,
        0.10,
        0.25,
        0.50,
        1,
        2,
        5,
        10,
        np.inf,
    ]


    labels = [
        "≤0.10%",
        "0.10–0.25%",
        "0.25–0.50%",
        "0.50–1.00%",
        "1.00–2.00%",
        "2.00–5.00%",
        "5.00–10.00%",
        ">10.00%",
    ]


    qualifying[
        "Contribution Bin"
    ] = pd.cut(

        qualifying[
            "Contribution %"
        ],

        bins=bins,

        labels=labels,

        include_lowest=True,

    )


    output = (

        qualifying

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


    total_row = pd.DataFrame(
        [
            {

                "Contribution Bin":
                    "Total",

                "No. of RMs":
                    output[
                        "No. of RMs"
                    ].sum(),

                "Projected NS":
                    output[
                        "Projected NS"
                    ].sum(),

                "Contribution %":
                    output[
                        "Contribution %"
                    ].sum(),

            }
        ]
    )


    return pd.concat(
        [
            output,
            total_row,
        ],
        ignore_index=True,
    )


# ============================================================
# PAGE 1 SCENARIO PANEL
# ============================================================

def scenario_panel(
    dataframe,
    name,
    threshold,
    months,
    uplift,
    market,
):

    qualifying = (
        dataframe[
            dataframe[
                "Crosses Threshold"
            ]
        ]
    )


    non_qualifying = (
        dataframe[
            ~dataframe[
                "Crosses Threshold"
            ]
        ]
    )


    hit_rate = (

        len(
            qualifying
        )

        /

        len(
            dataframe
        )

        *

        100

        if len(
            dataframe
        )

        else
        0

    )


    section(

        name,

        (
            f"{market} · "
            f"through {MONTHS[months]} · "
            f"uplift {uplift:.1f}% · "
            f"threshold {threshold:.1f}%"
        ),

    )


    # ========================================================
    # KPI ROW
    # ========================================================

    columns = st.columns(
        6
    )


    metrics = [

        (
            "RMs Analysed",
            len(
                dataframe
            ),
            market,
            False,
        ),

        (
            "RMs Crossing",
            len(
                qualifying
            ),
            f"{hit_rate:.1f}% of RMs",
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
            "Can include negative drag",
            False,
        ),

        (
            "All RM Projected NS",
            fmt(
                dataframe[
                    "Projected Final NS"
                ]
                .sum()
            ),
            (
                "Target "
                +
                fmt(
                    dataframe[
                        "FY 26 TGT EQ NS"
                    ]
                    .sum()
                )
            ),
            False,
        ),

        (
            "Median Achievement",
            pct(
                dataframe[
                    "Projected Achievement %"
                ]
                .median()
            ),
            "Projected FY achievement",
            False,
        ),

    ]


    for index, metric in enumerate(
        metrics
    ):

        with columns[index]:

            kpi(
                *metric
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
scenario monthly RR ×
selected future months.

<br>

Achievement % =
projected final NS ÷
FY target × 100.

<br>

Contribution % =
individual qualifying RM projected NS ÷
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
            "Projected Achievement % is grouped around the actual "
            "population mean and standard deviation. The normal-curve "
            "reference is shown only as a benchmark."
        ),

    )


    (
        bell_table,
        bell_mean,
        bell_median,
        bell_std,
    ) = bell_curve_analysis(
        dataframe
    )


    bell_kpis = st.columns(
        4
    )


    with bell_kpis[0]:

        kpi(
            "Mean Achievement",
            pct(
                bell_mean
            ),
            "Average projected achievement",
            False,
        )


    with bell_kpis[1]:

        kpi(
            "Median Achievement",
            pct(
                bell_median
            ),
            "50th percentile RM",
            False,
        )


    with bell_kpis[2]:

        kpi(
            "Standard Deviation",
            pct(
                bell_std
            ),
            "Spread of RM achievement",
            False,
        )


    with bell_kpis[3]:

        above_mean_count = int(
            (
                dataframe[
                    "Projected Achievement %"
                ]
                >=
                bell_mean
            )
            .sum()
        )


        above_share = (

            above_mean_count

            /

            len(
                dataframe
            )

            *

            100

            if len(
                dataframe
            )

            else
            0

        )


        kpi(
            "RMs Above Mean",
            f"{above_mean_count:,}",
            f"{above_share:.1f}% of RMs",
            True,
        )


    if bell_table.empty:

        st.info(
            "Bell curve analysis is unavailable."
        )

    else:

        showdf(
            bell_table
        )


    # ========================================================
    # MARKET TYPE CONTRIBUTION
    # ========================================================

    section(

        "Market Type Contribution Analysis",

        (
            "Contribution of T2, T6, B30, EM, T30 and other "
            "market types to target, current achievement, "
            "projected NS and qualifying NS."
        ),

    )


    market_table = (
        market_type_contribution_analysis(
            dataframe
        )
    )


    if market_table.empty:

        st.info(
            "Market Type data is not available."
        )

    else:

        showdf(
            market_table
        )


    # ========================================================
    # INDIVIDUAL RM CONTRIBUTION BINS
    # ========================================================

    section(

        "Individual RM Contribution Bins",

        (
            "Among qualifying RMs only, this shows how many "
            "people sit in each sales-contribution band."
        ),

    )


    bins = (
        contribution_bins(
            dataframe
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
    # QUALIFYING RM DETAIL
    # ========================================================

    section(
        "Qualifying RM Detail",
        "RM-level output for the selected filters.",
    )


    if qualifying.empty:

        st.info(
            "No qualifying RMs."
        )

    else:

        display_columns = [
            "Emp Code",
            "Employee Name",
            "Market Type",
        ]


        display_columns += [
            column
            for column
            in FILTER_COLS
            if column
            in qualifying.columns
        ]


        display_columns += [
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
                display_columns
            ]

            .sort_values(
                "Contribution %",
                ascending=False,
            ),

            520,

        )


# ============================================================
# TARGET BUCKET OPTION 1
# ============================================================

def bucket_table_option1(
    base_df,
    step=5.0,
    top=50.0,
):

    target = (
        base_df[
            "FY 26 TGT EQ NS"
        ]
    )


    rows = []
    upper_bounds = []

    current = (
        step
    )


    while current < top:

        upper_bounds.append(
            current
        )

        current += (
            step
        )


    upper_bounds.append(
        top
    )


    lower = (
        0.0
    )


    for upper in upper_bounds:

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
                    f"Upto {upper:g}",

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


    mask = (
        target
        >
        top
    )


    rows.append(
        {

            "Current Target Bucketing":
                f"Above {top:g}",

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


# ============================================================
# TARGET BUCKET OPTION 2
# ============================================================

def bucket_table_option2(
    base_df,
    current_scenario,
    cutoff=6.5,
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
            "Crosses Threshold"
        ]
        .to_numpy()
    )


    rows = []


    specifications = [

        (
            f"Upto {cutoff:g} crores",
            target
            <=
            cutoff,
        ),

        (
            f"{cutoff:g} and above",
            target
            >
            cutoff,
        ),

    ]


    for label, mask in specifications:

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
                        np.sum(
                            qualifies
                            &
                            mask.to_numpy()
                        )
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
                    current_scenario[
                        "Crosses Threshold"
                    ]
                    .sum()
                ),

        }
    )


    return pd.DataFrame(
        rows
    )


# ============================================================
# TARGET BUCKET OPTION 3
# ============================================================

def bucket_table_option3(
    base_df,
    cutoff_1=5.0,
    cutoff_2=10.0,
):

    target = (
        base_df[
            "FY 26 TGT EQ NS"
        ]
    )


    specifications = [

        (
            f"Upto {cutoff_1:g} crores",

            target
            <=
            cutoff_1,
        ),

        (
            f"{cutoff_1:g} to {cutoff_2:g} crores",

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
            f"Above {cutoff_2:g} crores",

            target
            >
            cutoff_2,
        ),

    ]


    rows = []


    for label, mask in specifications:

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
# QUALIFICATION + TRAVEL BUDGET
# ============================================================

def qualification_budget_section(
    base_df,
    current_scenario,
    uplift_scenario,
    projection_months,
):

    section(

        "Qualification & Travel Budget Planner",

        (
            "Convert qualification counts into an editable "
            "foreign-trip budget and inspect the RM population "
            "through different current-target buckets."
        ),

    )


    qualified_count = int(
        current_scenario[
            "Crosses Threshold"
        ]
        .sum()
    )


    scenario_incremental_ns = (

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


    control_1, control_2, metric_1, metric_2 = (
        st.columns(
            4
        )
    )


    with control_1:

        travelers = st.number_input(

            "Number of People for Trip",

            min_value=0,

            max_value=max(
                len(
                    base_df
                ),
                1
            ),

            value=min(
                40,
                len(
                    base_df
                )
            ),

            step=1,

            key=
                "budget_people",

        )


    with control_2:

        cost_lakh = st.number_input(

            "Trip Cost per Person (₹ lakh)",

            min_value=0.0,

            max_value=100.0,

            value=3.0,

            step=0.25,

            key=
                "budget_cost_lakh",

        )


    total_budget_crore = (

        travelers

        *

        cost_lakh

        /

        100.0

    )


    if scenario_incremental_ns > 0:

        budget_percentage = (

            total_budget_crore

            /

            scenario_incremental_ns

            *

            100

        )

    else:

        budget_percentage = (
            np.nan
        )


    with metric_1:

        kpi(
            "Trip Budget",
            f"₹{total_budget_crore:,.2f} Cr",
            (
                f"{travelers} people × "
                f"₹{cost_lakh:.2f} lakh"
            ),
            True,
        )


    with metric_2:

        kpi(
            "Budget / Incremental NS",
            pct(
                budget_percentage
            ),
            "Scenario 2 incremental NS used as denominator",
            False,
        )


    st.html(
        f"""
<div class="callout">

<b style="color:#D4AF37">
Current qualification
</b>

<br>

{qualified_count:,} RMs are projected to cross
100% at the current run rate through
{MONTHS[projection_months]}.

<br><br>

Example:

<b style="color:#F3F0E7">
40 people × ₹3 lakh = ₹1.20 Cr
</b>

</div>
"""
    )


    # ========================================================
    # TARGET BUCKET OPTIONS
    # ========================================================

    section(

        "Current Target Bucketing Options",

        (
            "Compare alternative ways of splitting the RM "
            "population by current Equity Net Sales target."
        ),

    )


    option_1_column, option_2_column, option_3_column = (
        st.columns(
            3
        )
    )


    # ========================================================
    # OPTION 1
    # ========================================================

    with option_1_column:

        st.markdown(
            "#### Option 1 · Detailed Bands"
        )


        bucket_step = st.number_input(

            "Bucket Step (Cr)",

            min_value=1.0,

            max_value=25.0,

            value=5.0,

            step=1.0,

            key=
                "option1_bucket_step",

        )


        last_cutoff = st.number_input(

            "Last Cutoff Before Above Bucket (Cr)",

            min_value=
                bucket_step,

            max_value=200.0,

            value=50.0,

            step=5.0,

            key=
                "option1_last_cutoff",

        )


        option_1 = (
            bucket_table_option1(
                base_df,
                bucket_step,
                last_cutoff,
            )
        )


        showdf(
            option_1
        )


    # ========================================================
    # OPTION 2
    # ========================================================

    with option_2_column:

        st.markdown(
            "#### Option 2 · Two Cohorts"
        )


        option_2_cutoff = st.number_input(

            "Split Cutoff (Cr)",

            min_value=0.5,

            max_value=100.0,

            value=6.5,

            step=0.5,

            key=
                "option2_cutoff",

        )


        option_2 = (
            bucket_table_option2(
                base_df,
                current_scenario,
                option_2_cutoff,
            )
        )


        showdf(
            option_2
        )


    # ========================================================
    # OPTION 3
    # ========================================================

    with option_3_column:

        st.markdown(
            "#### Option 3 · Three Cohorts"
        )


        option_3_cutoff_1 = st.number_input(

            "First Cutoff (Cr)",

            min_value=0.5,

            max_value=100.0,

            value=5.0,

            step=0.5,

            key=
                "option3_cutoff1",

        )


        option_3_cutoff_2 = st.number_input(

            "Second Cutoff (Cr)",

            min_value=
                option_3_cutoff_1
                +
                0.5,

            max_value=200.0,

            value=max(
                10.0,
                option_3_cutoff_1 + 0.5
            ),

            step=0.5,

            key=
                "option3_cutoff2",

        )


        option_3 = (
            bucket_table_option3(
                base_df,
                option_3_cutoff_1,
                option_3_cutoff_2,
            )
        )


        showdf(
            option_3
        )


    # ========================================================
    # EDITABLE COHORT PLAN
    # ========================================================

    section(

        "Editable Cohort Planning",

        (
            "Change proposed stretch % and planned traveler "
            "count for each target cohort."
        ),

    )


    editable_plan = (
        option_3
        .iloc[:-1]
        .copy()
    )


    default_stretches = [
        35.0,
        25.0,
        15.0,
    ]


    editable_plan[
        "Proposed Stretch %"
    ] = (
        default_stretches[
            :len(
                editable_plan
            )
        ]
    )


    editable_plan[
        "Planned Travelers"
    ] = [

        min(
            10,
            int(value)
        )

        for value
        in editable_plan[
            "#RMs"
        ]

    ]


    edited_plan = st.data_editor(

        editable_plan,

        width=
            "stretch",

        hide_index=
            True,

        key=
            "cohort_plan_editor",

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


    calculated_plan = (
        edited_plan.copy()
    )


    calculated_plan[
        "Implied Extra Target"
    ] = (

        calculated_plan[
            "Eq Target"
        ]

        *

        calculated_plan[
            "Proposed Stretch %"
        ]

        /

        100.0

    )


    calculated_plan[
        "Trip Budget (Cr)"
    ] = (

        calculated_plan[
            "Planned Travelers"
        ]

        *

        cost_lakh

        /

        100.0

    )


    calculated_plan[
        "Extra Target / Trip Budget (x)"
    ] = np.where(

        calculated_plan[
            "Trip Budget (Cr)"
        ]
        > 0,

        (

            calculated_plan[
                "Implied Extra Target"
            ]

            /

            calculated_plan[
                "Trip Budget (Cr)"
            ]

        ),

        np.nan,

    )


    showdf(
        calculated_plan
    )


# ============================================================
# PAGE 2 SUMMARY
# ============================================================

def summary_by(
    dataframe,
    dimension,
):

    if dimension not in dataframe.columns:

        return pd.DataFrame()


    result = (

        dataframe

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


    return (
        result
        .sort_values(
            "Projected_NS",
            ascending=False,
        )
    )


# ============================================================
# PAGE 2
# ============================================================

def new_insights(
    base,
    months,
    market,
):

    dataframe = scenario(
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
        geography, quick wins and
        high-value risk.

    </div>

</div>
"""
    )


    columns = st.columns(
        5
    )


    metrics = [

        (
            "FY Target",
            fmt(
                dataframe[
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
                dataframe[
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
                dataframe[
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
                dataframe[
                    "Crosses Threshold"
                ]
                .sum()
            ),
            (
                f"{dataframe['Crosses Threshold'].mean() * 100:.1f}% "
                "hit-rate"
            ),
            True,
        ),

        (
            "Median Achievement",
            pct(
                dataframe[
                    "Projected Achievement %"
                ]
                .median()
            ),
            "Current-RR projection",
            False,
        ),

    ]


    for index, metric in enumerate(
        metrics
    ):

        with columns[index]:

            kpi(
                *metric
            )


    # ========================================================
    # CUT TABLES
    # ========================================================

    for dimension, title in [

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

        if dimension in dataframe.columns:

            section(

                f"{title} Cut",

                (
                    f"Target, YTD, projection and "
                    f"hit-rate by {title.lower()}."
                ),

            )


            showdf(
                summary_by(
                    dataframe,
                    dimension,
                )
            )


    # ========================================================
    # TARGET VS PROJECTED
    # ========================================================

    section(

        "Target vs Projected NS",

        (
            "Dots above the diagonal are projected "
            "to beat their existing target."
        ),

    )


    fig = go.Figure(

        go.Scatter(

            x=
                dataframe[
                    "FY 26 TGT EQ NS"
                ],

            y=
                dataframe[
                    "Projected Final NS"
                ],

            mode=
                "markers",

            marker=dict(

                size=8,

                color=
                    dataframe[
                        "Projected Achievement %"
                    ],

                colorscale=
                    "Cividis",

                showscale=True,

            ),

            text=
                dataframe[
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


    maximum = max(

        dataframe[
            "FY 26 TGT EQ NS"
        ].max(),

        dataframe[
            "Projected Final NS"
        ].max(),

        1,

    )


    fig.add_scatter(

        x=[
            0,
            maximum
        ],

        y=[
            0,
            maximum
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
            "displayModeBar": False
        },
    )


    # ========================================================
    # REQUIRED RR
    # ========================================================

    dataframe[
        "Required Future RR for 100%"
    ] = (

        (

            dataframe[
                "FY 26 TGT EQ NS"
            ]

            -

            dataframe[
                "Equity NS Ach YTD June"
            ]

        )

        /

        months

    )


    dataframe[
        "Required RR Uplift %"
    ] = np.where(

        dataframe[
            "Current Monthly RR"
        ]
        > 0,

        (

            dataframe[
                "Required Future RR for 100%"
            ]

            /

            dataframe[
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

    quick_wins = (

        dataframe[

            (
                dataframe[
                    "Projected Achievement %"
                ]
                >=
                80
            )

            &

            (
                dataframe[
                    "Projected Achievement %"
                ]
                <
                100
            )

            &

            (
                dataframe[
                    "Required RR Uplift %"
                ]
                .between(
                    0,
                    30
                )
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


    if quick_wins.empty:

        st.info(
            "No quick wins under current filters."
        )

    else:

        display_columns = [
            "Emp Code",
            "Employee Name",
            "Market Type",
        ]


        display_columns += [
            column
            for column
            in [
                "ZONE",
                "REGION",
            ]
            if column
            in quick_wins.columns
        ]


        display_columns += [
            "FY 26 TGT EQ NS",
            "Projected Final NS",
            "Projected Achievement %",
            "Required RR Uplift %",
        ]


        showdf(

            quick_wins[
                display_columns
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

        dataframe[
            dataframe[
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
            "Prioritize absolute sales value at risk "
            "rather than percentage miss alone."
        ),

    )


    if risk.empty:

        st.success(
            "No RMs below 80%."
        )

    else:

        display_columns = [
            "Emp Code",
            "Employee Name",
            "Market Type",
        ]


        display_columns += [
            column
            for column
            in [
                "ZONE",
                "REGION",
            ]
            if column
            in risk.columns
        ]


        display_columns += [
            "FY 26 TGT EQ NS",
            "Projected Final NS",
            "Projected Achievement %",
            "Target at Risk",
        ]


        showdf(

            risk

            .nlargest(
                30,
                "Target at Risk",
            )[
                display_columns
            ],

            420,

        )


# ============================================================
# BONVOYAGE MODEL
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

    dataframe = scenario(
        base,
        months,
        0,
        100,
    )


    # ========================================================
    # NO-INCENTIVE EXPECTATION
    # ========================================================

    dataframe[
        "No-Incentive Expected NS"
    ] = np.maximum(

        dataframe[
            "FY 26 TGT EQ NS"
        ],

        dataframe[
            "Projected Final NS"
        ],

    )


    # ========================================================
    # TARGET SIZE PERCENTILE
    # ========================================================

    dataframe[
        "Target Size Percentile"
    ] = (

        dataframe[
            "FY 26 TGT EQ NS"
        ]

        .rank(
            pct=True,
            method="average",
        )

    )


    # ========================================================
    # PEER PERFORMANCE PERCENTILE
    # ========================================================

    dataframe[
        "Peer Performance Percentile"
    ] = (

        dataframe

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

        dataframe[
            "Projected Achievement %"
        ]

        /

        120

    ).clip(
        0,
        1,
    ).fillna(
        0
    )


    # ========================================================
    # CAPACITY SCORE
    # ========================================================

    dataframe[
        "Capacity Score"
    ] = (

        0.45

        *

        (
            1
            -
            dataframe[
                "Target Size Percentile"
            ]
        )

        +

        0.35

        *

        dataframe[
            "Peer Performance Percentile"
        ]

        +

        0.20

        *

        momentum

    ).clip(
        0,
        1,
    )


    # ========================================================
    # PLANNED FUTURE RR UPLIFT
    # ========================================================

    dataframe[
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

        dataframe[
            "Capacity Score"
        ]

    )


    # ========================================================
    # CAPACITY TARGET
    # ========================================================

    dataframe[
        "Capacity-Based Target"
    ] = (

        dataframe[
            "Equity NS Ach YTD June"
        ]

        +

        dataframe[
            "Current Monthly RR"
        ]

        *

        (

            1

            +

            dataframe[
                "Planned Future RR Uplift %"
            ]

            /

            100

        )

        *

        months

    )


    # ========================================================
    # STRETCH TARGET
    # ========================================================

    dataframe[
        "BonVoyage Stretch Target"
    ] = np.maximum(

        dataframe[
            "No-Incentive Expected NS"
        ],

        dataframe[
            "Capacity-Based Target"
        ],

    )


    # ========================================================
    # TRUE INCREMENTAL NS
    # ========================================================

    dataframe[
        "Incremental NS Required"
    ] = (

        dataframe[
            "BonVoyage Stretch Target"
        ]

        -

        dataframe[
            "No-Incentive Expected NS"
        ]

    ).clip(
        lower=0
    )


    # ========================================================
    # REQUIRED FUTURE RUN RATE
    # ========================================================

    dataframe[
        "Required Future RR for BonVoyage"
    ] = (

        (

            dataframe[
                "BonVoyage Stretch Target"
            ]

            -

            dataframe[
                "Equity NS Ach YTD June"
            ]

        )

        /

        months

    )


    dataframe[
        "Required Future RR Uplift %"
    ] = np.where(

        dataframe[
            "Current Monthly RR"
        ]
        > 0,

        (

            dataframe[
                "Required Future RR for BonVoyage"
            ]

            /

            dataframe[
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
    # TRIP ECONOMICS
    # ========================================================

    trip_cost_crore = (
        trip_lakh
        /
        100
    )


    dataframe[
        "Trip Budget Ceiling"
    ] = (

        dataframe[
            "Incremental NS Required"
        ]

        *

        allocation

        /

        100

    )


    dataframe[
        "Trip Funding Coverage x"
    ] = np.where(

        trip_cost_crore
        >
        0,

        (

            dataframe[
                "Trip Budget Ceiling"
            ]

            /

            trip_cost_crore

        ),

        np.inf,

    )


    dataframe[
        "Recommended Candidate"
    ] = (

        (
            dataframe[
                "Trip Funding Coverage x"
            ]
            >=
            1
        )

        &

        (
            dataframe[
                "Required Future RR Uplift %"
            ]
            .between(
                0,
                max_feasible
            )
        )

        &

        (
            dataframe[
                "Incremental NS Required"
            ]
            >
            0
        )

        &

        (
            dataframe[
                "Current Monthly RR"
            ]
            >
            0
        )

    )


    return dataframe


# ============================================================
# BONVOYAGE PAGE
# ============================================================

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

    dataframe = bonvoyage_model(
        base,
        months,
        min_uplift,
        max_uplift,
        allocation,
        trip_lakh,
        max_feasible,
    )


    candidates = (
        dataframe[
            dataframe[
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

        Targets are personalized by RM capacity.
        Sales the RM was already expected to generate
        are not counted as BonVoyage incremental sales.

    </div>

</div>
"""
    )


    columns = st.columns(
        5
    )


    metrics = [

        (
            "Official Target",
            fmt(
                dataframe[
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
                dataframe[
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
                dataframe[
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
                dataframe[
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
            f"≤{max_feasible}% required RR uplift",
            False,
        ),

    ]


    for index, metric in enumerate(
        metrics
    ):

        with columns[index]:

            kpi(
                *metric
            )


    # ========================================================
    # BONVOYAGE SCATTER
    # ========================================================

    section(

        "BonVoyage Target vs No-Incentive Baseline",

        (
            "The farther a point sits above the diagonal, "
            "the larger the genuinely incremental target."
        ),

    )


    fig = go.Figure(

        go.Scatter(

            x=
                dataframe[
                    "No-Incentive Expected NS"
                ],

            y=
                dataframe[
                    "BonVoyage Stretch Target"
                ],

            mode=
                "markers",

            marker=dict(

                size=np.clip(

                    7

                    +

                    dataframe[
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
                    dataframe[
                        "Required Future RR Uplift %"
                    ],

                colorscale=
                    "Cividis",

                showscale=True,

                colorbar=dict(
                    title=
                        "RR uplift %"
                ),

            ),

            text=
                dataframe[
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


    maximum = max(

        dataframe[
            "BonVoyage Stretch Target"
        ].max(),

        dataframe[
            "No-Incentive Expected NS"
        ].max(),

        1,

    )


    fig.add_scatter(

        x=[
            0,
            maximum
        ],

        y=[
            0,
            maximum
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
            "displayModeBar": False
        },

    )


    # ========================================================
    # RM TARGET BOOK
    # ========================================================

    section(

        "RM-Level BonVoyage Target Book",

        (
            "Personalized target, incremental ask, "
            "trip-funding coverage and eligibility."
        ),

    )


    display_columns = [
        "Emp Code",
        "Employee Name",
        "Market Type",
    ]


    display_columns += [
        column
        for column
        in [
            "ZONE",
            "REGION",
        ]
        if column
        in dataframe.columns
    ]


    display_columns += [
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

        dataframe[
            display_columns
        ]

        .sort_values(
            "Incremental NS Required",
            ascending=False,
        ),

        620,

    )


    # ========================================================
    # CUTS
    # ========================================================

    for dimension in [
        "Market Type",
        "ZONE",
        "REGION",
    ]:

        if dimension not in dataframe.columns:

            continue


        section(

            f"{dimension} BonVoyage Economics",

            (
                f"Incremental opportunity and "
                f"candidates by {dimension}."
            ),

        )


        summary = (

            dataframe

            .groupby(
                dimension
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
# MAIN HEADER
# ============================================================

st.html(
    """
<div class="hero">

    <div class="eyebrow">
        RM Equity Net Sales · Strategy Lab
    </div>

    <div class="hero-title">
        Target Analysis
    </div>

    <div class="hero-sub">

        Upload the RM workbook once.
        The application reads only the uploaded workbook.

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

        help=(
            "The app reads only the workbook uploaded here."
        ),

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


# ============================================================
# READ WORKBOOK
# ============================================================

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
        f"Could not read workbook: {error}"
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
# SIDEBAR FILTERS
# ============================================================

with st.sidebar:

    st.divider()

    st.markdown(
        "### Global Filters"
    )


    # ========================================================
    # MARKET TYPE
    # ========================================================

    market_values = sorted(

        [
            value

            for value
            in data[
                "Market Type"
            ]
            .dropna()
            .astype(str)
            .unique()

            if value
            !=
            "Unknown"
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

    filter_definitions = [

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
    ) in filter_definitions:


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
    # PROJECTION MONTHS
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

            lambda month:

            (
                f"{month} months · "
                f"through {MONTHS[month]}"
            ),

    )


    # ========================================================
    # SCENARIO 2
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
    # BONVOYAGE DEFAULTS
    # ========================================================

    minimum_uplift = (
        15
    )

    maximum_uplift = (
        60
    )

    allocation = (
        10
    )

    trip_lakh = (
        3.0
    )

    maximum_feasible = (
        50
    )


    if page == "3 · BonVoyage":

        st.divider()

        st.markdown(
            "### BonVoyage Economics"
        )


        minimum_uplift = st.slider(

            "Minimum Planned Future RR Uplift (%)",

            0,

            50,

            15,

        )


        maximum_uplift = st.slider(

            "Maximum Planned Future RR Uplift (%)",

            min_value=max(
                minimum_uplift,
                10
            ),

            max_value=100,

            value=max(
                60,
                minimum_uplift
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


        maximum_feasible = st.slider(

            "Maximum Feasible Future RR Uplift (%)",

            5,

            100,

            50,

            5,

        )


    st.divider()


    st.caption(

        (
            f"Uploaded {uploaded.name} | "
            f"Header row {header_row} | "
            f"MKT source {market_type_source} | "
            f"Excluded missing {excluded_missing} | "
            f"Excluded target≤0 {excluded_target}"
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
# PAGE 1
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


    tab_1, tab_2, tab_3 = st.tabs(
        [
            "Scenario 1 · Current Run Rate",
            "Scenario 2 · Increased Run Rate",
            "Scenario Comparison",
        ]
    )


    # ========================================================
    # SCENARIO 1
    # ========================================================

    with tab_1:

        scenario_panel(
            scenario_1,
            "Scenario 1 · Current Run Rate",
            100,
            months,
            0,
            market,
        )


    # ========================================================
    # SCENARIO 2
    # ========================================================

    with tab_2:

        scenario_panel(
            scenario_2,
            "Scenario 2 · Increased Run Rate",
            threshold,
            months,
            uplift,
            market,
        )


    # ========================================================
    # COMPARISON
    # ========================================================

    with tab_3:

        section(

            "Scenario Comparison",

            (
                f"Current monthly run rate vs "
                f"{uplift:.1f}% higher run rate."
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

                    scenario_1[
                        "Crosses Threshold"
                    ].sum(),

                    scenario_1[
                        "Crosses Threshold"
                    ].mean()
                    *
                    100,

                    scenario_1.loc[
                        scenario_1[
                            "Crosses Threshold"
                        ],
                        "Projected Final NS",
                    ].sum(),

                    scenario_1[
                        "Projected Final NS"
                    ].sum(),

                    scenario_1[
                        "Projected Achievement %"
                    ].median(),

                ],


                "Scenario 2": [

                    scenario_2[
                        "Crosses Threshold"
                    ].sum(),

                    scenario_2[
                        "Crosses Threshold"
                    ].mean()
                    *
                    100,

                    scenario_2.loc[
                        scenario_2[
                            "Crosses Threshold"
                        ],
                        "Projected Final NS",
                    ].sum(),

                    scenario_2[
                        "Projected Final NS"
                    ].sum(),

                    scenario_2[
                        "Projected Achievement %"
                    ].median(),

                ],

            }
        )


        showdf(
            comparison
        )


    # ========================================================
    # BELL CURVE SHIFT
    #
    # OUTSIDE TABS - APPEARS ONLY ONCE
    # ========================================================

    st.divider()


    section(

        "Bell Curve Shift · Before vs Now",

        (
            "Before represents the current-run-rate projection. "
            f"Now represents the {uplift:.1f}% increased-run-rate "
            "scenario. A rightward movement means the RM population "
            "is moving towards higher achievement."
        ),

    )


    before_values = (

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


    after_values = (

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
        before_values.mean()
    )


    after_mean = (
        after_values.mean()
    )


    before_std = (
        before_values.std()
    )


    after_std = (
        after_values.std()
    )


    before_qualifying = int(
        (
            before_values
            >=
            100
        )
        .sum()
    )


    after_qualifying = int(
        (
            after_values
            >=
            100
        )
        .sum()
    )


    curve_kpis = st.columns(
        4
    )


    with curve_kpis[0]:

        kpi(
            "Before Mean",
            pct(
                before_mean
            ),
            "Current run-rate distribution",
            False,
        )


    with curve_kpis[1]:

        kpi(

            "Now Mean",

            pct(
                after_mean
            ),

            (
                f"{after_mean - before_mean:+.1f} "
                "percentage-point movement"
            ),

            True,

        )


    with curve_kpis[2]:

        spread_change = (
            after_std
            -
            before_std
        )


        if spread_change > 0:

            spread_text = (
                "Wider spread"
            )

        elif spread_change < 0:

            spread_text = (
                "Narrower spread"
            )

        else:

            spread_text = (
                "No spread change"
            )


        kpi(

            "Spread Change",

            f"{spread_change:+.1f} pp",

            (
                f"{spread_text} · "
                f"{before_std:.1f} → {after_std:.1f}"
            ),

            False,

        )


    with curve_kpis[3]:

        kpi(

            "Additional RMs ≥100%",

            f"{after_qualifying - before_qualifying:+,}",

            (
                f"{before_qualifying:,} → "
                f"{after_qualifying:,} RMs"
            ),

            True,

        )


    # ========================================================
    # GRAPH
    # ========================================================

    st.plotly_chart(

        before_after_bell_curve(
            scenario_1,
            scenario_2,
        ),

        config={
            "displayModeBar": False
        },

    )


    st.caption(
        "The smooth curves are fitted normal curves using the actual "
        "mean and standard deviation of each scenario. The Bell Curve "
        "tables inside Scenario 1 and Scenario 2 show the actual RM "
        "population counts in each μ/σ band."
    )


    # ========================================================
    # DISTRIBUTION SUMMARY
    # ========================================================

    distribution_summary = (
        before_after_distribution_summary(
            scenario_1,
            scenario_2,
        )
    )


    showdf(
        distribution_summary
    )


    # ========================================================
    # INTERPRETATION
    # ========================================================

    mean_shift = (
        after_mean
        -
        before_mean
    )


    std_shift = (
        after_std
        -
        before_std
    )


    if mean_shift > 0:

        direction_text = (

            "The overall RM distribution shifts to the right. "
            "This means average projected target achievement "
            "improves under the selected uplift scenario."

        )

    elif mean_shift < 0:

        direction_text = (

            "The overall RM distribution shifts to the left. "
            "Average projected achievement is lower."

        )

    else:

        direction_text = (

            "There is no meaningful change in the centre "
            "of the distribution."

        )


    if std_shift < 0:

        spread_interpretation = (

            "The curve also becomes narrower, indicating that "
            "performance is becoming more consistent across RMs."

        )

    elif std_shift > 0:

        spread_interpretation = (

            "The curve becomes wider. The average improves, but "
            "the difference between stronger and weaker RMs also increases."

        )

    else:

        spread_interpretation = (

            "The overall spread of RM performance remains unchanged."

        )


    st.html(
        f"""
<div class="callout">

<b style="color:#D4AF37">
What changed?
</b>

<br><br>

{direction_text}

<br><br>

{spread_interpretation}

<br><br>

Mean Achievement:

<b style="color:#F3F0E7">
{before_mean:.1f}% → {after_mean:.1f}%
</b>

<br>

Standard Deviation:

<b style="color:#F3F0E7">
{before_std:.1f} → {after_std:.1f}
</b>

<br>

RMs ≥100%:

<b style="color:#F3F0E7">
{before_qualifying:,} → {after_qualifying:,}
</b>

</div>
"""
    )


    # ========================================================
    # TRAVEL BUDGET SECTION
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

else:

    bonvoyage_page(
        data,
        months,
        market,
        minimum_uplift,
        maximum_uplift,
        allocation,
        trip_lakh,
        maximum_feasible,
    )
