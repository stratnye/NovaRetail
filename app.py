"""
NovaRetail Interactive Dashboard
Built for: Sophia Martinez, Director of Customer Intelligence

Answers:
  - Which segments generate the most revenue, and by what dimensions?
  - Which segments are at risk (Decline label, low CSAT)?
  - Where should the company focus investment (Stable/Growth/Promising,
    high CSAT, high purchase amount)?

Data: NR_dataset.xlsx must sit in the same directory as this file.
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="NovaRetail | Customer Intelligence",
    page_icon="\U0001F4CA",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------------
# DESIGN SYSTEM — fonts, colors, CSS
# ----------------------------------------------------------------------------
COLOR_INK = "#1B2430"          # near-black navy, sidebar / headers
COLOR_BG = "#F6F5F1"           # warm paper background for content
COLOR_CARD = "#FFFFFF"
COLOR_LINE = "#E4E1D8"
COLOR_ACCENT = "#3A6EA5"       # steel blue — primary accent
COLOR_MUTED = "#6B7280"

# Segment palette: deliberate semantic progression, risk -> opportunity
SEGMENT_COLORS = {
    "Decline": "#C4433A",      # rust red   — at risk
    "Stable": "#7A8CA3",       # slate blue — holding steady
    "Growth": "#3F8F5F",       # green      — trending up
    "Promising": "#D6A03C",    # gold       — highest opportunity
}
SEGMENT_ORDER = ["Decline", "Stable", "Growth", "Promising"]

CUSTOM_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
}}
h1, h2, h3, .hero-title {{
    font-family: 'Fraunces', serif;
    letter-spacing: -0.01em;
}}
.stApp {{
    background-color: {COLOR_BG};
}}
section[data-testid="stSidebar"] {{
    background-color: {COLOR_INK};
}}
section[data-testid="stSidebar"] * {{
    color: #E9E7DF !important;
}}
section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] {{
    background-color: {COLOR_ACCENT} !important;
}}
.hero-band {{
    background: linear-gradient(135deg, {COLOR_INK} 0%, #2C3A4E 100%);
    padding: 1.6rem 2rem;
    border-radius: 14px;
    margin-bottom: 1.4rem;
}}
.hero-title {{
    color: #F6F5F1;
    font-size: 2rem;
    font-weight: 600;
    margin: 0;
}}
.hero-sub {{
    color: #B9C2CE;
    font-size: 0.95rem;
    margin-top: 0.3rem;
}}
.metric-card {{
    background-color: {COLOR_CARD};
    border: 1px solid {COLOR_LINE};
    border-radius: 12px;
    padding: 1rem 1.2rem;
    box-shadow: 0 1px 3px rgba(27,36,48,0.06);
}}
.metric-label {{
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: {COLOR_MUTED};
    font-weight: 600;
}}
.metric-value {{
    font-family: 'Fraunces', serif;
    font-size: 1.9rem;
    color: {COLOR_INK};
    font-weight: 600;
    margin-top: 0.15rem;
}}
.section-label {{
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: {COLOR_MUTED};
    font-weight: 700;
    margin-bottom: 0.2rem;
}}
.data-note {{
    background-color: #FBF3E6;
    border-left: 3px solid {SEGMENT_COLORS['Promising']};
    padding: 0.7rem 1rem;
    border-radius: 6px;
    font-size: 0.87rem;
    color: {COLOR_INK};
}}
div[data-testid="stMetricValue"] {{
    font-family: 'Fraunces', serif;
}}
.stTabs [data-baseweb="tab-list"] {{
    gap: 4px;
}}
.stTabs [data-baseweb="tab"] {{
    background-color: {COLOR_CARD};
    border-radius: 8px 8px 0 0;
    padding: 0.5rem 1rem;
}}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

PLOTLY_TEMPLATE = go.layout.Template(
    layout=go.Layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=COLOR_INK, size=13),
        colorway=[COLOR_ACCENT, COLOR_MUTED, "#8C6A9E", "#4E9C87"],
        xaxis=dict(gridcolor=COLOR_LINE, zeroline=False),
        yaxis=dict(gridcolor=COLOR_LINE, zeroline=False),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
        margin=dict(t=40, l=10, r=10, b=10),
    )
)

DIMENSIONS = {
    "Age Group": "CustomerAgeGroup",
    "Gender": "CustomerGender",
    "Region": "CustomerRegion",
    "Retail Channel": "RetailChannel",
    "Product Category": "ProductCategoryCondensed",
}
def fmt_money(value) -> str:
    """Format a number as currency: $4,000 if whole, $4,000.50 if not.
    Always rounds to a maximum of 2 decimal places."""
    rounded = round(float(value), 2)
    if rounded == int(rounded):
        return f"${rounded:,.0f}"
    return f"${rounded:,.2f}"

# ----------------------------------------------------------------------------
# DATA LOADING & CLEANING
# ----------------------------------------------------------------------------
@st.cache_data
def load_data(path: str = "NR_dataset.xlsx") -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="data")

    # Drop fully-blank trailing rows (export artifacts with no idx at all)
    df = df.dropna(subset=["idx"]).copy()

    # One record (idx 97) has no segment label — excluded from segment-based
    # views since it can't be attributed to Decline/Stable/Growth/Promising.
    df = df.dropna(subset=["label"]).copy()

    # Tidy dtypes
    df["CustomerSatisfaction"] = df["CustomerSatisfaction"].astype(int)
    df["PurchaseAmount"] = df["PurchaseAmount"].astype(float)
    df["label"] = pd.Categorical(df["label"], categories=SEGMENT_ORDER, ordered=True)

    age_order = ["18-24", "25-34", "35-44", "45-54", "55+"]
    df["CustomerAgeGroup"] = pd.Categorical(
        df["CustomerAgeGroup"], categories=age_order, ordered=True
    )

    df = df.rename(columns={"label": "Segment"})
    return df


df_raw = load_data()

# ----------------------------------------------------------------------------
# SIDEBAR — FILTERS
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### NovaRetail")
    st.caption("Customer Intelligence Dashboard")
    st.divider()
    st.markdown("**Filters**")

    seg_pick = st.multiselect("Segment", SEGMENT_ORDER, default=SEGMENT_ORDER)
    age_pick = st.multiselect(
        "Age Group", list(df_raw["CustomerAgeGroup"].cat.categories),
        default=list(df_raw["CustomerAgeGroup"].cat.categories),
    )
    gender_pick = st.multiselect(
        "Gender", sorted(df_raw["CustomerGender"].unique()),
        default=sorted(df_raw["CustomerGender"].unique()),
    )
    region_pick = st.multiselect(
        "Region", sorted(df_raw["CustomerRegion"].unique()),
        default=sorted(df_raw["CustomerRegion"].unique()),
    )
    channel_pick = st.multiselect(
        "Retail Channel", sorted(df_raw["RetailChannel"].unique()),
        default=sorted(df_raw["RetailChannel"].unique()),
    )
    cat_pick = st.multiselect(
        "Product Category", sorted(df_raw["ProductCategoryCondensed"].unique()),
        default=sorted(df_raw["ProductCategoryCondensed"].unique()),
    )

    st.divider()
    st.caption(
        "Note: CustomerID / TransactionID are unreliable in this dataset "
        "(ID collisions across records) and are excluded from analysis. "
        "Each row is treated as an independent customer record."
    )

df = df_raw[
    df_raw["Segment"].isin(seg_pick)
    & df_raw["CustomerAgeGroup"].isin(age_pick)
    & df_raw["CustomerGender"].isin(gender_pick)
    & df_raw["CustomerRegion"].isin(region_pick)
    & df_raw["RetailChannel"].isin(channel_pick)
    & df_raw["ProductCategoryCondensed"].isin(cat_pick)
].copy()

# ----------------------------------------------------------------------------
# HERO
# ----------------------------------------------------------------------------
st.markdown(
    f"""
    <div class="hero-band">
        <p class="hero-title">NovaRetail Customer Intelligence</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if df.empty:
    st.warning("No records match the current filters. Adjust the sidebar to see data.")
    st.stop()

# ----------------------------------------------------------------------------
# KPI ROW
# ----------------------------------------------------------------------------
avg_purchase = df["PurchaseAmount"].mean()
avg_csat = df["CustomerSatisfaction"].mean()
pct_decline = (df["Segment"] == "Decline").mean() * 100

k1, k2, k3 = st.columns(3)
for col, label, value in zip(
    [k1, k2, k3],
    ["Avg. Purchase", "Avg. CSAT (1-5)", "% At Risk (Decline)"],
    [fmt_money(avg_purchase), f"{avg_csat:.2f}", f"{pct_decline:.1f}%"],
):
    col.markdown(
        f"""<div class="metric-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
            </div>""",
        unsafe_allow_html=True,
    )

st.write("")

# ----------------------------------------------------------------------------
# TABS
# ----------------------------------------------------------------------------
tab_overview, tab_revenue, tab_risk, tab_invest, tab_data = st.tabs(
    ["Overview", "Revenue Drivers", "At-Risk Segments", "Investment Focus", "Data Explorer"]
)

# ---- OVERVIEW --------------------------------------------------------------
with tab_overview:
    c1, c2 = st.columns([1, 1.3])

    with c1:
        st.markdown('<p class="section-label">Segment Mix</p>', unsafe_allow_html=True)
        seg_counts = df["Segment"].value_counts().reindex(SEGMENT_ORDER).fillna(0)
        fig = go.Figure(
            go.Pie(
                labels=seg_counts.index,
                values=seg_counts.values,
                hole=0,
                marker=dict(colors=[SEGMENT_COLORS[s] for s in seg_counts.index]),
                textinfo="label+percent",
            )
        )
        fig.update_layout(template=PLOTLY_TEMPLATE, showlegend=False, height=340)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown('<p class="section-label">Revenue by Segment</p>', unsafe_allow_html=True)
        rev_by_seg = (
            df.groupby("Segment", observed=True)["PurchaseAmount"]
            .agg(["sum", "mean", "count"])
            .reindex(SEGMENT_ORDER)
            .reset_index()
        )
        rev_by_seg["sum_label"] = rev_by_seg["sum"].apply(fmt_money)
        fig = px.bar(
            rev_by_seg, x="Segment", y="sum",
            color="Segment", color_discrete_map=SEGMENT_COLORS,
            labels={"sum": "Total Revenue ($)"}, text="sum_label",
        )
        fig.update_layout(template=PLOTLY_TEMPLATE, showlegend=False, height=340)
        st.plotly_chart(fig, use_container_width=True)

    st.caption("Select a segment to filter the table below:")
    if "selected_segment" not in st.session_state:
        st.session_state.selected_segment = None

    btn_pad_l, btn_area, btn_pad_r = st.columns([1, 3, 1])
    with btn_area:
        btn_cols = st.columns(len(SEGMENT_ORDER) + 1)
        with btn_cols[0]:
            if st.button(
                "All", use_container_width=True,
                type="primary" if st.session_state.selected_segment is None else "secondary",
            ):
                st.session_state.selected_segment = None
        for i, seg_name in enumerate(SEGMENT_ORDER, start=1):
            with btn_cols[i]:
                if st.button(
                    seg_name, use_container_width=True,
                    type="primary" if st.session_state.selected_segment == seg_name else "secondary",
                ):
                    st.session_state.selected_segment = seg_name

    clicked_segment = st.session_state.selected_segment
    if clicked_segment:
        snapshot_df = df[df["Segment"] == clicked_segment]
        snapshot_title = f'Segment Snapshot — "{clicked_segment}" only'
    else:
        snapshot_df = df
        snapshot_title = "Segment Snapshot — all segments"

    st.markdown(f'<p class="section-label">{snapshot_title}</p>', unsafe_allow_html=True)
    st.caption(f"{len(snapshot_df)} matching record(s).")
    st.dataframe(
        snapshot_df[
            ["Segment", "CustomerAgeGroup", "CustomerGender", "CustomerRegion",
             "RetailChannel", "ProductCategoryCondensed", "PurchaseAmount", "CustomerSatisfaction"]
        ],
        use_container_width=True, height=360,
        column_config={
            "PurchaseAmount": st.column_config.NumberColumn("Purchase Amount", format="$%.2f"),
        },
    )
# ---- REVENUE DRIVERS --------------------------------------------------------
with tab_revenue:
    st.markdown('<p class="section-label">Customer Segments that Generate the Most Revenue</p>', unsafe_allow_html=True)
    dim_label = st.radio(
        "Compare average purchase size and total revenue amount by:",
        list(DIMENSIONS.keys()), horizontal=True, key="rev_dim",
    )
    dim_col = DIMENSIONS[dim_label]

    grp = (
        df.groupby(dim_col, observed=True)["PurchaseAmount"]
        .agg(["mean", "sum", "count"])
        .reset_index()
    )
    if dim_col == "CustomerAgeGroup":
        grp = grp.sort_values(dim_col)
    else:
        grp = grp.sort_values("mean", ascending=False)
    grp.columns = [dim_label, "Avg Purchase", "Total Revenue", "Records"]
    grp["avg_label"] = grp["Avg Purchase"].apply(fmt_money)
    grp["total_label"] = grp["Total Revenue"].apply(fmt_money)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(
            grp, x=dim_label, y="Avg Purchase", text="avg_label",
            color="Avg Purchase", color_continuous_scale=["#DCE6F0", COLOR_ACCENT],
        )
        fig.update_layout(template=PLOTLY_TEMPLATE, height=380, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(
            grp, x=dim_label, y="Total Revenue", text="total_label",
            color="Total Revenue", color_continuous_scale=["#F0E6D2", SEGMENT_COLORS["Promising"]],
        )
        fig.update_layout(template=PLOTLY_TEMPLATE, height=380, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<p class="section-label">Revenue mix: segment × dimension</p>', unsafe_allow_html=True)
    pivot = (
        df.groupby([dim_col, "Segment"], observed=True)["PurchaseAmount"]
        .sum()
        .reset_index()
    )
    fig = px.bar(
        pivot, x=dim_col, y="PurchaseAmount", color="Segment",
        color_discrete_map=SEGMENT_COLORS, category_orders={"Segment": SEGMENT_ORDER},
        labels={"PurchaseAmount": "Total Revenue ($)", dim_col: dim_label},
    )
    fig.update_layout(template=PLOTLY_TEMPLATE, height=380, barmode="stack")
    st.plotly_chart(fig, use_container_width=True)

# ---- AT-RISK SEGMENTS -------------------------------------------------------
with tab_risk:
    st.markdown('<p class="section-label">Breakdown of Customers Labeled as in Decline</p>', unsafe_allow_html=True)
    dim_label_r = st.radio(
        "View Decline records by:", list(DIMENSIONS.keys()), horizontal=True, key="risk_dim",
    )
    dim_col_r = DIMENSIONS[dim_label_r]

    decline_df = df[df["Segment"] == "Decline"]
    total_by_dim = df.groupby(dim_col_r, observed=True).size()
    decline_by_dim = decline_df.groupby(dim_col_r, observed=True).size()
    risk_rate = (decline_by_dim / total_by_dim * 100).fillna(0).round(1)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(
            x=decline_by_dim.index.astype(str), y=decline_by_dim.values,
            labels={"x": dim_label_r, "y": "Decline Count"},
            color_discrete_sequence=[SEGMENT_COLORS["Decline"]],
        )
        fig.update_layout(template=PLOTLY_TEMPLATE, height=360, title="Decline count")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(
            x=risk_rate.index.astype(str), y=risk_rate.values,
            labels={"x": dim_label_r, "y": "% of group in Decline"},
            color_discrete_sequence=[SEGMENT_COLORS["Decline"]],
        )
        fig.update_layout(template=PLOTLY_TEMPLATE, height=360, title="Decline rate within group")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<p class="section-label">Lowest CSAT records (flagged)</p>', unsafe_allow_html=True)
    csat_threshold = st.slider("Flag records with CSAT at or below:", 1, 5, 2)
    flagged = df[df["CustomerSatisfaction"] <= csat_threshold].sort_values("CustomerSatisfaction")

    dim_label_c = st.radio(
        "Break down low-CSAT customers by:", list(DIMENSIONS.keys()), horizontal=True, key="csat_dim",
    )
    dim_col_c = DIMENSIONS[dim_label_c]

    total_by_dim_c = df.groupby(dim_col_c, observed=True).size()
    flagged_by_dim = flagged.groupby(dim_col_c, observed=True).size()
    flagged_rate = (flagged_by_dim / total_by_dim_c * 100).fillna(0).round(1)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(
            x=flagged_by_dim.index.astype(str), y=flagged_by_dim.values,
            labels={"x": dim_label_c, "y": "Low-CSAT Count"},
            color_discrete_sequence=[SEGMENT_COLORS["Decline"]],
        )
        fig.update_layout(template=PLOTLY_TEMPLATE, height=340, title="Low-CSAT count")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(
            x=flagged_rate.index.astype(str), y=flagged_rate.values,
            labels={"x": dim_label_c, "y": "% of group with low CSAT"},
            color_discrete_sequence=[SEGMENT_COLORS["Decline"]],
        )
        fig.update_layout(template=PLOTLY_TEMPLATE, height=340, title="Low-CSAT rate within group")
        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        flagged[
            ["Segment", "CustomerSatisfaction", "PurchaseAmount", "CustomerAgeGroup",
             "CustomerGender", "CustomerRegion", "RetailChannel", "ProductCategoryCondensed"]
        ],
        use_container_width=True, height=300,
        column_config={
            "PurchaseAmount": st.column_config.NumberColumn("Purchase Amount", format="$%.2f"),
        },
    )
    st.caption(f"{len(flagged)} record(s) at or below a CSAT score of {csat_threshold}.")

# ---- INVESTMENT FOCUS -------------------------------------------------------
with tab_invest:
    st.markdown(
        '<p class="section-label">High CSAT + high purchase amount = investment priority</p>',
        unsafe_allow_html=True,
    )
    dim_label_i = st.radio(
        "Group opportunity view by:", list(DIMENSIONS.keys()), horizontal=True, key="invest_dim",
    )
    dim_col_i = DIMENSIONS[dim_label_i]

    opp = df[df["Segment"].isin(["Stable", "Growth", "Promising"])]
    opp_grp = (
        opp.groupby([dim_col_i, "Segment"], observed=True)
        .agg(Avg_CSAT=("CustomerSatisfaction", "mean"),
             Avg_Purchase=("PurchaseAmount", "mean"),
             Records=("PurchaseAmount", "count"))
        .reset_index()
    )

    fig = px.scatter(
        opp_grp, x="Avg_CSAT", y="Avg_Purchase", size="Records", color="Segment",
        color_discrete_map=SEGMENT_COLORS, category_orders={"Segment": SEGMENT_ORDER},
        hover_name=dim_col_i, size_max=40,
        labels={"Avg_CSAT": "Average CSAT", "Avg_Purchase": "Average Purchase ($)"},
        text=dim_col_i,
    )
    fig.update_traces(textposition="top center")
    fig.update_layout(template=PLOTLY_TEMPLATE, height=460)
    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        "Bubbles further right and higher up — with larger size — mark "
        "groups combining high satisfaction, high spend, and enough volume "
        "to justify investment."
    )

    st.markdown('<p class="section-label">Opportunity table</p>', unsafe_allow_html=True)
    min_csat = st.slider("Minimum average CSAT", 1.0, 5.0, 4.0, 0.1)
    min_purchase = st.slider("Minimum average purchase ($)", 0, 1000, 100, 10)
    qualifying = opp_grp[
        (opp_grp["Avg_CSAT"] >= min_csat) & (opp_grp["Avg_Purchase"] >= min_purchase)
    ].sort_values("Avg_Purchase", ascending=False)
    st.dataframe(
        qualifying.round(2),
        use_container_width=True,
        column_config={
            "Avg_Purchase": st.column_config.NumberColumn("Avg Purchase", format="$%.2f"),
        },
    )

# ---- DATA EXPLORER -----------------------------------------------------------
with tab_data:
    st.markdown('<p class="section-label">Filtered records</p>', unsafe_allow_html=True)
    display_cols = [
        "Segment", "CustomerAgeGroup", "CustomerGender", "CustomerRegion",
        "RetailChannel", "ProductCategoryCondensed", "PurchaseAmount", "CustomerSatisfaction",
    ]
    st.dataframe(
        df[display_cols], use_container_width=True, height=420,
        column_config={
            "PurchaseAmount": st.column_config.NumberColumn("Purchase Amount", format="$%.2f"),
        },
    )
    st.download_button(
        "Download filtered data (CSV)",
        df[display_cols].to_csv(index=False).encode("utf-8"),
        file_name="novaretail_filtered.csv",
        mime="text/csv",
    )
    st.markdown(
        """<div class="data-note">
        <strong>Data notes:</strong> Two blank trailing rows and one record
        with a missing segment label were removed during cleaning.
        CustomerID and TransactionID are not reliable unique identifiers in
        this dataset (collisions were found across unrelated records), so
        each row is analyzed as an independent customer record rather than
        deduplicated to a unique-customer level.
        </div>""",
        unsafe_allow_html=True,
    )
