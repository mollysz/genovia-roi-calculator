import streamlit as st
import pandas as pd
from pathlib import Path

# ------------------------------
# Load configuration from Excel
# ------------------------------
CONFIG_PATH = Path("data/genovia_config.xlsx")

@st.cache_data
def load_config(path: Path):
    xls = pd.ExcelFile(path)
    tiers_df = pd.read_excel(xls, sheet_name="tiers")
    shipping_df = pd.read_excel(xls, sheet_name="shipping")
    global_df = pd.read_excel(xls, sheet_name="global_settings")

    # Basic validation
    required_tier_cols = {
        "tier_name", "description", "case_price", "cost_per_tx",
        "savings_vs_standard_pct", "tx_per_case",
        "default_clinic_price_per_tx", "default_extra_cost_per_tx",
        "default_min_cases", "default_max_cases",
    }
    missing_tier = required_tier_cols - set(tiers_df.columns)
    if missing_tier:
        raise ValueError(f"Missing columns in tiers sheet: {missing_tier}")

    required_ship_cols = {"shipping_name", "shipping_cost"}
    missing_ship = required_ship_cols - set(shipping_df.columns)
    if missing_ship:
        raise ValueError(f"Missing columns in shipping sheet: {missing_ship}")

    return tiers_df, shipping_df, global_df

tiers_df, shipping_df, global_df = load_config(CONFIG_PATH)

# Convert global settings to dict
GLOBAL_SETTINGS = {
    row["key"]: row["value"] for _, row in global_df.iterrows()
}
CURRENCY = str(GLOBAL_SETTINGS.get("currency_symbol", "$"))

# Fallbacks for case ranges
DEFAULT_MIN_CASES_GLOBAL = int(GLOBAL_SETTINGS.get("default_min_cases_global", 1))
DEFAULT_MAX_CASES_GLOBAL = int(GLOBAL_SETTINGS.get("default_max_cases_global", 500))

# Build TIERS structure from Excel
TIERS_BASE = {}
for _, row in tiers_df.iterrows():
    TIERS_BASE[row["tier_name"]] = {
        "description": row["description"],
        "case_price": float(row["case_price"]),
        "cost_per_tx": float(row["cost_per_tx"]),
        "savings_vs_standard_pct": float(row["savings_vs_standard_pct"]),
        "tx_per_case": int(row["tx_per_case"]),
        "default_clinic_price_per_tx": float(row["default_clinic_price_per_tx"]),
        "default_extra_cost_per_tx": float(row["default_extra_cost_per_tx"]),
        "default_min_cases": int(row["default_min_cases"]),
        "default_max_cases": int(row["default_max_cases"]),
    }

# Assume same tx_per_case for all tiers (can be changed per tier if needed)
TX_PER_CASE = list(TIERS_BASE.values())[0]["tx_per_case"]

# Build shipping dict from Excel
SHIPPING_BASE = {
    row["shipping_name"]: float(row["shipping_cost"])
    for _, row in shipping_df.iterrows()
}


# ------------------------------
# Helper functions
# ------------------------------
def calc_roi(
    tier: dict,
    num_cases: int,
    price_per_tx: float,
    extra_cost_per_tx: float,
    shipping_cost: float,
):
    """Return a dict with all ROI metrics for a given tier and order size."""
    case_price = tier["case_price"]
    cost_per_tx_product = tier["cost_per_tx"]

    # Core quantities
    total_cases = num_cases
    total_txs = total_cases * tier["tx_per_case"]

    # Costs
    product_cost = total_cases * case_price
    total_cost = product_cost + shipping_cost

    # Per-treatment economics
    total_cost_per_tx = cost_per_tx_product + extra_cost_per_tx
    revenue_per_tx = price_per_tx
    profit_per_tx = revenue_per_tx - total_cost_per_tx

    # Revenue & profit
    total_revenue = revenue_per_tx * total_txs
    total_profit = profit_per_tx * total_txs

    # Margins & ROI
    margin_pct = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
    roi_pct = (total_profit / total_cost * 100) if total_cost > 0 else 0

    # Break-even
    if profit_per_tx > 0:
        breakeven_txs = total_cost / profit_per_tx
    else:
        breakeven_txs = None

    return {
        "total_cases": total_cases,
        "total_txs": total_txs,
        "product_cost": product_cost,
        "shipping_cost": shipping_cost,
        "total_cost": total_cost,
        "revenue_per_tx": revenue_per_tx,
        "cost_per_tx_product": cost_per_tx_product,
        "extra_cost_per_tx": extra_cost_per_tx,
        "total_cost_per_tx": total_cost_per_tx,
        "profit_per_tx": profit_per_tx,
        "total_revenue": total_revenue,
        "total_profit": total_profit,
        "margin_pct": margin_pct,
        "roi_pct": roi_pct,
        "breakeven_txs": breakeven_txs,
    }


def fc(x):
    return f"{CURRENCY}{x:,.0f}"


def fc1(x):
    return f"{CURRENCY}{x:,.1f}"


# ------------------------------
# Streamlit UI
# ------------------------------
st.set_page_config(
    page_title="Genovia ROI Calculator",
    page_icon="💧",
    layout="centered",
)

st.title("Genovia™ ROI Calculator")
st.caption("All pricing and costs are loaded from Excel. You can override values for this session inside the app.")

# ---- Advanced override controls ----
with st.sidebar:
    st.header("Advanced Controls")
    use_overrides = st.checkbox(
        "Enable manual overrides (session only)", value=False,
        help="Turn this on to adjust pricing and shipping in the app without editing Excel."
    )

    # Start with base dicts
    tiers_runtime = {k: v.copy() for k, v in TIERS_BASE.items()}
    shipping_runtime = SHIPPING_BASE.copy()

    if use_overrides:
        st.subheader("Tier pricing overrides")
        for tier_name, tier in tiers_runtime.items():
            with st.expander(f"{tier_name} tier settings", expanded=False):
                tier["case_price"] = st.number_input(
                    f"{tier_name} case price",
                    min_value=0.0,
                    value=tier["case_price"],
                    step=10.0,
                    key=f"case_price_{tier_name}",
                )
                tier["cost_per_tx"] = st.number_input(
                    f"{tier_name} cost per treatment (product)",
                    min_value=0.0,
                    value=tier["cost_per_tx"],
                    step=1.0,
                    key=f"cost_per_tx_{tier_name}",
                )
                tier["tx_per_case"] = st.number_input(
                    f"{tier_name} treatments per case",
                    min_value=1,
                    value=tier["tx_per_case"],
                    step=1,
                    key=f"tx_per_case_{tier_name}",
                )
                tier["default_clinic_price_per_tx"] = st.number_input(
                    f"{tier_name} default clinic price per treatment",
                    min_value=0.0,
                    value=tier["default_clinic_price_per_tx"],
                    step=50.0,
                    key=f"default_clinic_price_{tier_name}",
                )
                tier["default_extra_cost_per_tx"] = st.number_input(
                    f"{tier_name} default other cost per treatment",
                    min_value=0.0,
                    value=tier["default_extra_cost_per_tx"],
                    step=10.0,
                    key=f"default_extra_cost_{tier_name}",
                )

        st.subheader("Shipping overrides")
        for ship_name, cost in list(shipping_runtime.items()):
            shipping_runtime[ship_name] = st.number_input(
                f"Shipping cost – {ship_name}",
                min_value=0.0,
                value=cost,
                step=5.0,
                key=f"shipping_{ship_name}",
            )
    else:
        tiers_runtime = TIERS_BASE
        shipping_runtime = SHIPPING_BASE

# ---- Main controls ----
st.markdown("### Step 1 – Clinic pricing & volume")

tier_names = list(tiers_runtime.keys())
default_tier = tier_names[0]

col1, col2 = st.columns(2)
with col1:
    tier_choice = st.selectbox(
        "Genovia pricing tier",
        tier_names,
        help="Choose the tier you are offering this clinic."
    )

tier_selected = tiers_runtime[tier_choice]

# Determine cases range from tier / global
min_cases = tier_selected.get("default_min_cases", DEFAULT_MIN_CASES_GLOBAL)
max_cases = tier_selected.get("default_max_cases", DEFAULT_MAX_CASES_GLOBAL)

with col2:
    num_cases = st.number_input(
        "Number of cases for this order",
        min_value=min_cases,
        max_value=max_cases,
        value=min_cases,
        step=1,
    )

st.markdown("---")

col3, col4 = st.columns(2)
with col3:
    price_per_tx = st.number_input(
        "Clinic price charged per treatment",
        min_value=0.0,
        value=tier_selected["default_clinic_price_per_tx"],
        step=50.0,
        help="What the MedSpa plans to charge patients per treatment."
    )
with col4:
    extra_cost_per_tx = st.number_input(
        "Other cost per treatment",
        min_value=0.0,
        value=tier_selected["default_extra_cost_per_tx"],
        step=10.0,
        help="Staff time, room cost, numbing, etc. Excluding Genovia product cost."
    )

shipping_name = st.selectbox(
    "Shipping option",
    list(shipping_runtime.keys()),
    index=0
)
shipping_cost = shipping_runtime[shipping_name]

st.markdown("---")

st.markdown("### Step 2 – ROI summary")

results = calc_roi(
    tier=tier_selected,
    num_cases=int(num_cases),
    price_per_tx=price_per_tx,
    extra_cost_per_tx=extra_cost_per_tx,
    shipping_cost=shipping_cost,
)

st.subheader(f"{tier_choice} Tier Overview")
st.write(tier_selected["description"])
st.write(
    f"- **Genovia cost per case:** {fc1(tier_selected['case_price'])}  \n"
    f"- **Genovia cost per treatment:** {fc1(tier_selected['cost_per_tx'])}  \n"
    f"- **Treatments per case:** {tier_selected['tx_per_case']}"
)
if tier_selected.get("savings_vs_standard_pct", 0):
    st.write(f"- **Savings vs Standard:** {tier_selected['savings_vs_standard_pct']}%")

# Key metrics in cards
m1, m2, m3 = st.columns(3)
m1.metric("Total cases", f"{results['total_cases']}")
m2.metric("Total treatments", f"{results['total_txs']}")
m3.metric("Total product cost", fc(results["product_cost"]))

m4, m5, m6 = st.columns(3)
m4.metric("Total revenue", fc(results["total_revenue"]))
m5.metric("Total profit", fc(results["total_profit"]))
m6.metric("ROI", f"{results['roi_pct']:.1f}%")

m7, m8, m9 = st.columns(3)
m7.metric("Profit per treatment", fc1(results["profit_per_tx"]))
m8.metric("Margin", f"{results['margin_pct']:.1f}%")
if results["breakeven_txs"] is not None:
    m9.metric("Break-even treatments", f"{results['breakeven_txs']:.0f}")
else:
    m9.metric("Break-even treatments", "Not reachable (price too low)")

st.markdown("### Detailed breakdown")

detail_rows = {
    "Genovia cost per treatment": fc1(results["cost_per_tx_product"]),
    "Other cost per treatment": fc1(results["extra_cost_per_tx"]),
    "Total cost per treatment": fc1(results["total_cost_per_tx"]),
    "Price charged per treatment": fc1(results["revenue_per_tx"]),
    "Product cost (Genovia only)": fc(results["product_cost"]),
    "Shipping cost": fc(results["shipping_cost"]),
    "Total cost (product + shipping)": fc(results["total_cost"]),
    "Total revenue": fc(results["total_revenue"]),
    "Total profit": fc(results["total_profit"]),
    "Profit margin": f"{results['margin_pct']:.1f}%",
    "ROI on order": f"{results['roi_pct']:.1f}%",
}

detail_df = pd.DataFrame(
    {"Metric": list(detail_rows.keys()), "Value": list(detail_rows.values())}
)
st.table(detail_df)

st.markdown("### Optional – Compare tiers at the same clinic price")

if st.checkbox("Show tier comparison table"):
    comparison = []
    for t_name, t_info in tiers_runtime.items():
        r = calc_roi(
            tier=t_info,
            num_cases=int(num_cases),
            price_per_tx=price_per_tx,
            extra_cost_per_tx=extra_cost_per_tx,
            shipping_cost=shipping_cost,
        )
        comparison.append(
            {
                "Tier": t_name,
                "Cost per treatment (Genovia)": r["cost_per_tx_product"],
                "Total profit": r["total_profit"],
                "Profit per treatment": r["profit_per_tx"],
                "Margin %": r["margin_pct"],
                "ROI %": r["roi_pct"],
            }
        )
    comp_df = pd.DataFrame(comparison)
    comp_df_display = comp_df.copy()
    comp_df_display["Cost per treatment (Genovia)"] = comp_df_display[
        "Cost per treatment (Genovia)"
    ].map(fc1)
    comp_df_display["Total profit"] = comp_df_display["Total profit"].map(fc)
    comp_df_display["Profit per treatment"] = comp_df_display[
        "Profit per treatment"
    ].map(fc1)
    comp_df_display["Margin %"] = comp_df_display["Margin %"].map(
        lambda x: f"{x:.1f}%"
    )
    comp_df_display["ROI %"] = comp_df_display["ROI %"].map(lambda x: f"{x:.1f}%")
    st.table(comp_df_display)
