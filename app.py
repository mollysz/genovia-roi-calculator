import streamlit as st
import pandas as pd
from pathlib import Path

# =========================================================
# 1. LOAD CONFIG FROM CSV FILES
# =========================================================
TIERS_PATH = Path("data/tiers.csv")
SHIPPING_PATH = Path("data/shipping.csv")
GLOBAL_PATH = Path("data/global_settings.csv")


@st.cache_data
def load_config():
    tiers_df = pd.read_csv(TIERS_PATH)
    shipping_df = pd.read_csv(SHIPPING_PATH)
    global_df = pd.read_csv(GLOBAL_PATH)

    return tiers_df, shipping_df, global_df


tiers_df, shipping_df, global_df = load_config()

# global settings
GLOBAL_SETTINGS = {row["key"]: row["value"] for _, row in global_df.iterrows()}
CURRENCY = str(GLOBAL_SETTINGS.get("currency_symbol", "$"))
DEFAULT_MIN_CASES_GLOBAL = int(GLOBAL_SETTINGS.get("default_min_cases_global", 1))
DEFAULT_MAX_CASES_GLOBAL = int(GLOBAL_SETTINGS.get("default_max_cases_global", 500))

# build TIERS dict
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

TX_PER_CASE_DEFAULT = list(TIERS_BASE.values())[0]["tx_per_case"]

# shipping dict
SHIPPING_BASE = {
    row["shipping_name"]: float(row["shipping_cost"])
    for _, row in shipping_df.iterrows()
}

# =========================================================
# 2. HELPERS
# =========================================================
def calc_roi(tier, num_cases, price_per_tx, extra_cost_per_tx, shipping_cost):
    case_price = tier["case_price"]
    cost_per_tx_product = tier["cost_per_tx"]
    tx_per_case = tier.get("tx_per_case", TX_PER_CASE_DEFAULT)

    total_cases = num_cases
    total_txs = total_cases * tx_per_case

    product_cost = total_cases * case_price
    total_cost = product_cost + shipping_cost

    total_cost_per_tx = cost_per_tx_product + extra_cost_per_tx
    revenue_per_tx = price_per_tx
    profit_per_tx = revenue_per_tx - total_cost_per_tx

    total_revenue = revenue_per_tx * total_txs
    total_profit = profit_per_tx * total_txs

    margin_pct = (total_profit / total_revenue * 100) if total_revenue else 0
    roi_pct = (total_profit / total_cost * 100) if total_cost else 0

    breakeven_txs = total_cost / profit_per_tx if profit_per_tx > 0 else None

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


# =========================================================
# 3. UI LAYOUT
# =========================================================
st.set_page_config(page_title="Genovia ROI Calculator", page_icon="💧", layout="centered")

st.title("Genovia™ ROI Calculator")
st.caption("All pricing and costs load from CSV files in your GitHub repo.")

# ------------------ SIDEBAR OVERRIDES ------------------
with st.sidebar:
    st.header("Advanced Overrides")
    use_overrides = st.checkbox("Enable manual overrides", value=False)

    tiers_runtime = {k: v.copy() for k, v in TIERS_BASE.items()}
    shipping_runtime = SHIPPING_BASE.copy()

    if use_overrides:
        for tier_name, tier in tiers_runtime.items():
            with st.expander(f"{tier_name} Tier"):
                tier["case_price"] = st.number_input(
                    f"{tier_name} case price", 0.0, step=10.0, value=tier["case_price"]
                )
                tier["cost_per_tx"] = st.number_input(
                    f"{tier_name} cost per treatment", 0.0, step=1.0, value=tier["cost_per_tx"]
                )
                tier["tx_per_case"] = st.number_input(
                    f"{tier_name} treatments per case",
                    min_value=1,
                    step=1,
                    value=tier["tx_per_case"],
                )

# ------------------ MAIN INPUTS ------------------
st.subheader("Step 1 — Clinic Inputs")

tier_choice = st.selectbox("Choose a Genovia tier", list(tiers_runtime.keys()))
tier_selected = tiers_runtime[tier_choice]

num_cases = st.number_input(
    "Number of cases",
    min_value=int(tier_selected["default_min_cases"]),
    max_value=int(tier_selected["default_max_cases"]),
    value=int(tier_selected["default_min_cases"]),
)

price_per_tx = st.number_input(
    "Clinic price per treatment ($)",
    value=float(tier_selected["default_clinic_price_per_tx"]),
    min_value=0.0,
    step=50.0,
)

extra_cost_per_tx = st.number_input(
    "Other per-treatment cost (staff, room, etc.)",
    value=float(tier_selected["default_extra_cost_per_tx"]),
    min_value=0.0,
    step=10.0,
)

shipping_name = st.selectbox("Shipping option", list(shipping_runtime.keys()))
shipping_cost = shipping_runtime[shipping_name]

st.markdown("---")

# ------------------ ROI SUMMARY ------------------
st.subheader("Step 2 — ROI Summary")

results = calc_roi(
    tier=tier_selected,
    num_cases=num_cases,
    price_per_tx=price_per_tx,
    extra_cost_per_tx=extra_cost_per_tx,
    shipping_cost=shipping_cost,
)

# metrics
col1, col2, col3 = st.columns(3)
col1.metric("Total Revenue", fc(results["total_revenue"]))
col2.metric("Total Cost", fc(results["total_cost"]))
col3.metric("Total Profit", fc(results["total_profit"]))

# charts
summary_df = pd.DataFrame(
    {
        "Metric": ["Total Revenue", "Total Cost", "Total Profit"],
        "Amount": [
            results["total_revenue"],
            results["total_cost"],
            results["total_profit"],
        ],
    }
).set_index("Metric")

st.bar_chart(summary_df)

# ------------------ TIER COMPARISON ------------------
# st.subheader("Compare Tiers at Same Clinic Price")

# if st.checkbox("Show comparison table"):
#     comparison = []
#     for tname, info in tiers_runtime.items():
#         r = calc_roi(
#             tier=info,
#             num_cases=num_cases,
#             price_per_tx=price_per_tx,
#             extra_cost_per_tx=extra_cost_per_tx,
#             shipping_cost=shipping_cost,
#         )
#         comparison.append(
#             {
#                 "Tier": tname,
#                 "Cost per treatment": r["cost_per_tx_product"],
#                 "Total Profit": r["total_profit"],
#                 "ROI %": r["roi_pct"],
#             }
#         )

#     comp_df = pd.DataFrame(comparison)
#     st.table(comp_df)

#     # Profit chart
#     st.markdown("#### Profit by Tier")
#     st.bar_chart(comp_df.set_index("Tier")["Total Profit"])

#     # ROI chart
#     st.markdown("#### ROI by Tier")
#     st.bar_chart(comp_df.set_index("Tier")["ROI %"])

st.subheader("Compare Tiers at Same Clinic Price")

# Always compute comparison — no checkbox needed
comparison = []
for tname, info in tiers_runtime.items():
    r = calc_roi(
        tier=info,
        num_cases=num_cases,
        price_per_tx=price_per_tx,
        extra_cost_per_tx=extra_cost_per_tx,
        shipping_cost=shipping_cost,
    )
    comparison.append(
        {
            "Tier": tname,
            "Cost per treatment": r["cost_per_tx_product"],
            "Total Profit": r["total_profit"],
            "ROI %": r["roi_pct"],
        }
    )

comp_df = pd.DataFrame(comparison)

st.table(comp_df)

# Profit chart
st.markdown("#### Profit by Tier")
st.bar_chart(comp_df.set_index("Tier")["Total Profit"])

# ROI chart
st.markdown("#### ROI by Tier")
st.bar_chart(comp_df.set_index("Tier")["ROI %"])
