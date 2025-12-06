import streamlit as st
import pandas as pd
from pathlib import Path
import io
from docx import Document

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


def build_word_report(
    tier_choice,
    tier_selected,
    num_cases,
    price_per_tx,
    extra_cost_per_tx,
    shipping_name,
    results,
    comp_df,
):
    """Create a Word document (as BytesIO) summarizing the scenario + comparison."""
    doc = Document()

    # Title
    doc.add_heading("Genovia ROI Report", level=1)
    doc.add_paragraph(
        "This report summarizes the ROI scenario generated from the Genovia ROI Calculator. "
        "It is intended for internal review by management and sales leadership."
    )

    # Section 1 – Scenario Overview
    doc.add_heading("1. Scenario Overview", level=2)
    p = doc.add_paragraph()
    p.add_run("Selected tier: ").bold = True
    p.add_run(tier_choice)
    doc.add_paragraph(f"Number of cases: {num_cases}")
    doc.add_paragraph(f"Clinic price per treatment: {fc1(price_per_tx)}")
    doc.add_paragraph(f"Other cost per treatment: {fc1(extra_cost_per_tx)}")
    doc.add_paragraph(f"Shipping option: {shipping_name}")

    # Section 2 – Key Financial Metrics
    doc.add_heading("2. Key Financial Metrics", level=2)
    metrics = [
        ("Total treatments", f"{results['total_txs']:,}"),
        ("Total revenue", fc(results["total_revenue"])),
        ("Total cost (product + shipping)", fc(results["total_cost"])),
        ("Total profit", fc(results["total_profit"])),
        ("Profit per treatment", fc1(results["profit_per_tx"])),
        ("Profit margin", f"{results['margin_pct']:.1f}%"),
        ("ROI on order", f"{results['roi_pct']:.1f}%"),
    ]
    table = doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = "Metric"
    hdr[1].text = "Value"
    for name, value in metrics:
        row = table.add_row().cells
        row[0].text = name
        row[1].text = value

    # Section 3 – Interpretation
    doc.add_heading("3. Interpretation of Results", level=2)
    doc.add_paragraph(
        f"At the {tier_choice} tier, with {num_cases} cases ordered, the clinic generates "
        f"{fc(results['total_revenue'])} in revenue and {fc(results['total_profit'])} in profit, "
        f"corresponding to a margin of {results['margin_pct']:.1f}% and an ROI of {results['roi_pct']:.1f}%."
    )
    if results["breakeven_txs"]:
        doc.add_paragraph(
            f"The estimated break-even point is {results['breakeven_txs']:.0f} treatments."
        )

    # Section 4 – Tier Comparison
    doc.add_heading("4. Tier Comparison at Same Clinic Price", level=2)
    comp_table = doc.add_table(rows=1, cols=4)
    ch = comp_table.rows[0].cells
    ch[0].text = "Tier"
    ch[1].text = "Cost per Treatment (Genovia)"
    ch[2].text = "Total Profit"
    ch[3].text = "ROI %"

    for _, row in comp_df.iterrows():
        r = comp_table.add_row().cells
        r[0].text = str(row["Tier"])
        r[1].text = fc1(row["Cost per treatment"])
        r[2].text = fc(row["Total Profit"])
        r[3].text = f"{row['ROI %']:.1f}%"

    # highlight best tier
    best_row = comp_df.loc[comp_df["ROI %"].idxmax()]
    best_tier = best_row["Tier"]
    best_roi = best_row["ROI %"]
    doc.add_paragraph(
        f"Across tiers, {best_tier} delivers the highest ROI at approximately {best_roi:.1f}% "
        "under the same clinic pricing and volume assumptions."
    )

    # Section 5 – Recommendations
    doc.add_heading("5. Recommendations", level=2)
    recs = [
        "Highlight Gold and Diamond tiers in sales conversations as they deliver higher ROI versus Standard.",
        "Use this ROI model live during demos to show how changes in price and volume impact profit.",
        "Share tailored ROI reports with top clinic prospects to support pricing and volume commitments.",
    ]
    for r in recs:
        doc.add_paragraph(f"• {r}")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# =========================================================
# 3. UI LAYOUT
# =========================================================
st.set_page_config(page_title="Genovia ROI Calculator", page_icon="💧", layout="centered")

st.title("Genovia™ ROI Calculator")
st.caption("All pricing and costs load from CSV files in your GitHub repo.")

# ------------------ TIER SETTINGS & PRICING INPUTS ------------------
st.markdown("### Tier Settings & Pricing Inputs")
st.caption(
    "Adjust price and cost parameters for each Genovia tier to model different business scenarios. "
    "These settings affect only this session and do not change your master CSV files."
)

tiers_runtime = {k: v.copy() for k, v in TIERS_BASE.items()}
shipping_runtime = SHIPPING_BASE.copy()

settings_col1, settings_col2 = st.columns(2)

with settings_col1:
    st.markdown("**Tier Pricing Parameters**")
    for tier_name, tier in tiers_runtime.items():
        with st.expander(f"{tier_name} — Pricing Settings", expanded=False):
            tier["case_price"] = st.number_input(
                f"{tier_name} · Case Price",
                min_value=0.0,
                step=10.0,
                value=tier["case_price"],
                key=f"case_price_{tier_name}",
            )
            tier["cost_per_tx"] = st.number_input(
                f"{tier_name} · Cost per Treatment (Genovia)",
                min_value=0.0,
                step=1.0,
                value=tier["cost_per_tx"],
                key=f"cost_per_tx_{tier_name}",
            )
            tier["tx_per_case"] = st.number_input(
                f"{tier_name} · Treatments per Case",
                min_value=1,
                step=1,
                value=tier["tx_per_case"],
                key=f"tx_per_case_{tier_name}",
            )

with settings_col2:
    st.markdown("**Shipping Cost Assumptions**")
    for ship_name, cost in list(shipping_runtime.items()):
        shipping_runtime[ship_name] = st.number_input(
            f"{ship_name}",
            min_value=0.0,
            step=5.0,
            value=cost,
            key=f"shipping_{ship_name}",
        )

st.markdown("---")

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
    num_cases=int(num_cases),
    price_per_tx=price_per_tx,
    extra_cost_per_tx=extra_cost_per_tx,
    shipping_cost=shipping_cost,
)

col1, col2, col3 = st.columns(3)
col1.metric("Total Revenue", fc(results["total_revenue"]))
col2.metric("Total Cost", fc(results["total_cost"]))
col3.metric("Total Profit", fc(results["total_profit"]))

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

# CSV download for scenario
summary_export_df = pd.DataFrame(
    [{
        "Tier": tier_choice,
        "Number of cases": num_cases,
        "Clinic price per treatment": price_per_tx,
        "Other cost per treatment": extra_cost_per_tx,
        "Shipping option": shipping_name,
        "Total treatments": results["total_txs"],
        "Total revenue": results["total_revenue"],
        "Total cost": results["total_cost"],
        "Total profit": results["total_profit"],
        "Profit per treatment": results["profit_per_tx"],
        "Margin %": results["margin_pct"],
        "ROI %": results["roi_pct"],
    }]
)
st.download_button(
    label="⬇️ Download current scenario (CSV)",
    data=summary_export_df.to_csv(index=False),
    file_name="genovia_roi_scenario.csv",
    mime="text/csv",
)

st.markdown("---")

# ------------------ TIER COMPARISON ------------------
st.subheader("Compare Tiers at Same Clinic Price")

comparison = []
for tname, info in tiers_runtime.items():
    r = calc_roi(
        tier=info,
        num_cases=int(num_cases),
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

st.markdown("#### Profit by Tier")
st.bar_chart(comp_df.set_index("Tier")["Total Profit"])

st.markdown("#### ROI by Tier")
st.bar_chart(comp_df.set_index("Tier")["ROI %"])

# CSV download for comparison
st.download_button(
    label="⬇️ Download tier comparison (CSV)",
    data=comp_df.to_csv(index=False),
    file_name="genovia_tier_comparison.csv",
    mime="text/csv",
)

# ------------------ WORD REPORT DOWNLOAD ------------------
report_buffer = build_word_report(
    tier_choice=tier_choice,
    tier_selected=tier_selected,
    num_cases=int(num_cases),
    price_per_tx=price_per_tx,
    extra_cost_per_tx=extra_cost_per_tx,
    shipping_name=shipping_name,
    results=results,
    comp_df=comp_df,
)

st.markdown("---")
st.download_button(
    label="⬇️ Download Word report (management summary)",
    data=report_buffer,
    file_name="Genovia_ROI_Report.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)
