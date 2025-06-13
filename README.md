# - Financial Risk Analytics Toolkit (Excel/VBA)

This project simulates and visualizes key financial risk indicators using an Excel-based interface, designed to support asset-liability management and liquidity risk monitoring across multiple entities. The tool is VBA-powered for automation, enabling monthly data processing, risk flagging, and dynamic report generation.

## 📌 Key Features

- **Duration Gap Analysis**  
  Calculates monthly and quarterly duration gaps by country using a modified duration formula. Highlights duration mismatches and interest rate sensitivity.

- **Liquidity Ratio Tracking**  
  Monitors trends in liquidity ratios from January to June 2025, supporting early detection of short-term funding misalignments.

- **FX Risk Distribution**  
  Aggregates foreign exchange exposure (EUR, USD, GBP) across time periods to visualize currency risk concentration.

- **Interactive Dashboards**  
  Includes pivot-table-driven dashboards with slicers for multi-dimensional filtering by country, quarter, currency and risk type.

- **Automated Risk Flagging**  
  Uses VBA scripts to detect outliers and high-risk entries, such as low liquidity or excessive duration gaps.

- **Scenario-Ready Design**  
  Changes in asset/liability durations or liquidity assumptions can be simulated to observe their impact on risk metrics.

## 🔧 Technologies Used

- Microsoft Excel (Pivot Tables, Charts)
- VBA for automation
- Data validation & conditional formatting

## 📂 File Structure

- `financial_risk.xlsm` — Main macro-enabled file with all functionalities
- `Simulated Risk Data` — Input sheet for monthly risk observations

- `Dashboard` — Visual summaries of risk metrics
- `Quarterly Duration Gap`, `Duration Gap Trend`, `Liquidity Ratio`, `FX Risk Distribution` — Processed analytical outputs

## 📎 Getting Started

1. Open the `.xlsm` file in Excel with macros enabled.
2. Navigate to the `Simulated Risk Data` sheet to input or modify raw data.
3. Use the Dashboard or other output sheets to explore dynamic visualizations.
4. Run VBA macros (if needed) to refresh risk flags and summaries.

## ✅ Use Case

This tool was built to simulate monitoring key metrics, automating reports, and assisting in ALM-related decisions. It serves both as a learning exercise and a practical tool for portfolio or risk analysts.

## 📄 License

This project is released under the MIT License.

---

*Author: Qiushi Guo — [LinkedIn](https://www.linkedin.com/in/qiushi-guo/)*
