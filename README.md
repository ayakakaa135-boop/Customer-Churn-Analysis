# 📉 Corporate Customer Churn Intelligence System

[![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://www.python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.30-FF4B4B.svg)](https://streamlit.io)
[![License](https://img.shields.io/badge/License-Corporate_Portfolio-green.svg)](https://opensource.org/licenses/MIT)

## 📌 Executive Summary
Customer churn is one of the most critical threats to profitability in the telecom industry. This project delivers an **end-to-end AI-powered solution** to identify, analyze, and retain high-value customers.

By combining **Machine Learning (Logistic Regression)** with an **Advanced Analytical Data Model (Star Schema)**, this system provides actionable insights through a premium interactive dashboard.

---

## 🛠️ Tech Stack & Architecture
- **Data Engineering:** Star Schema Implementation (Fact & Dimension Tables in Excel/Python).
- **Analytics Engine:** Scikit-Learn for Predictive Risk Scoring.
- **Reporting:** **Streamlit + Plotly** (High-end interactive dashboard).
- **Environment:** Unified via `requirements.txt` for enterprise portability.

---

## 🚀 Deployment Instructions

### 1. Environment Setup
Clone this repository and install necessary libraries:
```bash
pip install -r requirements.txt
```

### 2. Generate Data Model (If needed)
The system uses a pre-calculated star schema model. If you modify the source data, refresh the model:
```bash
python notebooks/build_data_model.py
```

### 3. Launch Intelligence Dashboard
Run the professional Streamlit application:
```bash
streamlit run dashboard/app.py
```

---

## 📊 Dashboard Modules

### 1. Executive Overview
Real-time tracking of:
- **Churn Rate:** Benchmarked against industry standards.
- **Revenue at Risk:** The financial impact of projected churn.
- **Retention Performance:** Gauge-based KPI tracking.

### 2. Churn Driver Analysis
Identification of behavioral patterns:
- **Contract Impact:** Monthly vs. Multi-year churn dynamics.
- **Internet Service:** Quality-to-Churn correlation.
- **Pricing:** Behavioral shifts based on monthly billing amounts.

### 3. Smart Risk Segmentation
An actionable list of customers for the retention team:
- **Risk Tiers:** Tiering customers into High, Medium, and Low risk cohorts.
- **Action Plans:** Specific recommendations for each risk level.
- **Financial Simulation:** Real-time ROI calculation of retention efforts.

---
*Developed for a Business Intelligence Portfolio. Dataset: IBM Telco Customer Churn.*
