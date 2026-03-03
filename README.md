# 📉 Customer Churn Analysis — Telecom

> End-to-end data analysis project identifying at-risk customers using behavioral patterns and predictive modeling.
>
> 
> live demo https://customerchurnanalysis2.netlify.app/

![Python](https://img.shields.io/badge/Python-3.10-blue?style=flat-square&logo=python)
![Pandas](https://img.shields.io/badge/Pandas-2.0-green?style=flat-square)
![Scikit-learn](https://img.shields.io/badge/Scikit--learn-1.3-orange?style=flat-square)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen?style=flat-square)

---

## 🎯 Business Problem

Customer churn costs telecom companies billions annually. Acquiring a new customer costs **5–7x more** than retaining an existing one. This project answers:

- What is our churn rate and how does it compare to industry benchmarks?
- Which customers are most likely to churn, and why?
- Can we predict churn with enough accuracy to enable proactive retention?
- What business actions should we take?

---

## 📊 Key Findings

| Metric | Value |
|--------|-------|
| Overall Churn Rate | **26.5%** (industry avg: ~18%) |
| Highest Risk Segment | Month-to-month contracts (**42% churn**) |
| Critical Retention Window | First **12 months** (63% of all churn) |
| Estimated Revenue at Risk | **$435K/year** |

### 🔑 Top Insights
- **Contract type** is the strongest churn predictor — month-to-month customers churn at 5x the rate of two-year customers
- **Fiber optic** users have surprisingly high churn (41.9%) despite being a premium service
- Customers **without tech support** are 2.7x more likely to churn
- **Electronic check** payment method correlates strongly with churn (45.3%)

---

## 🤖 Model Performance

| Metric | Score |
|--------|-------|
| Accuracy | 82.1% |
| AUC-ROC | 0.847 |
| Precision | 79.3% |
| Recall | 76.8% |
| F1-Score | 78.0% |

---

## 🛠️ Tech Stack

- **Python** — Core analysis
- **Pandas & NumPy** — Data manipulation
- **Matplotlib & Seaborn** — Visualization
- **Scikit-learn** — Machine learning
- **Jupyter Notebook** — Interactive analysis

---

## 📁 Project Structure

```
customer-churn-analysis/
│
├── data/
│   └── WA_Fn-UseC_-Telco-Customer-Churn.csv    # Dataset (Kaggle)
│
├── notebooks/
│   └── customer_churn_analysis.ipynb            # Main analysis notebook
│
├── outputs/
│   ├── at_risk_customers.csv                    # Model predictions
│   └── figures/                                 # All visualizations
│
├── dashboard/
│   └── customer_churn_dashboard.html            # Interactive dashboard
│
├── requirements.txt
└── README.md
```

---

## 🚀 Getting Started

```bash
# Clone the repo
git clone https://github.com/yourusername/customer-churn-analysis.git
cd customer-churn-analysis

# Install dependencies
pip install -r requirements.txt

# Download dataset from Kaggle
# https://www.kaggle.com/datasets/blastchar/telco-customer-churn
# Place in data/ folder

# Launch notebook
jupyter notebook notebooks/customer_churn_analysis.ipynb
```

---

## 💡 Business Recommendations

1. **Launch contract upgrade campaign** — offer 15% discount for month-to-month → annual upgrade
2. **90-day onboarding program** — personalized check-ins for first-year customers
3. **Tech support bundle** — include free 3-month tech support with new fiber sign-ups
4. **Auto-pay incentive** — 5% discount for switching from electronic check

---

## 📈 Next Steps

- [ ] Experiment with Random Forest & XGBoost for improved accuracy
- [ ] Add SHAP values for model explainability
- [ ] Build real-time scoring API (FastAPI)
- [ ] Deploy interactive Streamlit dashboard

---

## 📄 Dataset

IBM Telco Customer Churn — available on [Kaggle](https://www.kaggle.com/datasets/blastchar/telco-customer-churn)

---

*Built as part of my data analytics portfolio. Feel free to fork and use!*
