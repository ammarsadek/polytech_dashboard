# 📘 Polytech OEE & Production Dashboard
*A Streamlit App for Real-Time KPI Monitoring & OEE Analysis*

---

## Overview

This Streamlit dashboard provides advanced analytics for factory production lines, machines, and molds.  
It automatically calculates:

- **OEE (Overall Equipment Effectiveness)**
- **Availability**
- **Performance**
- **Quality**
- Production KPIs  
- Reject analysis  
- Downtime impact  
- Machine & Product performance comparison  
- Daily and monthly trends  

The app is built using **Python**, **Pandas**, and **Streamlit**, with a clean filtering system and interactive visualizations.

---

## Features

### Filters
- Date range  
- Machine selection  
- Product/Mold selection  
- Configurable working hours/day  

### Dashboards
#### **Overview Dashboard**
- OEE by machine & product  
- KPI summary cards  
- Bar charts & ranked tables  

#### **Machine Drill-Down**
- Detailed OEE breakdown  
- Daily machine trends  
- Full machine-specific records  

#### **Product / Mold Drill-Down**
- Product/Mold OEE  
- OEE trends  
- Reject & performance analysis  

#### **Trend Dashboard**
- Daily OEE trend  
- Monthly OEE trend  

#### **Raw Data Viewer**
- Cleaned dataset with all calculated metrics  

---

## 🧮 OEE Calculation Logic

The dashboard uses the standard OEE formula:
Where:

- **Availability** = 1 – (Downtime / Planned Hours)  
- **Performance** = Average performance ratio  
- **Quality** = Good Units / Total Units  

---

## 📁 Project Structure
```
oee-dashboard/
│
├── app.py                     # Streamlit application
├── Production data.xlsx        # Production dataset (optional for public repos)
├── requirements.txt           # Python dependencies
└── README.md                  # Documentation
```

## ⚙️ Installation

### 1️⃣ Clone the Repository

```bash
git clone https://github.com/yourusername/oee-dashboard.git
cd oee-dashboard
```

### 2️⃣ Install Dependencies
```bash
pip install -r requirements.txt
```
### 3️⃣ Run the Application
```bash
streamlit run app.py
````
