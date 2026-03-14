# Center Location Dashboard

A Flask web application that replicates the Power BI dashboard from `Center_Location_Work_V1.pbix`.

## Project Structure

```
center_dashboard/
├── app.py                  ← Main Flask application (run this)
├── requirements.txt        ← Python dependencies
├── README.md
├── templates/
│   └── index.html          ← Full dashboard UI
└── data/                   ← ⚠️ Place your Excel source files here
    ├── MASTER_NEW.xlsx
    ├── Impact.xlsx
    ├── Cluster_Master.xlsx
    └── Criteria.xlsx
```

## Setup in PyCharm

### Step 1 – Place Source Files
Copy the 4 Excel files into the `data/` folder:
- `MASTER_NEW.xlsx`
- `Impact.xlsx`
- `Cluster_Master.xlsx`
- `Criteria.xlsx`

### Step 2 – Create Virtual Environment (in PyCharm)
1. Open **File → Settings → Project → Python Interpreter**
2. Click the gear icon → **Add Interpreter → Virtual Environment**
3. Click **OK**

### Step 3 – Install Dependencies
In PyCharm's Terminal, run:
```bash
pip install -r requirements.txt
```

Or in PyCharm:
1. Go to **File → Settings → Project → Python Interpreter**
2. Click **+** and search for `flask` and `openpyxl`

### Step 4 – Run the App
- **Method 1 (PyCharm):** Right-click `app.py` → **Run 'app'**
- **Method 2 (Terminal):** `python app.py`

### Step 5 – Open in Browser
Navigate to: **http://localhost:5000**

---

## Dashboard Pages

### 📍 Active Centers
- Interactive map showing all center locations (from `MASTER_NEW.xlsx`)
- Filters: Entity, Cluster, Sub Cluster, Status, Type, Distance Slab
- Summary cards: Total Centers, Training Camps, Fixed Centers
- Center details table

### 📋 QP & Criteria
- Job roles / Qualification Packs per center (from `Criteria.xlsx`)
- Filters: Entity, Cluster, Sector, Business Vertical
- Bar charts by Sector and Business Vertical
- Full criteria table with age and qualification requirements

### 📊 Impact
- Enrollment / Certification / Placement metrics (from `Impact.xlsx`)
- Summary cards: Enrolled, Certified, Placed, Dropout
- Cluster-wise comparison chart
- Top clusters by placement ranking

---

## Data Sources (mirrors PBIX)

| Page in PBIX       | Excel Source File     | Sheet Used       |
|--------------------|-----------------------|------------------|
| Active Centers     | `MASTER_NEW.xlsx`     | `Sheet1`         |
| Cluster mapping    | `Cluster_Master.xlsx` | `Sheet1`         |
| QP and Criteria    | `Criteria.xlsx`       | `Center Wise`    |
| Impact             | `Impact.xlsx`         | `Sheet1`         |
