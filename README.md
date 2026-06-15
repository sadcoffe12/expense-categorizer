# 💰 Expense Categorizer

> **Smart expense tracking with AI-powered automatic categorization**

A modern, full-featured application to track, categorize, and analyze your expenses with machine learning assistance. Automatically learns from your spending patterns and suggests budget alerts.

![Status](https://img.shields.io/badge/status-production%20ready-brightgreen)
![License](https://img.shields.io/badge/license-MIT-blue)
![Python](https://img.shields.io/badge/python-3.13-blue)
![Node.js](https://img.shields.io/badge/node.js-18%2B-green)

---

## 📋 Table of Contents

- [Features](#-features)
- [System Requirements](#-system-requirements)
- [Installation](#-installation)
- [Quick Start](#-quick-start)
- [Usage Guide](#-usage-guide)
- [Project Structure](#-project-structure)
- [Troubleshooting](#-troubleshooting)
- [Next Steps](#-next-steps)

---

## ✨ Features

### 📊 **Dashboard Analytics**
- **Summary Cards**: Total expenses, average transaction, transaction count, median
- **30-Day Trends**: Interactive line chart showing spending over time
- **Category Breakdown**: Pie chart visualization of expenses by category
- **Type Distribution**: Cards showing Gasto (Expense) vs Ingreso (Income) totals
- **Real-time Updates**: Data refreshes as you add/edit expenses

### 💳 **Expense Management**
- **CRUD Operations**: Create, read, update, delete expenses
- **Inline Editing**: Edit expenses directly in the table without dialog
- **Advanced Filtering**: Filter by category, date range, amount range
- **Rich Details**: Track date, description, amount, category, type, location, notes
- **Data Import**: Bulk import expenses from CSV files
- **Quick Search**: Find expenses by description

### 🤖 **Smart Categorization**
- **Auto-Categorization**: AI suggests categories based on description
- **Learning Engine**: Learns from your categorization patterns
- **Rule Management**: Create custom rules for automatic categorization
- **Confidence Scores**: See how confident the system is in suggestions
- **Pattern Analysis**: Analyzes 4 time periods (session, last month, 3 months, 6 months)
- **Rule Suggestions**: System recommends rules based on repeated patterns

### 💼 **Budget Management**
- **Budget Creation**: Set budgets per category with flexible periods
- **Budget Tracking**: Monitor spending vs budget in real-time
- **Visual Progress**: Color-coded bars show budget status (green = under, red = over)
- **Multiple Periods**: Monthly, yearly, or custom date range budgets
- **Budget Details**: See total spent, remaining amount, and percentage

### 🚨 **Alert System**
- **Budget Alerts**: Automatic notifications when budget is exceeded
- **Alert History**: View all alerts with timestamps
- **Acknowledgment**: Mark alerts as read/acknowledged
- **Alert Types**: Budget overflow, spending spike, and more
- **Summary Dashboard**: See unacknowledged alerts count by type

### 📈 **Advanced Analytics & Reporting**

#### Anomaly Detection
- **Statistical Analysis**: Identifies unusual spending using Z-score algorithm
- **Severity Levels**: Low, Medium, High severity classification
- **Category Breakdown**: See anomalies per expense category
- **Customizable Threshold**: Adjust sensitivity (1-4 standard deviations)

#### Spending Patterns
- **Day-of-Week Analysis**: Spending habits by weekday (Monday-Sunday)
- **Week-of-Month Analysis**: Patterns across different weeks of the month
- **Average Calculations**: See average spending per day/week
- **Transaction Counts**: How many transactions per period

#### Forecasting
- **Predictive Analytics**: Forecast spending for next 1, 3, or 6 months
- **Confidence Intervals**: Upper and lower bounds for predictions
- **Historical Basis**: Based on last 6 months of data
- **Accuracy Metrics**: Confidence level percentage

### 💾 **Data Export & Reports**

#### CSV Export
- **Full Export**: Export all expenses with filters
- **Filtered Export**: Export by category, date range
- **Complete Metadata**: Includes dates, descriptions, amounts, categories, types, locations, notes

#### Summary Reports
- **JSON Format**: Structured data for integration
- **CSV Format**: Spreadsheet-compatible format
- **Period Analysis**: Customizable report period (7-365 days)
- **Category Breakdown**: Total spending per category
- **Highest Categories**: Identify top spending categories

### 📌 **Saved Dashboard Views**

- **Custom Layouts**: Save your preferred dashboard configurations
- **Filter Presets**: Save common filter combinations
- **Quick Access**: Switch between saved views with one click
- **Persistent Storage**: Views saved to database

---

## 📦 System Requirements

### Minimum Requirements
- **Operating System**: Windows 10+, macOS 10.14+, Linux (Ubuntu 18.04+)
- **Python**: 3.13 or higher
- **Node.js**: 18.0 or higher
- **RAM**: 2 GB minimum
- **Disk Space**: 500 MB for installation + data
- **Browser**: Modern browser (Chrome, Firefox, Safari, Edge)

### Recommended
- **Python**: 3.13 or higher
- **Node.js**: 20 LTS or higher
- **RAM**: 4 GB or more
- **Disk Space**: 2 GB or more
- **Internet**: Connection for initial setup (not required for operation)

---

## 🔧 Installation

### Step 1: Download/Clone the Project

```bash
# If you have Git installed:
git clone <repository-url>
cd "Expense Categorizer"

# Or manually extract the folder
cd "Expense Categorizer/expense-categorizer"
```

### Step 2: Install Python Dependencies

**Windows:**
```cmd
cd backend
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

**macOS/Linux:**
```bash
cd backend
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### Step 3: Install Node.js Dependencies

**Windows:**
```cmd
cd ../frontend
npm install
```

**macOS/Linux:**
```bash
cd ../frontend
npm install
```

### Step 4: Verify Installation

**Backend Check:**
```bash
cd backend
python -m py_compile app/main.py
```
Should show no errors.

**Frontend Check:**
```bash
cd frontend
npm list vite
```
Should show vite version installed.

---

## 🚀 Quick Start

### Option 1: Automatic Startup (Recommended)

**Windows:**
Double-click `start.bat` in the project root

**macOS/Linux:**
```bash
./start.sh
```

**Cross-Platform:**
```bash
python3 start.py
```

### Option 2: Manual Startup

**Terminal 1 - Backend (Python):**
```bash
cd backend
python run.py
```
Expected output:
```
🚀 Iniciando Expense Categorizer API...
📍 http://localhost:8000
INFO:     Uvicorn running on http://0.0.0.0:8000
```

**Terminal 2 - Frontend (React):**
```bash
cd frontend
npm run dev
```
Expected output:
```
VITE v5.0.0  ready in XXXms
➜  Local:   http://localhost:5173/
```

### Step 3: Access the Application

Open your browser and go to:
```
http://localhost:5173
```

---

## 📖 Usage Guide

### First Time Setup

1. **Open Dashboard**: You'll see an empty dashboard on first visit
2. **Import Data** (Optional):
   - Prepare a CSV file with your expenses
   - Columns: `date`, `description`, `amount`, `category` (optional)
3. **Create Categories**: Go to **Settings** → Add categories for your spending
4. **Add First Expense**: 
   - Click **+ Add Expense**
   - Fill in details (date, description, amount, category)
   - Click **Save**

### 💳 Managing Expenses

#### Adding an Expense
1. Click **+ New Expense** button
2. Fill in the form:
   - **Date**: When the expense occurred
   - **Description**: What was purchased (e.g., "Coffee at Starbucks")
   - **Amount**: How much was spent
   - **Category**: Select from dropdown (or system suggests)
   - **Type**: Gasto (expense) or Ingreso (income)
   - **Location** (optional): Where the purchase happened
   - **Notes** (optional): Additional details
3. Click **Save**

#### Editing an Expense
1. In **Expenses** page, hover over any row
2. Click the **Edit** icon (pencil)
3. Modify the fields you want to change
4. Click **Update**

Or use inline editing:
1. Double-click the value you want to change
2. Edit directly in the table
3. Click outside to save

#### Deleting an Expense
1. In **Expenses** page, hover over the row
2. Click the **Delete** icon (trash)
3. Confirm deletion

#### Filtering Expenses
1. Use the **Category Filter** dropdown
2. Select a category or "All"
3. Table updates automatically
4. For more filters, use the **Advanced Filters** section

### 🤖 Smart Categorization

#### Let the System Suggest Categories
1. Start typing in the **Description** field
2. The system analyzes the text
3. A suggested category appears below
4. Accept by clicking the suggestion or change manually

#### Create Custom Rules
1. Go to **Settings** → **Rules & Learning** tab
2. Click **Add New Rule**
3. Enter:
   - **Keyword**: Text to match (e.g., "starbucks")
   - **Category**: Category to assign
   - **Confidence**: 0.5-1.0 (higher = more confident)
4. Click **Create Rule**
5. The system will use this rule for matching

#### View Recommended Rules
1. Go to **Settings** → **Rules & Learning** tab
2. Click **Refresh Suggestions**
3. System shows rules it recommends based on patterns
4. Accept good recommendations

### 💼 Setting Budgets

#### Create a Budget
1. Go to **Settings** → **Budgets** tab
2. Click **Create New Budget**
3. Fill in:
   - **Category**: Which category to budget for
   - **Amount**: Monthly/period limit
   - **Period**: Monthly (default), yearly, or custom
4. Click **Create Budget**

#### Monitor Budgets
- Progress bars show:
  - **Green** = Under budget (safe)
  - **Yellow** = Approaching limit (caution)
  - **Red** = Over budget (alert)
- Values shown: Spent / Budget Amount

#### Delete a Budget
- Click the **X** button next to the budget
- Confirm deletion

### 📊 Viewing Alerts

1. Go to **Settings** → **Alerts** tab
2. See summary cards:
   - Total unacknowledged alerts
   - Breakdown by alert type
3. Table shows all alerts:
   - **Type**: What triggered the alert
   - **Message**: Details about the alert
   - **Status**: Acknowledged or Pending
   - **Date**: When the alert was created
4. Click **Mark as Acknowledged** to dismiss

### 📈 Advanced Analytics

#### Anomaly Detection
1. Go to **Reports** → **Anomalies** tab
2. Choose time period: Last 7, 30, 60, or 90 days
3. View:
   - Total expenses in period
   - Number of anomalies found
   - Table with details:
     - **Category**: Where the unusual spending occurred
     - **Amount**: The unusual transaction
     - **Z-Score**: How unusual (higher = more unusual)
     - **Severity**: Low/Medium/High
4. Red highlights indicate high-severity anomalies

#### Spending Patterns
1. Go to **Reports** → **Patterns** tab
2. View analysis by:
   - **Day of Week**: Which days you spend the most
   - **Week of Month**: Spending patterns within the month
3. Linear progress bars show spending per period
4. Totals and averages displayed

#### Forecasting
1. Go to **Reports** → **Forecasting** tab
2. Select forecast period: 1, 3, or 6 months
3. View:
   - **Average Monthly Spending**: Historical average
   - **Confidence Level**: How reliable the forecast is
   - **Forecast Chart**: Predicted spending with upper/lower bounds
4. Bounds show range of likely values

#### Export Data
1. Click **Export to CSV** button (on any Reports tab)
2. Choose filters if desired:
   - **Date Range**: From and to dates
   - **Categories**: Select specific categories or all
3. CSV file downloads automatically
4. Open in Excel, Google Sheets, or any spreadsheet app

### 📌 Saved Views

#### Save Current View
1. After setting up filters and preferences
2. Click **Save View** button
3. Enter a name (e.g., "Home Office Expenses")
4. Click **Save**

#### Load Saved View
1. Click **View** dropdown or **Saved Views** menu
2. Select the view you want
3. Dashboard instantly updates with saved filters

#### Delete Saved View
1. From the **Saved Views** menu
2. Hover over the view name
3. Click **Delete**
4. Confirm deletion

---

## 📂 Project Structure

```
expense-categorizer/
├── README.md                          # This file
├── STARTUP_GUIDE.md                   # Detailed startup instructions
├── PHASES_7_9_GUIDE.md               # Technical API documentation
│
├── start.sh                           # Linux/macOS startup script
├── start.bat                          # Windows startup script
├── start.py                           # Cross-platform startup script
│
├── backend/                           # FastAPI Backend
│   ├── run.py                         # Entry point
│   ├── requirements.txt               # Python dependencies
│   ├── data/                          # SQLite database folder (auto-created)
│   │   └── expense.db
│   └── app/
│       ├── main.py                    # FastAPI app initialization
│       ├── database.py                # Database configuration
│       ├── schemas.py                 # Pydantic models
│       ├── models.py                  # SQLAlchemy ORM models
│       └── api/                       # API endpoints
│           ├── setup.py               # Initialization endpoint
│           ├── expenses.py            # Expense CRUD
│           ├── analytics.py           # Dashboard analytics
│           ├── categorize.py          # Auto-categorization
│           ├── budgets.py             # Budget management
│           ├── alerts.py              # Alert system
│           ├── rules_engine.py        # ML rule learning
│           ├── advanced_analytics.py  # Anomalies, patterns, forecasting
│           ├── exports.py             # CSV export, reports
│           └── saved_views.py         # Dashboard view persistence
│
└── frontend/                          # React + TypeScript Frontend
    ├── package.json                   # npm dependencies
    ├── vite.config.ts                 # Vite configuration
    ├── tsconfig.json                  # TypeScript configuration
    ├── public/
    │   └── favicon.svg
    ├── src/
    │   ├── main.tsx                   # Entry point
    │   ├── App.tsx                    # Main app component
    │   ├── api/
    │   │   └── client.ts              # Axios API client
    │   ├── types/
    │   │   └── index.ts               # TypeScript type definitions
    │   └── pages/
    │       ├── DashboardPage.tsx      # Main dashboard
    │       ├── ExpensesPage.tsx       # Expense management
    │       ├── ReportsPage.tsx        # Advanced analytics
    │       └── SettingsPage.tsx       # Budgets, alerts, rules
    └── index.html
```

---

## 🎯 Common Tasks

### Import Expenses from CSV

**CSV Format:**
```csv
date,description,amount,category,type
2024-01-15,Coffee,5.50,Food,Gasto
2024-01-16,Salary,3000.00,Income,Ingreso
2024-01-17,Gas,45.00,Transportation,Gasto
```

**Columns:**
- `date`: YYYY-MM-DD format
- `description`: Any text
- `amount`: Numeric value
- `category`: Category name (must exist)
- `type`: "Gasto" (expense) or "Ingreso" (income)

### Check API Documentation

1. With backend running, go to:
   ```
   http://localhost:8000/docs
   ```
2. Interactive Swagger UI shows all endpoints
3. Click any endpoint to see:
   - Parameters
   - Request format
   - Response examples
   - Try it button to test

### Change Backend Port

If port 8000 is already in use:

1. Open `backend/run.py`
2. Find the line with `port=8000`
3. Change to different port (e.g., 8001)
4. Restart backend

### Change Frontend Port

If port 5173 is already in use:

1. Vite automatically tries next ports (5174, 5175, etc.)
2. Or edit `frontend/vite.config.ts`:
   ```typescript
   server: {
     port: 3000  // Change this
   }
   ```
3. Restart frontend

---

## ❌ Troubleshooting

### Backend won't start

**Error: "unable to open database file"**
```bash
# Create data directory
mkdir backend/data
```

**Error: "Port 8000 already in use"**
```bash
# Find process using port (macOS/Linux)
lsof -i :8000

# Kill process
kill -9 <PID>
```

**Error: "ModuleNotFoundError"**
```bash
# Reinstall Python dependencies
cd backend
pip install -r requirements.txt
```

### Frontend won't start

**Error: "vite: command not found"**
```bash
# Install dependencies
cd frontend
npm install
npm run dev
```

**Error: "Cannot connect to backend"**
1. Verify backend is running: `http://localhost:8000/docs`
2. Check firewall settings
3. Try refreshing the page

### Application loads but shows no data

1. **First time?** You need to add expenses first
2. **Expecting imported data?** Go to Settings and import CSV
3. **Data not showing?** Check browser console for errors (F12)

### Performance is slow

1. **Too many expenses?** Try filtering by date or category
2. **Too many years of data?** Forecasting/anomaly detection slow with large datasets
3. **Browser issue?** Try different browser or clear cache

### Lost connection between frontend and backend

1. **Restart both services:**
   ```bash
   # Stop both terminals
   # Restart with ./start.sh or start.bat
   ```
2. **Clear browser cache:** Ctrl+Shift+Delete (Cmd+Shift+Delete on Mac)
3. **Check firewall:** Allow localhost traffic

---

## 🚀 Next Steps

### Want More Features?

- **Mobile App**: Web version works on mobile (responsive design)
- **Advanced ML**: More sophisticated categorization with neural networks
- **Banking Integration**: Auto-import from bank APIs
- **Multi-User**: Shared expense tracking with families
- **Cloud Sync**: Backup and sync across devices

### Contribute

Have ideas? Found a bug? Want to improve the application?

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

### Get Help

- Check `STARTUP_GUIDE.md` for detailed setup help
- See `PHASES_7_9_GUIDE.md` for technical API details
- Review API docs at `http://localhost:8000/docs`

---

## 📚 Documentation

| Document | Purpose |
|----------|---------|
| [README.md](README.md) | This file - Overview and usage |
| [STARTUP_GUIDE.md](STARTUP_GUIDE.md) | Step-by-step startup instructions |
| [PHASES_7_9_GUIDE.md](PHASES_7_9_GUIDE.md) | Technical API reference |

---

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

---

## 🎉 Enjoy Your Expense Categorizer!

**Version**: 2.0 (Phase 5-9 Complete)  
**Last Updated**: 2026-06-14  
**Status**: Production Ready ✅

---

## 📞 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review detailed guides in STARTUP_GUIDE.md
3. Check API documentation at http://localhost:8000/docs
4. Review browser console (F12) for error messages

**Happy expense tracking! 💰**
