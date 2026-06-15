"""Advanced analytics with anomaly detection for Phase 9"""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from sqlalchemy import and_
from datetime import datetime, timedelta
from typing import List
import statistics

from ..database import get_db
from ..models import Expense, Category

router = APIRouter(prefix="/api/advanced-analytics", tags=["advanced_analytics"])


@router.get("/anomalies")
def detect_anomalies(
    days: int = 30,
    std_dev_threshold: float = 2.0,
    db: Session = Depends(get_db)
):
    """Detect anomalous spending patterns using statistical analysis"""
    cutoff_date = datetime.now() - timedelta(days=days)
    
    expenses = db.query(Expense).filter(
        and_(
            Expense.date >= cutoff_date,
            Expense.type == "Gasto"
        )
    ).all()
    
    if not expenses:
        return {"anomalies": [], "analysis": {}}
    
    # Group by category
    by_category = {}
    for expense in expenses:
        if expense.category_id not in by_category:
            by_category[expense.category_id] = []
        by_category[expense.category_id].append(expense.amount)
    
    anomalies = []
    
    for category_id, amounts in by_category.items():
        if len(amounts) < 3:
            continue
        
        mean = statistics.mean(amounts)
        stdev = statistics.stdev(amounts) if len(amounts) > 1 else 0
        
        if stdev > 0:
            for amount in amounts:
                z_score = (amount - mean) / stdev
                if abs(z_score) >= std_dev_threshold:
                    category = db.query(Category).filter(Category.id == category_id).first()
                    anomalies.append({
                        "amount": float(amount),
                        "category": category.name if category else "Unknown",
                        "z_score": round(z_score, 2),
                        "mean": round(mean, 2),
                        "deviation": round(stdev, 2),
                        "severity": "high" if abs(z_score) > 3 else "medium"
                    })
    
    return {
        "anomalies": sorted(anomalies, key=lambda x: abs(x["z_score"]), reverse=True),
        "analysis": {
            "total_expenses": len(expenses),
            "period_days": days,
            "anomalies_found": len(anomalies),
            "threshold": std_dev_threshold
        }
    }


@router.get("/spending-patterns")
def analyze_spending_patterns(
    months: int = 3,
    db: Session = Depends(get_db)
):
    """Analyze spending patterns and trends"""
    cutoff_date = datetime.now() - timedelta(days=months * 30)
    
    expenses = db.query(Expense).filter(
        and_(
            Expense.date >= cutoff_date,
            Expense.type == "Gasto"
        )
    ).all()
    
    # Analyze by day of week
    day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    by_day_of_week = {i: [] for i in range(7)}
    
    for expense in expenses:
        dow = expense.date.weekday()
        by_day_of_week[dow].append(expense.amount)
    
    day_analysis = {}
    for dow, amounts in by_day_of_week.items():
        if amounts:
            day_analysis[day_names[dow]] = {
                "count": len(amounts),
                "total": round(sum(amounts), 2),
                "average": round(statistics.mean(amounts), 2)
            }
    
    # Analyze by time of month
    by_week_of_month = {i: [] for i in range(1, 5)}
    for expense in expenses:
        week = (expense.date.day - 1) // 7 + 1
        by_week_of_month[week].append(expense.amount)
    
    week_analysis = {}
    for week, amounts in by_week_of_month.items():
        if amounts:
            week_analysis[f"Week {week}"] = {
                "count": len(amounts),
                "total": round(sum(amounts), 2),
                "average": round(statistics.mean(amounts), 2)
            }
    
    return {
        "analysis_period_months": months,
        "by_day_of_week": day_analysis,
        "by_week_of_month": week_analysis,
        "total_transactions": len(expenses),
        "total_spending": round(sum(e.amount for e in expenses), 2)
    }


@router.get("/forecasting")
def forecast_spending(months_ahead: int = 3, db: Session = Depends(get_db)):
    """Forecast future spending based on historical data"""
    # Use last 6 months of data
    cutoff_date = datetime.now() - timedelta(days=180)
    
    expenses = db.query(Expense).filter(
        and_(
            Expense.date >= cutoff_date,
            Expense.type == "Gasto"
        )
    ).all()
    
    # Calculate monthly averages
    by_month = {}
    for expense in expenses:
        month_key = expense.date.strftime("%Y-%m")
        if month_key not in by_month:
            by_month[month_key] = []
        by_month[month_key].append(expense.amount)
    
    monthly_totals = [sum(amounts) for amounts in by_month.values()]
    
    if not monthly_totals:
        return {"forecast": [], "confidence": 0}
    
    # Simple average forecast
    avg_monthly = statistics.mean(monthly_totals)
    stdev = statistics.stdev(monthly_totals) if len(monthly_totals) > 1 else 0
    
    forecast = []
    for i in range(months_ahead):
        future_date = datetime.now() + timedelta(days=30 * (i + 1))
        forecast.append({
            "month": future_date.strftime("%Y-%m"),
            "predicted_spending": round(avg_monthly, 2),
            "lower_bound": round(max(0, avg_monthly - 2 * stdev), 2),
            "upper_bound": round(avg_monthly + 2 * stdev, 2)
        })
    
    return {
        "forecast": forecast,
        "confidence": 0.6 if len(monthly_totals) > 2 else 0.3,
        "based_on_months": len(monthly_totals),
        "average_monthly_spending": round(avg_monthly, 2)
    }
