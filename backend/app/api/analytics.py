from fastapi import APIRouter, Depends, Query, HTTPException
from sqlalchemy.orm import Session
from sqlalchemy import func
from datetime import date, datetime, timedelta
from typing import Optional, List
import statistics

from ..database import get_db
from ..models import Expense, Category

router = APIRouter(prefix="/api/analytics", tags=["analytics"])

@router.get("/summary")
async def get_summary(
    db: Session = Depends(get_db),
    date_from: Optional[date] = None,
    date_to: Optional[date] = None,
):
    """Get summary statistics for expenses"""
    query = db.query(Expense)
    
    if date_from:
        query = query.filter(Expense.date >= date_from)
    if date_to:
        query = query.filter(Expense.date <= date_to)
    
    expenses = query.all()
    
    if not expenses:
        return {
            "total": 0,
            "count": 0,
            "average": 0,
            "min": 0,
            "max": 0,
            "median": 0,
            "by_type": {},
            "by_category": {}
        }
    
    amounts = [float(e.amount) for e in expenses]
    
    # By type
    type_summary = {}
    for exp_type in set(e.type for e in expenses if e.type):
        type_expenses = [e for e in expenses if e.type == exp_type]
        type_summary[exp_type] = {
            "count": len(type_expenses),
            "total": sum(float(e.amount) for e in type_expenses),
            "average": sum(float(e.amount) for e in type_expenses) / len(type_expenses)
        }
    
    # By category
    cat_summary = {}
    for exp in expenses:
        cat_name = exp.category.name if exp.category else "Sin categoría"
        if cat_name not in cat_summary:
            cat_summary[cat_name] = {"count": 0, "total": 0}
        cat_summary[cat_name]["count"] += 1
        cat_summary[cat_name]["total"] += float(exp.amount)
    
    return {
        "total": sum(amounts),
        "count": len(amounts),
        "average": sum(amounts) / len(amounts),
        "min": min(amounts),
        "max": max(amounts),
        "median": statistics.median(amounts) if len(amounts) > 0 else 0,
        "by_type": type_summary,
        "by_category": cat_summary
    }

@router.get("/trends")
async def get_trends(
    db: Session = Depends(get_db),
    period: str = Query("daily", regex="^(daily|weekly|monthly)$"),
    months: int = Query(6, ge=1, le=24),
):
    """Get expense trends over time"""
    end_date = datetime.now().date()
    start_date = end_date - timedelta(days=months * 30)
    
    expenses = db.query(Expense).filter(
        Expense.date >= start_date,
        Expense.date <= end_date
    ).order_by(Expense.date).all()
    
    trends = {}
    
    for expense in expenses:
        if period == "daily":
            key = expense.date.isoformat()
        elif period == "weekly":
            week_start = expense.date - timedelta(days=expense.date.weekday())
            key = week_start.isoformat()
        else:  # monthly
            key = expense.date.strftime("%Y-%m")
        
        if key not in trends:
            trends[key] = {"total": 0, "count": 0, "average": 0}
        
        trends[key]["total"] += float(expense.amount)
        trends[key]["count"] += 1
    
    # Calculate averages
    for key in trends:
        trends[key]["average"] = trends[key]["total"] / trends[key]["count"]
    
    return {
        "period": period,
        "data": trends
    }

@router.get("/category/{category_id}")
async def get_category_detail(
    category_id: int,
    db: Session = Depends(get_db),
    date_from: Optional[date] = None,
    date_to: Optional[date] = None,
):
    """Get detailed analytics for a specific category"""
    category = db.query(Category).filter(Category.id == category_id).first()
    
    if not category:
        raise HTTPException(status_code=404, detail="Category not found")
    
    query = db.query(Expense).filter(Expense.category_id == category_id)
    
    if date_from:
        query = query.filter(Expense.date >= date_from)
    if date_to:
        query = query.filter(Expense.date <= date_to)
    
    expenses = query.all()
    amounts = [float(e.amount) for e in expenses]
    
    return {
        "category": {
            "id": category.id,
            "name": category.name,
            "type": category.type,
            "color": category.color_hex
        },
        "stats": {
            "total": sum(amounts) if amounts else 0,
            "count": len(amounts),
            "average": sum(amounts) / len(amounts) if amounts else 0,
            "min": min(amounts) if amounts else 0,
            "max": max(amounts) if amounts else 0
        },
        "expenses": [
            {
                "id": e.id,
                "date": e.date.isoformat(),
                "description": e.description,
                "amount": float(e.amount),
                "location": e.location,
                "notes": e.notes
            }
            for e in expenses
        ]
    }

@router.get("/budget-vs-actual")
async def get_budget_vs_actual(
    db: Session = Depends(get_db),
    year: int = None,
    month: int = None,
):
    """Compare budgeted vs actual expenses"""
    from ..models import Budget
    
    if not year:
        year = datetime.now().year
    if not month:
        month = datetime.now().month
    
    start_date = date(year, month, 1)
    if month == 12:
        end_date = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(year, month + 1, 1) - timedelta(days=1)
    
    # Get budgets for this period
    budgets = db.query(Budget).filter(
        Budget.period == "month",
        Budget.start_date <= end_date,
        Budget.end_date >= start_date
    ).all()
    
    result = []
    
    for budget in budgets:
        # Get actual expenses for this category in this period
        actual_expenses = db.query(func.sum(Expense.amount)).filter(
            Expense.category_id == budget.category_id,
            Expense.date >= start_date,
            Expense.date <= end_date
        ).scalar() or 0
        
        result.append({
            "category": budget.category.name,
            "budget": float(budget.amount),
            "actual": float(actual_expenses),
            "variance": float(budget.amount) - float(actual_expenses),
            "percentage": (float(actual_expenses) / float(budget.amount) * 100) if budget.amount > 0 else 0
        })
    
    return {"period": f"{year}-{month:02d}", "data": result}
