"""Budget management endpoints for Phase 7"""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from sqlalchemy import and_
from datetime import datetime, timedelta
from typing import List

from ..database import get_db
from ..models import Budget, Category, Expense
from ..schemas import BudgetResponse, BudgetCreate

router = APIRouter(prefix="/api/budgets", tags=["budgets"])


@router.get("/", response_model=List[BudgetResponse])
def get_budgets(
    category_id: int | None = None,
    period: str | None = None,
    db: Session = Depends(get_db)
):
    """Get all budgets with optional filtering"""
    query = db.query(Budget)
    
    if category_id:
        query = query.filter(Budget.category_id == category_id)
    if period:
        query = query.filter(Budget.period == period)
    
    budgets = query.all()
    return [
        {
            "id": b.id,
            "category_id": b.category_id,
            "category": db.query(Category).filter(Category.id == b.category_id).first().name,
            "amount": float(b.amount),
            "period": b.period,
            "start_date": b.start_date.isoformat() if b.start_date else None,
            "end_date": b.end_date.isoformat() if b.end_date else None,
            "created_at": b.created_at.isoformat()
        }
        for b in budgets
    ]


@router.get("/{budget_id}")
def get_budget(budget_id: int, db: Session = Depends(get_db)):
    """Get budget with current spending analysis"""
    budget = db.query(Budget).filter(Budget.id == budget_id).first()
    if not budget:
        raise HTTPException(status_code=404, detail="Budget not found")
    
    category = db.query(Category).filter(Category.id == budget.category_id).first()
    
    # Calculate current spending for the period
    now = datetime.now()
    if budget.period == "month":
        period_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        period_end = (period_start + timedelta(days=32)).replace(day=1) - timedelta(seconds=1)
    elif budget.period == "year":
        period_start = now.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
        period_end = now.replace(month=12, day=31, hour=23, minute=59, second=59)
    else:
        period_start = budget.start_date if budget.start_date else now - timedelta(days=30)
        period_end = budget.end_date if budget.end_date else now
    
    spent = db.query(Expense).filter(
        and_(
            Expense.category_id == budget.category_id,
            Expense.date >= period_start,
            Expense.date <= period_end,
            Expense.type == "Gasto"
        )
    ).all()
    
    total_spent = sum(e.amount for e in spent)
    count = len(spent)
    
    return {
        "id": budget.id,
        "category_id": budget.category_id,
        "category": category.name,
        "amount": float(budget.amount),
        "period": budget.period,
        "start_date": budget.start_date.isoformat() if budget.start_date else None,
        "end_date": budget.end_date.isoformat() if budget.end_date else None,
        "spent": float(total_spent),
        "count": count,
        "remaining": float(budget.amount - total_spent),
        "percentage": (total_spent / budget.amount * 100) if budget.amount > 0 else 0,
        "created_at": budget.created_at.isoformat()
    }


@router.post("/", response_model=dict)
def create_budget(
    budget_data: BudgetCreate,
    db: Session = Depends(get_db)
):
    """Create new budget"""
    # Verify category exists
    category = db.query(Category).filter(Category.id == budget_data.category_id).first()
    if not category:
        raise HTTPException(status_code=404, detail="Category not found")
    
    # Check if budget already exists for this category and period
    existing = db.query(Budget).filter(
        and_(
            Budget.category_id == budget_data.category_id,
            Budget.period == budget_data.period
        )
    ).first()
    
    if existing:
        raise HTTPException(status_code=400, detail="Budget already exists for this category and period")
    
    new_budget = Budget(
        category_id=budget_data.category_id,
        amount=budget_data.amount,
        period=budget_data.period,
        start_date=budget_data.start_date,
        end_date=budget_data.end_date
    )
    
    db.add(new_budget)
    db.commit()
    db.refresh(new_budget)
    
    return {
        "id": new_budget.id,
        "category_id": new_budget.category_id,
        "amount": float(new_budget.amount),
        "period": new_budget.period,
        "created_at": new_budget.created_at.isoformat(),
        "message": "Budget created successfully"
    }


@router.put("/{budget_id}", response_model=dict)
def update_budget(
    budget_id: int,
    budget_data: BudgetCreate,
    db: Session = Depends(get_db)
):
    """Update existing budget"""
    budget = db.query(Budget).filter(Budget.id == budget_id).first()
    if not budget:
        raise HTTPException(status_code=404, detail="Budget not found")
    
    # Verify category exists
    category = db.query(Category).filter(Category.id == budget_data.category_id).first()
    if not category:
        raise HTTPException(status_code=404, detail="Category not found")
    
    budget.amount = budget_data.amount
    budget.period = budget_data.period
    budget.start_date = budget_data.start_date
    budget.end_date = budget_data.end_date
    
    db.commit()
    db.refresh(budget)
    
    return {
        "id": budget.id,
        "amount": float(budget.amount),
        "period": budget.period,
        "message": "Budget updated successfully"
    }


@router.delete("/{budget_id}", response_model=dict)
def delete_budget(budget_id: int, db: Session = Depends(get_db)):
    """Delete budget"""
    budget = db.query(Budget).filter(Budget.id == budget_id).first()
    if not budget:
        raise HTTPException(status_code=404, detail="Budget not found")
    
    db.delete(budget)
    db.commit()
    
    return {"message": "Budget deleted successfully"}
