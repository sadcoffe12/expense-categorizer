from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from sqlalchemy import and_, or_
from datetime import date
from typing import List, Optional
from decimal import Decimal

from ..database import get_db
from ..models import Expense, Category
from ..schemas import ExpenseBase

router = APIRouter(prefix="/api/expenses", tags=["expenses"])

@router.get("/", response_model=List[dict])
async def get_expenses(
    db: Session = Depends(get_db),
    skip: int = Query(0, ge=0),
    limit: int = Query(100, ge=1, le=1000),
    category_id: Optional[int] = None,
    date_from: Optional[date] = None,
    date_to: Optional[date] = None,
    min_amount: Optional[float] = None,
    max_amount: Optional[float] = None,
):
    """Get expenses with optional filtering"""
    query = db.query(Expense)
    
    if category_id:
        query = query.filter(Expense.category_id == category_id)
    if date_from:
        query = query.filter(Expense.date >= date_from)
    if date_to:
        query = query.filter(Expense.date <= date_to)
    if min_amount is not None:
        query = query.filter(Expense.amount >= min_amount)
    if max_amount is not None:
        query = query.filter(Expense.amount <= max_amount)
    
    expenses = query.order_by(Expense.date.desc()).offset(skip).limit(limit).all()
    
    return [
        {
            "id": e.id,
            "date": e.date.isoformat(),
            "description": e.description,
            "amount": float(e.amount),
            "category": e.category.name if e.category else "Sin categoría",
            "category_id": e.category_id,
            "type": e.type,
            "location": e.location,
            "notes": e.notes,
            "created_at": e.created_at.isoformat()
        }
        for e in expenses
    ]

@router.get("/{expense_id}", response_model=dict)
async def get_expense(expense_id: int, db: Session = Depends(get_db)):
    """Get a single expense by ID"""
    expense = db.query(Expense).filter(Expense.id == expense_id).first()
    
    if not expense:
        raise HTTPException(status_code=404, detail="Expense not found")
    
    return {
        "id": expense.id,
        "date": expense.date.isoformat(),
        "description": expense.description,
        "amount": float(expense.amount),
        "category": expense.category.name if expense.category else "Sin categoría",
        "category_id": expense.category_id,
        "type": expense.type,
        "location": expense.location,
        "notes": expense.notes,
        "created_at": expense.created_at.isoformat()
    }

@router.put("/{expense_id}", response_model=dict)
async def update_expense(
    expense_id: int,
    category_id: Optional[int] = None,
    description: Optional[str] = None,
    amount: Optional[float] = None,
    type_: Optional[str] = None,
    location: Optional[str] = None,
    notes: Optional[str] = None,
    db: Session = Depends(get_db)
):
    """Update an expense"""
    expense = db.query(Expense).filter(Expense.id == expense_id).first()
    
    if not expense:
        raise HTTPException(status_code=404, detail="Expense not found")
    
    if category_id is not None:
        category = db.query(Category).filter(Category.id == category_id).first()
        if not category:
            raise HTTPException(status_code=400, detail="Category not found")
        expense.category_id = category_id
    
    if description is not None:
        expense.description = description
    if amount is not None:
        expense.amount = amount
    if type_ is not None:
        expense.type = type_
    if location is not None:
        expense.location = location
    if notes is not None:
        expense.notes = notes
    
    db.commit()
    db.refresh(expense)
    
    return {
        "id": expense.id,
        "date": expense.date.isoformat(),
        "description": expense.description,
        "amount": float(expense.amount),
        "category_id": expense.category_id,
        "type": expense.type,
        "location": expense.location,
        "notes": expense.notes
    }

@router.delete("/{expense_id}")
async def delete_expense(expense_id: int, db: Session = Depends(get_db)):
    """Delete an expense"""
    expense = db.query(Expense).filter(Expense.id == expense_id).first()
    
    if not expense:
        raise HTTPException(status_code=404, detail="Expense not found")
    
    db.delete(expense)
    db.commit()
    
    return {"deleted": True, "id": expense_id}
