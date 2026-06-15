"""Alert system endpoints for Phase 7"""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from datetime import datetime
from typing import List

from ..database import get_db
from ..models import Alert, Expense

router = APIRouter(prefix="/api/alerts", tags=["alerts"])


@router.get("/", response_model=List[dict])
def get_alerts(
    expense_id: int | None = None,
    acknowledged: bool | None = None,
    alert_type: str | None = None,
    db: Session = Depends(get_db)
):
    """Get all alerts with optional filtering"""
    query = db.query(Alert)
    
    if expense_id:
        query = query.filter(Alert.expense_id == expense_id)
    if acknowledged is not None:
        query = query.filter(Alert.acknowledged == acknowledged)
    if alert_type:
        query = query.filter(Alert.alert_type == alert_type)
    
    alerts = query.order_by(Alert.created_at.desc()).all()
    
    return [
        {
            "id": a.id,
            "expense_id": a.expense_id,
            "alert_type": a.alert_type,
            "message": a.message,
            "acknowledged": a.acknowledged,
            "created_at": a.created_at.isoformat()
        }
        for a in alerts
    ]


@router.get("/summary")
def get_alert_summary(db: Session = Depends(get_db)):
    """Get summary of unacknowledged alerts"""
    alerts = db.query(Alert).filter(Alert.acknowledged == False).all()
    
    by_type = {}
    for alert in alerts:
        alert_type = alert.alert_type
        by_type[alert_type] = by_type.get(alert_type, 0) + 1
    
    return {
        "total_unacknowledged": len(alerts),
        "by_type": by_type,
        "alerts": [
            {
                "id": a.id,
                "alert_type": a.alert_type,
                "message": a.message,
                "created_at": a.created_at.isoformat()
            }
            for a in alerts[:10]  # Latest 10
        ]
    }


@router.put("/{alert_id}/acknowledge", response_model=dict)
def acknowledge_alert(alert_id: int, db: Session = Depends(get_db)):
    """Mark alert as acknowledged"""
    alert = db.query(Alert).filter(Alert.id == alert_id).first()
    if not alert:
        raise HTTPException(status_code=404, detail="Alert not found")
    
    alert.acknowledged = True
    db.commit()
    db.refresh(alert)
    
    return {
        "id": alert.id,
        "acknowledged": True,
        "message": "Alert acknowledged"
    }


@router.delete("/{alert_id}", response_model=dict)
def delete_alert(alert_id: int, db: Session = Depends(get_db)):
    """Delete alert"""
    alert = db.query(Alert).filter(Alert.id == alert_id).first()
    if not alert:
        raise HTTPException(status_code=404, detail="Alert not found")
    
    db.delete(alert)
    db.commit()
    
    return {"message": "Alert deleted successfully"}


@router.post("/check-budgets")
def check_budgets(db: Session = Depends(get_db)):
    """Check all budgets and create alerts for exceeded ones"""
    # This would be called by a scheduled task
    from ..models import Budget, Category
    from datetime import timedelta
    from sqlalchemy import and_
    
    now = datetime.now()
    budgets = db.query(Budget).all()
    alerts_created = 0
    
    for budget in budgets:
        # Calculate period
        if budget.period == "month":
            period_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            period_end = (period_start + timedelta(days=32)).replace(day=1) - timedelta(seconds=1)
        elif budget.period == "year":
            period_start = now.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
            period_end = now.replace(month=12, day=31, hour=23, minute=59, second=59)
        else:
            period_start = budget.start_date if budget.start_date else now - timedelta(days=30)
            period_end = budget.end_date if budget.end_date else now
        
        # Get spending
        spent = sum(
            e.amount for e in db.query(Expense).filter(
                and_(
                    Expense.category_id == budget.category_id,
                    Expense.date >= period_start,
                    Expense.date <= period_end,
                    Expense.type == "Gasto"
                )
            ).all()
        )
        
        # Check if exceeded
        if spent > budget.amount:
            category = db.query(Category).filter(Category.id == budget.category_id).first()
            existing_alert = db.query(Alert).filter(
                and_(
                    Alert.alert_type == "budget_exceeded",
                    Alert.message.contains(f"Category: {category.name}")
                )
            ).first()
            
            if not existing_alert:
                alert = Alert(
                    expense_id=None,
                    alert_type="budget_exceeded",
                    message=f"Budget exceeded for {category.name}. Spent: ${spent:.2f} / Budget: ${budget.amount:.2f}",
                    acknowledged=False
                )
                db.add(alert)
                alerts_created += 1
    
    db.commit()
    return {"alerts_created": alerts_created}
