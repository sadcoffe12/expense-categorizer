"""Data export endpoints for Phase 9"""
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session
from sqlalchemy import and_
from datetime import datetime, timedelta
from typing import Optional
import csv
import io
from pathlib import Path

from ..database import get_db
from ..models import Expense, Category

router = APIRouter(prefix="/api/exports", tags=["exports"])


@router.get("/expenses-csv")
def export_expenses_csv(
    category_id: Optional[int] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    db: Session = Depends(get_db)
):
    """Export expenses to CSV format"""
    query = db.query(Expense)
    
    if category_id:
        query = query.filter(Expense.category_id == category_id)
    
    if date_from:
        from_date = datetime.fromisoformat(date_from).date()
        query = query.filter(Expense.date >= from_date)
    
    if date_to:
        to_date = datetime.fromisoformat(date_to).date()
        query = query.filter(Expense.date <= to_date)
    
    expenses = query.order_by(Expense.date.desc()).all()
    
    # Create CSV content
    output = io.StringIO()
    writer = csv.writer(output)
    
    # Write header
    writer.writerow([
        "Date", "Description", "Amount", "Category", "Type",
        "Location", "Notes", "Created At"
    ])
    
    # Write data
    for expense in expenses:
        category = db.query(Category).filter(
            Category.id == expense.category_id
        ).first()
        
        writer.writerow([
            expense.date.isoformat(),
            expense.description,
            expense.amount,
            category.name if category else "Unknown",
            expense.type,
            expense.location or "",
            expense.notes or "",
            expense.created_at.isoformat()
        ])
    
    # Create file
    file_content = output.getvalue()
    filename = f"expenses_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    
    return FileResponse(
        io.BytesIO(file_content.encode()),
        media_type="text/csv",
        filename=filename
    )


@router.get("/summary-report")
def generate_summary_report(
    format: str = "json",
    period_days: int = 30,
    db: Session = Depends(get_db)
):
    """Generate spending summary report"""
    cutoff_date = datetime.now() - timedelta(days=period_days)
    
    expenses = db.query(Expense).filter(
        and_(
            Expense.date >= cutoff_date,
            Expense.type == "Gasto"
        )
    ).all()
    
    total = sum(e.amount for e in expenses)
    
    # By category
    by_category = {}
    for expense in expenses:
        cat = db.query(Category).filter(Category.id == expense.category_id).first()
        cat_name = cat.name if cat else "Unknown"
        
        if cat_name not in by_category:
            by_category[cat_name] = {"amount": 0, "count": 0}
        
        by_category[cat_name]["amount"] += expense.amount
        by_category[cat_name]["count"] += 1
    
    report = {
        "report_date": datetime.now().isoformat(),
        "period_days": period_days,
        "total_spending": round(total, 2),
        "transaction_count": len(expenses),
        "average_transaction": round(total / len(expenses), 2) if expenses else 0,
        "by_category": by_category,
        "highest_category": max(
            by_category.items(),
            key=lambda x: x[1]["amount"],
            default=(None, {})
        )[0]
    }
    
    if format.lower() == "csv":
        output = io.StringIO()
        writer = csv.writer(output)
        
        writer.writerow(["Spending Report", datetime.now().strftime("%Y-%m-%d")])
        writer.writerow([])
        writer.writerow(["Period (days)", period_days])
        writer.writerow(["Total Spending", f"${report['total_spending']:.2f}"])
        writer.writerow(["Transactions", report["transaction_count"]])
        writer.writerow(["Average Per Transaction", f"${report['average_transaction']:.2f}"])
        writer.writerow([])
        writer.writerow(["Category", "Amount", "Count"])
        
        for category, data in report["by_category"].items():
            writer.writerow([
                category,
                f"${data['amount']:.2f}",
                data["count"]
            ])
        
        return FileResponse(
            io.BytesIO(output.getvalue().encode()),
            media_type="text/csv",
            filename=f"report_{datetime.now().strftime('%Y%m%d')}.csv"
        )
    
    return report
