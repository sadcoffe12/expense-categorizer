"""Saved views/dashboards endpoints for Phase 9"""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from datetime import datetime
from typing import List
import json

from ..database import get_db
from ..models import DashboardConfig

router = APIRouter(prefix="/api/views", tags=["saved_views"])


@router.get("/", response_model=List[dict])
def list_saved_views(db: Session = Depends(get_db)):
    """List all saved dashboard views"""
    views = db.query(DashboardConfig).all()
    
    return [
        {
            "id": v.id,
            "name": v.config_name,
            "filters": json.loads(v.filters) if v.filters else {},
            "layout": json.loads(v.layout_preferences) if v.layout_preferences else {},
            "created_at": v.created_at.isoformat()
        }
        for v in views
    ]


@router.get("/{view_id}")
def get_saved_view(view_id: int, db: Session = Depends(get_db)):
    """Get a specific saved view configuration"""
    view = db.query(DashboardConfig).filter(DashboardConfig.id == view_id).first()
    
    if not view:
        raise HTTPException(status_code=404, detail="View not found")
    
    return {
        "id": view.id,
        "name": view.config_name,
        "filters": json.loads(view.filters) if view.filters else {},
        "layout": json.loads(view.layout_preferences) if view.layout_preferences else {},
        "created_at": view.created_at.isoformat()
    }


@router.post("/", response_model=dict)
def create_saved_view(
    name: str,
    filters: dict | None = None,
    layout: dict | None = None,
    db: Session = Depends(get_db)
):
    """Create a new saved view"""
    new_view = DashboardConfig(
        config_name=name,
        filters=json.dumps(filters or {}),
        layout_preferences=json.dumps(layout or {})
    )
    
    db.add(new_view)
    db.commit()
    db.refresh(new_view)
    
    return {
        "id": new_view.id,
        "name": new_view.config_name,
        "message": "View saved successfully"
    }


@router.put("/{view_id}", response_model=dict)
def update_saved_view(
    view_id: int,
    name: str | None = None,
    filters: dict | None = None,
    layout: dict | None = None,
    db: Session = Depends(get_db)
):
    """Update a saved view"""
    view = db.query(DashboardConfig).filter(DashboardConfig.id == view_id).first()
    
    if not view:
        raise HTTPException(status_code=404, detail="View not found")
    
    if name:
        view.config_name = name
    if filters is not None:
        view.filters = json.dumps(filters)
    if layout is not None:
        view.layout_preferences = json.dumps(layout)
    
    db.commit()
    db.refresh(view)
    
    return {"id": view.id, "message": "View updated successfully"}


@router.delete("/{view_id}", response_model=dict)
def delete_saved_view(view_id: int, db: Session = Depends(get_db)):
    """Delete a saved view"""
    view = db.query(DashboardConfig).filter(DashboardConfig.id == view_id).first()
    
    if not view:
        raise HTTPException(status_code=404, detail="View not found")
    
    db.delete(view)
    db.commit()
    
    return {"message": "View deleted successfully"}
