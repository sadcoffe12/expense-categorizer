from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from typing import Optional, List

from ..database import get_db
from ..models import Expense, Category, Rule
from ..utils.categorizer import normalize_text, guess_category

router = APIRouter(prefix="/api", tags=["categorization"])

@router.post("/categorize")
async def categorize_expense(
    description: str,
    db: Session = Depends(get_db),
):
    """
    Suggest category for an expense description.
    Uses existing rules and pattern matching.
    """
    # Normalize the input
    cleaned = normalize_text(description, is_transaction=True)
    
    # Load active rules
    rules = db.query(Rule).filter(Rule.active == True).all()
    rules_list = [
        (r.keyword, r.category.type if r.category else "", 
         r.category.name if r.category else "", "")
        for r in rules
    ]
    
    # Get suggestion
    tipo_sug, cat_sug = guess_category(cleaned, rules_list)
    
    if not cat_sug:
        return {
            "description": description,
            "cleaned": cleaned,
            "suggested_category": None,
            "suggested_type": None,
            "confidence": 0
        }
    
    # Find the category
    category = db.query(Category).filter(Category.name == cat_sug).first()
    
    return {
        "description": description,
        "cleaned": cleaned,
        "suggested_category": cat_sug,
        "suggested_type": tipo_sug,
        "confidence": 0.75,  # Placeholder confidence
        "category_id": category.id if category else None
    }

@router.get("/rules", response_model=List[dict])
async def get_rules(
    db: Session = Depends(get_db),
    active_only: bool = True,
):
    """Get all categorization rules"""
    query = db.query(Rule)
    if active_only:
        query = query.filter(Rule.active == True)
    
    rules = query.all()
    
    return [
        {
            "id": r.id,
            "keyword": r.keyword,
            "category": r.category.name if r.category else None,
            "category_id": r.category_id,
            "confidence": r.confidence,
            "active": r.active,
            "created_at": r.created_at.isoformat()
        }
        for r in rules
    ]

@router.post("/rules", response_model=dict)
async def create_rule(
    keyword: str,
    category_id: int,
    confidence: float = 0.8,
    db: Session = Depends(get_db),
):
    """Create a new categorization rule"""
    # Validate category exists
    category = db.query(Category).filter(Category.id == category_id).first()
    if not category:
        raise HTTPException(status_code=400, detail="Category not found")
    
    # Normalize keyword
    keyword_normalized = normalize_text(keyword, is_transaction=True)
    
    # Check if rule already exists
    existing = db.query(Rule).filter(
        Rule.keyword == keyword_normalized,
        Rule.category_id == category_id
    ).first()
    
    if existing:
        raise HTTPException(status_code=400, detail="Rule already exists")
    
    rule = Rule(
        keyword=keyword_normalized,
        category_id=category_id,
        confidence=min(max(confidence, 0.0), 1.0),  # Clamp to 0-1
        active=True
    )
    
    db.add(rule)
    db.commit()
    db.refresh(rule)
    
    return {
        "id": rule.id,
        "keyword": rule.keyword,
        "category": category.name,
        "category_id": rule.category_id,
        "confidence": rule.confidence,
        "active": rule.active,
        "created_at": rule.created_at.isoformat()
    }

@router.put("/rules/{rule_id}", response_model=dict)
async def update_rule(
    rule_id: int,
    keyword: Optional[str] = None,
    category_id: Optional[int] = None,
    confidence: Optional[float] = None,
    active: Optional[bool] = None,
    db: Session = Depends(get_db),
):
    """Update a categorization rule"""
    rule = db.query(Rule).filter(Rule.id == rule_id).first()
    
    if not rule:
        raise HTTPException(status_code=404, detail="Rule not found")
    
    if keyword is not None:
        rule.keyword = normalize_text(keyword, is_transaction=True)
    
    if category_id is not None:
        category = db.query(Category).filter(Category.id == category_id).first()
        if not category:
            raise HTTPException(status_code=400, detail="Category not found")
        rule.category_id = category_id
    
    if confidence is not None:
        rule.confidence = min(max(confidence, 0.0), 1.0)
    
    if active is not None:
        rule.active = active
    
    db.commit()
    db.refresh(rule)
    
    return {
        "id": rule.id,
        "keyword": rule.keyword,
        "category": rule.category.name if rule.category else None,
        "category_id": rule.category_id,
        "confidence": rule.confidence,
        "active": rule.active
    }

@router.delete("/rules/{rule_id}")
async def delete_rule(rule_id: int, db: Session = Depends(get_db)):
    """Delete a categorization rule"""
    rule = db.query(Rule).filter(Rule.id == rule_id).first()
    
    if not rule:
        raise HTTPException(status_code=404, detail="Rule not found")
    
    db.delete(rule)
    db.commit()
    
    return {"deleted": True, "id": rule_id}
