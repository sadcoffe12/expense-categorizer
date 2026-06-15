"""Rules suggestion engine for Phase 7"""
from sqlalchemy.orm import Session
from sqlalchemy import and_
from datetime import datetime, timedelta
from collections import Counter
import re

from ..models import Rule, Category, CategorizationHistory


def suggest_rules_from_history(db: Session, min_confidence: float = 0.7):
    """Suggest new rules based on categorization history"""
    # Get uncategorized items that users have manually categorized
    history_items = db.query(CategorizationHistory).filter(
        and_(
            CategorizationHistory.user_action == "accepted",
            CategorizationHistory.suggested_category_id == None
        )
    ).all()
    
    if not history_items:
        return []
    
    # Group by category
    by_category = {}
    for item in history_items:
        if item.suggested_category_id not in by_category:
            by_category[item.suggested_category_id] = []
        by_category[item.suggested_category_id].append(item.original_text)
    
    suggestions = []
    for category_id, texts in by_category.items():
        # Extract common keywords
        keywords = extract_keywords(texts)
        
        # For each keyword, check if rule already exists
        for keyword, count in keywords.items():
            confidence = count / len(texts)
            
            if confidence >= min_confidence:
                existing_rule = db.query(Rule).filter(
                    and_(
                        Rule.keyword == keyword,
                        Rule.category_id == category_id
                    )
                ).first()
                
                if not existing_rule:
                    suggestions.append({
                        "keyword": keyword,
                        "category_id": category_id,
                        "confidence": round(confidence, 2),
                        "occurrences": count,
                        "frequency": f"{round(confidence * 100, 0):.0f}%"
                    })
    
    return suggestions


def extract_keywords(texts: list[str]) -> dict:
    """Extract common keywords from a list of texts"""
    # Clean and split texts
    all_words = []
    for text in texts:
        # Remove special characters and convert to lowercase
        cleaned = re.sub(r"[^a-záéíóúñ\s]", " ", text.lower())
        words = [w for w in cleaned.split() if len(w) > 2]
        all_words.extend(words)
    
    # Count word frequencies
    word_counts = Counter(all_words)
    
    # Filter common words
    stopwords = {"para", "con", "del", "de", "la", "el", "los", "las", "y", "por", "en"}
    keywords = {word: count for word, count in word_counts.items() if word not in stopwords and count > 0}
    
    return dict(sorted(keywords.items(), key=lambda x: x[1], reverse=True)[:5])


def get_recommended_rules(db: Session, category_id: int | None = None):
    """Get recommended rules for categories"""
    query = db.query(Rule).filter(Rule.active == True)
    
    if category_id:
        query = query.filter(Rule.category_id == category_id)
    
    rules = query.all()
    
    return [
        {
            "id": r.id,
            "keyword": r.keyword,
            "category_id": r.category_id,
            "confidence": round(float(r.confidence), 2),
            "active": r.active,
            "created_at": r.created_at.isoformat()
        }
        for r in rules
    ]


def auto_create_rules_from_accepted(db: Session):
    """Automatically create rules when users accept suggestions"""
    suggestions = suggest_rules_from_history(db, min_confidence=0.8)
    
    for suggestion in suggestions:
        new_rule = Rule(
            keyword=suggestion["keyword"],
            category_id=suggestion["category_id"],
            confidence=suggestion["confidence"],
            active=True
        )
        db.add(new_rule)
    
    db.commit()
    return {"rules_created": len(suggestions)}
