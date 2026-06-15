from sqlalchemy import Column, Integer, String, Float, Date, DateTime, ForeignKey, Boolean, Text, JSON
from sqlalchemy.orm import relationship
from datetime import datetime
from .database import Base

class Category(Base):
    __tablename__ = "categories"
    
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True, nullable=False)
    type = Column(String)  # Fixed, Variable, Income
    color_hex = Column(String, default="#666666")
    icon = Column(String, default="💰")
    created_at = Column(DateTime, default=datetime.utcnow)
    
    expenses = relationship("Expense", back_populates="category")
    rules = relationship("Rule", back_populates="category")
    budgets = relationship("Budget", back_populates="category")

class Expense(Base):
    __tablename__ = "expenses"
    
    id = Column(Integer, primary_key=True)
    date = Column(Date, nullable=False)
    description = Column(String, nullable=False)
    description_cleaned = Column(String)
    amount = Column(Float, nullable=False)
    category_id = Column(Integer, ForeignKey("categories.id"))
    type = Column(String)  # Gasto, Ingreso
    location = Column(String)
    notes = Column(String)
    source = Column(String)  # manual, import, api
    created_at = Column(DateTime, default=datetime.utcnow)
    
    category = relationship("Category", back_populates="expenses")
    alerts = relationship("Alert", back_populates="expense")
    tags = relationship("ExpenseTag", back_populates="expense")
    history = relationship("CategorizationHistory", back_populates="expense")

class Rule(Base):
    __tablename__ = "rules"
    
    id = Column(Integer, primary_key=True)
    keyword = Column(String, nullable=False)
    category_id = Column(Integer, ForeignKey("categories.id"))
    confidence = Column(Float, default=0.8)
    active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    
    category = relationship("Category", back_populates="rules")

class Budget(Base):
    __tablename__ = "budgets"
    
    id = Column(Integer, primary_key=True)
    category_id = Column(Integer, ForeignKey("categories.id"))
    amount = Column(Float)
    period = Column(String)  # month, year
    start_date = Column(Date)
    end_date = Column(Date)
    created_at = Column(DateTime, default=datetime.utcnow)
    
    category = relationship("Category", back_populates="budgets")

class Alert(Base):
    __tablename__ = "alerts"
    
    id = Column(Integer, primary_key=True)
    expense_id = Column(Integer, ForeignKey("expenses.id"))
    alert_type = Column(String)  # budget_exceeded, anomaly
    message = Column(String)
    acknowledged = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    
    expense = relationship("Expense", back_populates="alerts")

class CategorizationHistory(Base):
    __tablename__ = "categorization_history"
    
    id = Column(Integer, primary_key=True)
    expense_id = Column(Integer, ForeignKey("expenses.id"))
    original_text = Column(String)
    cleaned_text = Column(String)
    suggested_category_id = Column(Integer)
    user_action = Column(String)  # confirmed, rejected, skipped
    created_at = Column(DateTime, default=datetime.utcnow)
    
    expense = relationship("Expense", back_populates="history")

class ExpenseTag(Base):
    __tablename__ = "expense_tags"
    
    id = Column(Integer, primary_key=True)
    expense_id = Column(Integer, ForeignKey("expenses.id"))
    tag_name = Column(String)
    created_at = Column(DateTime, default=datetime.utcnow)
    
    expense = relationship("Expense", back_populates="tags")

class DashboardConfig(Base):
    __tablename__ = "dashboard_config"
    
    id = Column(Integer, primary_key=True)
    config_name = Column(String)
    filters = Column(JSON)  # Filtros guardados
    layout_preferences = Column(JSON)  # Preferencias de layout
    created_at = Column(DateTime, default=datetime.utcnow)
