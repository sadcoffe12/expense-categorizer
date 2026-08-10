"""
Configuration for Expense Categorizer.
Centralized settings for file paths, thresholds, and triggers.
"""

# File paths (relative to project root)
RULES_FILE = "categorization_rules.csv"
TEMPLATES_FILE = "templates.txt"
HISTORY_FILE = "uncategorized_history.csv"

# Text similarity threshold for pattern matching (0.0 - 1.0)
SIMILARITY_THRESHOLD = 0.8  # 80% similarity required

# Maximum age for history entries (in days)
# Entries older than this are automatically cleaned up
MAX_HISTORY_AGE_DAYS = 213  # ~7 months

# Pattern trigger thresholds (number of repetitions required to suggest a rule)
# These values determine when the learning system suggests a new categorization rule

# This session: minimum repetitions within current session
TRIGGER_THIS_SESSION = 5

# Last month: minimum repetitions in the past 30 days
TRIGGER_LAST_MONTH = 3

# Last 3 months: minimum repetitions in the past 90 days
TRIGGER_LAST_3_MONTHS = 3

# Last 6 months: minimum repetitions in the past 180 days
TRIGGER_LAST_6_MONTHS = 3

# Minimum confidence score for suggesting a category via pattern matching (0.0 - 1.0)
MIN_CONFIDENCE_FOR_SUGGESTION = 0.6  # 60% confidence
