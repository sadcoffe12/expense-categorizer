import unicodedata
import re
from difflib import SequenceMatcher
from typing import Tuple, Optional

def normalize_text(text: str, is_transaction: bool = False) -> str:
    """
    Normaliza texto (quitar tildes, símbolos, minúsculas).
    Si is_transaction=True, también elimina ruido bancario (tarjetas, IDs).
    
    Replica la lógica del main.py original.
    """
    if not isinstance(text, str):
        return str(text) if text is not None else ""

    # 1. Limpieza básica (Minúsculas y acentos)
    text = text.lower().strip()
    text = ''.join(c for c in unicodedata.normalize('NFD', text)
                  if unicodedata.category(c) != 'Mn')

    # 2. Limpieza de Ruido Bancario (Solo si es una descripción de gasto)
    if is_transaction:
        # Quita "tarj nro. 1234" o "tarjeta 1234"
        text = re.sub(r'tarj\s?nro\.?\s?\d+', '', text)
        # Quita números largos de 5 o más dígitos (IDs de transacción)
        text = re.sub(r'\d{5,}', '', text)

    # 3. Colapsar espacios múltiples
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text

def get_similarity(a: str, b: str) -> float:
    """Retorna el ratio de similitud entre dos strings (0-1)."""
    return SequenceMatcher(None, a, b).ratio()

def guess_category(cleaned_desc: str, rules: list) -> Tuple[Optional[str], Optional[str]]:
    """
    Intenta predecir categoría basándose en reglas existentes.
    Replica la lógica del main.py original.
    
    Args:
        cleaned_desc: Descripción normalizada
        rules: Lista de tuplas (keyword, tipo, categoría, new_description)
    
    Returns:
        Tupla (tipo, categoría) o (None, None) si no hay buena coincidencia
    """
    best_score = 0
    best_match = (None, None)  # (Tipo, Categoría)
    
    for keyword, t_val, c_val, _ in rules:
        # Comparamos la keyword de la regla con el texto actual
        score = SequenceMatcher(None, cleaned_desc, keyword).ratio()
        if score > best_score:
            best_score = score
            best_match = (t_val, c_val)
    
    # Si la confianza es mayor al 60%, sugerimos
    SIMILARITY_THRESHOLD = 0.6
    return best_match if best_score > SIMILARITY_THRESHOLD else (None, None)

def suggest_rule_from_patterns(
    pattern: str, 
    rules: list,
    period_info: dict
) -> Tuple[Optional[str], Optional[str]]:
    """
    Sugiere categoría basada en un patrón detectado en el historial.
    
    Args:
        pattern: Patrón normalizado de descripción
        rules: Reglas existentes
        period_info: Info del período donde se detectó (para mostrar contexto)
    
    Returns:
        Tupla sugerida (tipo, categoría)
    """
    tipo_sug, cat_sug = guess_category(pattern, rules)
    return (tipo_sug, cat_sug)

def extract_keywords_from_description(description: str) -> list:
    """Extrae palabras claves de una descripción para análisis de patrones."""
    cleaned = normalize_text(description, is_transaction=True)
    # Palabras de al menos 4 caracteres, excluyendo números
    words = [w for w in cleaned.split() if len(w) >= 4 and not w.isdigit()]
    return words
