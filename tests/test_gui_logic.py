import sys
import os
import pytest
from unittest.mock import MagicMock

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import scoring_gui

def test_header_aliases():
    # Verify canonical keys are correctly mapped
    aliases = scoring_gui.HEADER_ALIASES
    assert aliases.get('nilai') == 'nilai'
    assert aliases.get('score') == 'nilai'
    assert aliases.get('nilai_sertifikat') == 'nilai'
    
    assert aliases.get('komentar') == 'komentar'
    assert aliases.get('comment') == 'komentar'
    assert aliases.get('catatan') == 'komentar'
    assert aliases.get('notes') == 'komentar'
    
    assert aliases.get('link') == 'url'
    assert aliases.get('url') == 'url'

def test_validate_score():
    # We can mock the ScoringApp class and test _validate_score
    app = MagicMock()
    # Bind the actual method to the mock instance
    app._validate_score = scoring_gui.ScoringApp._validate_score.__get__(app, scoring_gui.ScoringApp)
    
    # Valid scores
    ok, msg = app._validate_score("10")
    assert ok is True
    assert msg == ""
    
    ok, msg = app._validate_score("0")
    assert ok is True
    
    ok, msg = app._validate_score("20")
    assert ok is True
    
    # Invalid score - Empty
    ok, msg = app._validate_score("")
    assert ok is False
    assert "Please enter a nilai" in msg
    
    # Invalid score - Not a number
    ok, msg = app._validate_score("abc")
    assert ok is False
    assert "must be a number" in msg
    
    # Invalid score - Out of range
    ok, msg = app._validate_score("21")
    assert ok is False
    assert "must be between 0 and 20" in msg
    
    ok, msg = app._validate_score("-1")
    assert ok is False
    assert "must be between 0 and 20" in msg
