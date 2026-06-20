import sys
import os

# Add scripts directory to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'scripts')))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pytest
from scoring import classify_championship

def test_classify_esport():
    # Esport always gets 0 score
    score, flag = classify_championship("Juara 1", "Nasional", "Penyelenggara Esport", "Sertifikat Mobile Legends Championship")
    assert score == 0
    assert "E-SPORT" in flag

def test_classify_snbp():
    score, flag = classify_championship("Juara 1", "Nasional", "Lomba SNBP", "Sertifikat Lomba SNBP")
    assert score == 0
    assert "SNBP" in flag

def test_classify_non_competition():
    # Non competition certificate gets 0
    score, flag = classify_championship("Peserta", "Nasional", "Google", "Google Cloud Workshop Participant")
    assert score == 0
    assert "Peserta/Partisipasi" in flag or "NON-KOMPETISI" in flag

def test_classify_tahfidz_rekognisi():
    # Tahfidz with >5 juz on Nasional level
    score, flag = classify_championship("Rekognisi", "Nasional", "Kemenag", "Sertifikat Tahfidz Quran 10 Juz")
    assert score == 8
    assert "Tahfidz" in flag

    # Tahfidz with <=5 juz gets 0 score
    score, flag = classify_championship("Rekognisi", "Nasional", "Kemenag", "Sertifikat Tahfidz Quran 3 Juz")
    assert score == 0
    assert "NON-KOMPETISI" in flag or "Tahfidz" in flag

def test_classify_kota_non_government():
    # Kota/Kabupaten non-government organizer gets 0
    score, flag = classify_championship("Juara 1", "Kota/Kabupaten", "Klub Catur Swasta", "Lomba Catur Kota")
    assert score == 0
    assert "non-pemerintah" in flag
