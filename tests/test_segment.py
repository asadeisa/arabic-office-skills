"""segment() decides which characters share a run — and so where they render."""
import pytest

import arabic_docx
import arabic_pptx

CASES = [
    ("Vue.js 3 مع Nuxt 3.", [("Vue.js 3", False), (" مع ", True), ("Nuxt 3", False), (".", True)]),
    ("(20 فأكثر)", [("(", True), ("20", False), (" فأكثر)", True)]),
    ("DECIMAL(5,2) هنا", [("DECIMAL(5,2)", False), (" هنا", True)]),
    ("نستخدم C# و C++ و .NET", [("نستخدم ", True), ("C#", False), (" و ", True),
                                ("C++", False), (" و ", True), (".NET", False)]),
    ("الهاتف +963 912", [("الهاتف ", True), ("+963 912", False)]),
    ("نسبة 50% فقط", [("نسبة ", True), ("50", False), ("% فقط", True)]),
]


@pytest.mark.parametrize("module", [arabic_docx, arabic_pptx])
@pytest.mark.parametrize("text,expected", CASES)
def test_segment(module, text, expected):
    assert module.segment(text) == expected


@pytest.mark.parametrize("module", [arabic_docx, arabic_pptx])
def test_segment_is_lossless(module):
    text = "نص (API) مع Array<String> و a < b و c > d، C# و .NET."
    assert "".join(chunk for chunk, _ in module.segment(text)) == text


def test_guard_wraps_only_tokens_with_affixes():
    assert arabic_docx._guard("C#") == "\u202aC#\u202c"
    assert arabic_docx._guard(".NET") == "\u202a.NET\u202c"
    assert arabic_docx._guard("PostgreSQL") == "PostgreSQL"


def test_digits_keep_latin_context():
    out = arabic_docx.arabic_digits("عام 2024 مع Nuxt 3", keep_latin_context=True)
    assert out == "عام ٢٠٢٤ مع Nuxt 3"
