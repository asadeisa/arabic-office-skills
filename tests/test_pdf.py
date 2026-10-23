"""shape() output is in visual order: what it returns is what the page shows."""
import pytest

from arabic_pdf import shape

SIGNED = ["-1.9", "−0.42", "+0.0224", "±0.5"]


@pytest.mark.parametrize("number", SIGNED)
def test_sign_stays_left_of_a_bare_number(number):
    assert shape(number) == number


def test_sign_stays_left_inside_arabic():
    assert "−0.42" in shape("الفرق −0.42 لصالح النموذج")


@pytest.mark.parametrize("text", ["انتهت 1-0 للمضيف", "بين 2024-01-01 و 2026-09-19"])
def test_hyphen_between_digits_is_untouched(text):
    for token in text.split():
        if token[0].isdigit():
            assert token in shape(text)
