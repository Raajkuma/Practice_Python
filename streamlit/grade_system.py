"""Grade System - a small Streamlit app that converts a mark (0-100) to a letter grade.

Run with:
    streamlit run grade_system.py
"""

import math

import streamlit as st

# Grade bands: (minimum mark, letter grade), checked from highest to lowest.
GRADE_BANDS = [
    (90, "A"),
    (80, "B"),
    (70, "C"),
    (60, "D"),
    (0, "F"),
]


def get_grade(mark: float) -> str:
    """Return the letter grade for a mark between 0 and 100."""
    for minimum, grade in GRADE_BANDS:
        if mark >= minimum:
            return grade
    raise ValueError("Mark must be between 0 and 100.")


def parse_mark(text: str):
    """Convert user text to a float. Returns (mark, error_message)."""
    text = text.strip()
    if not text:
        return None, None  # empty: not an error, just waiting for input
    try:
        mark = float(text)
    except ValueError:
        return None, "Please enter a valid number (for example 76 or 82.5)."
    if math.isnan(mark) or math.isinf(mark):
        return None, "Please enter a valid number (for example 76 or 82.5)."
    if mark < 0 or mark > 100:
        return None, "The mark must be between 0 and 100."
    return mark, None


def main() -> None:
    st.set_page_config(page_title="Grade System", page_icon="🎓")
    st.title("🎓 Grade System")
    st.write("Enter a mark between 0 and 100 to see the matching letter grade.")

    text = st.text_input("Mark (0-100)", placeholder="e.g. 76")
    mark, error = parse_mark(text)

    if error:
        st.error(error)
    elif mark is None:
        st.info("Waiting for a mark. Type a number above and press Enter.")
    else:
        shown = int(mark) if mark == int(mark) else mark
        st.success(f"Mark: {shown} → Grade: {get_grade(mark)}")

    with st.expander("Grade bands"):
        st.markdown(
            "| Mark | Grade |\n|---|---|\n"
            "| 90 – 100 | A |\n"
            "| 80 – 89.99 | B |\n"
            "| 70 – 79.99 | C |\n"
            "| 60 – 69.99 | D |\n"
            "| 0 – 59.99 | F |"
        )


main()