"""Calculator - a clean button-based calculator built with Streamlit only.

Run with:
    streamlit run calculator.py
"""

import ast
import html
import operator
import re

import streamlit as st

st.set_page_config(page_title="Calculator", page_icon="🧮", layout="centered")

OPERATORS = ("+", "−", "×", "÷")

# Buttons, row by row. Operators and "=" get the accent colour.
LAYOUT = [
    ["C", "⌫", "%", "÷"],
    ["7", "8", "9", "×"],
    ["4", "5", "6", "−"],
    ["1", "2", "3", "+"],
    ["±", "0", ".", "="],
]
ACCENT = set(OPERATORS) | {"="}

# ----------------------------------------------------------------------------
# Styling
# ----------------------------------------------------------------------------
CSS = """
<style>
.block-container { max-width: 420px; padding-top: 2rem; }
[data-testid="stVerticalBlock"] { gap: 0.6rem; }

/* Keep the 4-column grid on small phone screens */
[data-testid="stHorizontalBlock"] { flex-wrap: nowrap !important; gap: 0.6rem; }
[data-testid="stColumn"] { min-width: 0 !important; flex: 1 1 0 !important; }

/* Display screen */
.calc-display {
    background: linear-gradient(145deg, #1e2530, #11151c);
    border-radius: 22px;
    padding: 1.2rem 1.4rem;
    text-align: right;
    box-shadow: 0 10px 28px rgba(0, 0, 0, 0.28);
    min-height: 130px;
    display: flex; flex-direction: column; justify-content: flex-end;
    overflow: hidden;
}
.calc-history { color: #8b97a8; font-size: 1.05rem; min-height: 1.5rem;
                word-break: break-all; }
.calc-main { color: #ffffff; font-weight: 600; line-height: 1.2;
             word-break: break-all; font-family: "SF Mono", Menlo, Consolas, monospace; }
.calc-main.error { color: #ff6b6b; }

/* Make every wrapper around a button fill its column */
[data-testid="stColumn"] > div,
[data-testid="stColumn"] [data-testid="stVerticalBlock"],
[data-testid="stColumn"] [data-testid="stElementContainer"],
[data-testid="stColumn"] [data-testid="stButton"],
div.stButton {
    width: 100% !important;
    min-width: 0 !important;
}

/* Buttons */
div.stButton > button {
    width: 100% !important; height: 64px; border-radius: 18px;
    padding: 0;
    transition: transform 0.08s ease, filter 0.15s ease;
}
div.stButton > button p { font-size: 1.4rem; font-weight: 600; margin: 0; }
div.stButton > button:active { transform: scale(0.95); }
div.stButton > button:hover { filter: brightness(1.08); }
div.stButton > button[kind="primary"],
div.stButton > button[data-testid="stBaseButton-primary"] {
    background: linear-gradient(135deg, #ff9f0a, #ff7a00);
    border: none; color: #ffffff;
}
div.stButton > button[kind="primary"] p,
div.stButton > button[data-testid="stBaseButton-primary"] p { color: #ffffff; }
</style>
"""

# ----------------------------------------------------------------------------
# Safe expression evaluation (no eval)
# ----------------------------------------------------------------------------
_BIN_OPS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
}


def _eval(node):
    if isinstance(node, ast.Expression):
        return _eval(node.body)
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return node.value
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
        return -_eval(node.operand)
    if isinstance(node, ast.BinOp) and type(node.op) in _BIN_OPS:
        return _BIN_OPS[type(node.op)](_eval(node.left), _eval(node.right))
    raise ValueError("Unsupported expression")


def format_number(value: float) -> str:
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    return f"{value:.10g}"


def calculate(expr: str) -> str:
    cleaned = expr.rstrip("+−×÷")  # ignore a dangling operator, e.g. "5+"
    cleaned = cleaned.replace("×", "*").replace("÷", "/").replace("−", "-")
    return format_number(_eval(ast.parse(cleaned, mode="eval")))


# ----------------------------------------------------------------------------
# State and button handling
# ----------------------------------------------------------------------------
def init_state() -> None:
    st.session_state.setdefault("expr", "")
    st.session_state.setdefault("history", "")
    st.session_state.setdefault("done", False)  # last action was "="
    st.session_state.setdefault("error", "")


def reset() -> None:
    s = st.session_state
    s.expr, s.history, s.done, s.error = "", "", False, ""


def press(key: str) -> None:
    s = st.session_state

    if s.error:  # start fresh after an error
        reset()
    if key == "C":
        reset()
        return

    if key == "⌫":
        if s.done:
            reset()
        elif s.expr.endswith(")"):
            s.expr = re.sub(r"\(-[\d.]+\)$", "", s.expr)
        else:
            s.expr = s.expr[:-1]
        return

    if key.isdigit():
        if s.done:
            s.expr, s.history, s.done = "", "", False
        if s.expr.endswith(")"):
            return
        segment = re.search(r"[\d.]*$", s.expr).group()
        s.expr = s.expr[:-1] + key if segment == "0" else s.expr + key
        return

    if key == ".":
        if s.done:
            s.expr, s.history, s.done = "", "", False
        if s.expr.endswith(")"):
            return
        segment = re.search(r"[\d.]*$", s.expr).group()
        if "." in segment:
            return
        s.expr += "0." if segment == "" else "."
        return

    if key in OPERATORS:
        s.done, s.history = False, ""
        if s.expr == "":
            if key == "−":
                s.expr = "-"
            return
        if s.expr == "-":
            return
        if s.expr[-1] in OPERATORS:
            s.expr = s.expr[:-1] + key
        else:
            s.expr += key
        return

    if key == "%":
        match = re.search(r"\d+\.?\d*$|\.\d+$", s.expr)
        if match:
            s.done, s.history = False, ""
            s.expr = s.expr[: match.start()] + format_number(float(match.group()) / 100)
        return

    if key == "±":
        if not s.expr or s.expr == "-":
            return
        s.done, s.history = False, ""
        if re.fullmatch(r"-[\d.]+", s.expr):
            s.expr = s.expr[1:]
        elif re.fullmatch(r"[\d.]+", s.expr):
            s.expr = "-" + s.expr
        elif s.expr.endswith(")"):
            s.expr = re.sub(r"\(-([\d.]+)\)$", r"\1", s.expr)
        else:
            match = re.search(r"[\d.]+$", s.expr)
            if match:
                s.expr = s.expr[: match.start()] + f"(-{match.group()})"
        return

    if key == "=":
        if not s.expr or s.done:
            return
        try:
            result = calculate(s.expr)
        except ZeroDivisionError:
            s.error = "Cannot divide by 0"
            return
        except Exception:
            s.error = "Invalid input"
            return
        s.history = s.expr.rstrip("+−×÷") + " ="
        s.expr, s.done = result, True


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
def render_display() -> None:
    s = st.session_state
    text = s.error or s.expr or "0"
    length = len(text)
    size = 3.0 if length <= 8 else 2.2 if length <= 12 else 1.6 if length <= 18 else 1.2
    css_class = "calc-main error" if s.error else "calc-main"
    st.markdown(
        f"""
        <div class="calc-display">
            <div class="calc-history">{html.escape(s.history) or "&nbsp;"}</div>
            <div class="{css_class}" style="font-size:{size}rem">{html.escape(text)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_buttons() -> None:
    for r, row in enumerate(LAYOUT):
        cols = st.columns(4)
        for c, label in enumerate(row):
            cols[c].button(
                label,
                key=f"btn_{r}_{c}",
                type="primary" if label in ACCENT else "secondary",
                on_click=press,
                args=(label,),
            )


def main() -> None:
    st.markdown(CSS, unsafe_allow_html=True)
    init_state()
    st.markdown("### 🧮 Calculator")
    render_display()
    render_buttons()
    st.caption("Supports + − × ÷, %, ±, decimals and backspace.")


main()