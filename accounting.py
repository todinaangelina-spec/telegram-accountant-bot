"""Exact, safe accounting calculations for the Russian-language assistant."""
from __future__ import annotations

import ast
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
import re


# Spaces may be used only as three-digit grouping separators.  Commas and dots
# are decimal separators, never grouping separators, so ``1,234.56`` remains
# deliberately invalid rather than ambiguous.
_MONEY_RE = re.compile(
    r"^(?P<sign>[+-]?)(?P<integer>(?:\d{1,3}(?: \d{3})+)|\d+)(?:(?P<separator>[,.])(?P<fraction>\d{1,2}))?$"
)
_NUMBER_LITERAL_RE = re.compile(r"^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")

MAX_EXPRESSION_LENGTH = 200
MAX_AST_NODES = 100
MAX_ABS_VALUE = Decimal("1E+18")
MIN_ADJUSTED_EXPONENT = -100
_CALC_PRECISION = 256


def _as_decimal(value: Decimal, label: str) -> Decimal:
    if isinstance(value, bool) or isinstance(value, float):
        raise ValueError(f"{label} должен быть точным Decimal, а не float")
    try:
        decimal_value = value if isinstance(value, Decimal) else Decimal(value)
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise ValueError(f"Некорректное значение: {label}") from exc
    if not decimal_value.is_finite():
        raise ValueError(f"{label} должен быть конечным числом")
    return decimal_value


def parse_money(text: str) -> Decimal:
    """Parse an unambiguous money amount without rounding.

    Both ``1 234,50`` and ``1234.50`` are accepted.  A comma or dot may appear
    once as a decimal separator; comma/dot thousands notation is rejected.
    """
    if not isinstance(text, str):
        raise ValueError("Сумма должна быть строкой")
    # Russian keyboards commonly emit NBSP/narrow NBSP when grouping numbers.
    normalized = text.strip().replace("\u00a0", " ").replace("\u202f", " ")
    if len(normalized) > 64:
        raise ValueError("Слишком длинное число")
    match = _MONEY_RE.fullmatch(normalized)
    if match is None:
        raise ValueError("Некорректный денежный формат")
    number = f"{match.group('sign')}{match.group('integer').replace(' ', '')}"
    if match.group("fraction") is not None:
        number += "." + match.group("fraction")
    try:
        result = Decimal(number)
    except InvalidOperation as exc:  # Defensive; the regex already constrains it.
        raise ValueError("Некорректный денежный формат") from exc
    if not result.is_finite():
        raise ValueError("Сумма должна быть конечной")
    if abs(result) > MAX_ABS_VALUE:
        raise ValueError("Число по модулю не должно превышать 10^18")
    return result


def money_text(amount: Decimal) -> str:
    """Format a Decimal as Russian money text, rounded HALF_UP to kopecks."""
    value = _as_decimal(amount, "Сумма")
    with localcontext() as context:
        context.prec = max(28, value.adjusted() + 4)
        rounded = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if rounded == 0:
        # Do not display a confusing negative zero for tiny negative values.
        rounded = Decimal("0.00")
    return format(rounded, ",.2f").replace(",", " ").replace(".", ",")


def _ensure_magnitude(value: Decimal) -> Decimal:
    if not value.is_finite():
        raise ValueError("Результат должен быть конечным")
    if value != 0 and (abs(value) > MAX_ABS_VALUE or value.adjusted() < MIN_ADJUSTED_EXPONENT):
        raise ValueError("Выражение выходит за допустимые пределы")
    return value


def _decimal_literal(expression: str, node: ast.Constant) -> Decimal:
    # Reading the original source token avoids routing a numeric literal through
    # Python's binary float representation (e.g. 0.1 stays exact Decimal('0.1')).
    source = ast.get_source_segment(expression, node)
    if source is None or _NUMBER_LITERAL_RE.fullmatch(source) is None:
        raise ValueError("Разрешены только десятичные числовые литералы")
    try:
        value = Decimal(source)
    except InvalidOperation as exc:
        raise ValueError("Некорректный числовой литерал") from exc
    return _ensure_magnitude(value)


def safe_calculate(expression: str) -> Decimal:
    """Evaluate only arithmetic grammar using Decimal, never ``eval``.

    Allowed syntax is decimal literals, parentheses, unary ``+``/``-`` and
    binary ``+``, ``-``, ``*`` and ``/``.  Length, AST size and intermediate
    magnitude are bounded before returning a result.
    """
    if not isinstance(expression, str) or not expression.strip():
        raise ValueError("Введите арифметическое выражение")
    if len(expression) > MAX_EXPRESSION_LENGTH:
        raise ValueError(f"Выражение не должно превышать {MAX_EXPRESSION_LENGTH} символов")
    expression = expression.strip()
    try:
        tree = ast.parse(expression, mode="eval")
    except SyntaxError as exc:
        raise ValueError("Некорректное арифметическое выражение") from exc
    if sum(1 for _ in ast.walk(tree)) > MAX_AST_NODES:
        raise ValueError(f"Выражение не должно содержать более {MAX_AST_NODES} узлов")

    def calculate(node: ast.AST) -> Decimal:
        if isinstance(node, ast.Expression):
            return calculate(node.body)
        if isinstance(node, ast.Constant):
            # bool is a subclass of int, so source validation alone is not
            # enough for an AST crafted programmatically.
            if isinstance(node.value, bool) or not isinstance(node.value, (int, float)):
                raise ValueError("Разрешены только числа")
            return _decimal_literal(expression, node)
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            value = calculate(node.operand)
            return _ensure_magnitude(value if isinstance(node.op, ast.UAdd) else -value)
        if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
            left = calculate(node.left)
            right = calculate(node.right)
            if isinstance(node.op, ast.Add):
                result = left + right
            elif isinstance(node.op, ast.Sub):
                result = left - right
            elif isinstance(node.op, ast.Mult):
                result = left * right
            else:
                if right == 0:
                    raise ValueError("Деление на ноль")
                result = left / right
            return _ensure_magnitude(result)
        raise ValueError("Разрешены только +, -, *, /, скобки и числа")

    try:
        with localcontext() as context:
            # Source length is capped at 200, so 256 digits preserves all input
            # digits and practical intermediate results without unbounded work.
            context.prec = _CALC_PRECISION
            return calculate(tree)
    except (InvalidOperation, OverflowError, ZeroDivisionError) as exc:
        raise ValueError("Невозможно вычислить выражение") from exc


def vat(amount: Decimal, rate: Decimal) -> dict[str, Decimal]:
    """Calculate VAT where ``rate`` is a percentage (for example, ``20``).

    ``amount`` is interpreted as VAT-inclusive for ``base`` and
    ``included_tax``, and as VAT-exclusive for ``added_tax``.
    """
    decimal_amount = _as_decimal(amount, "Сумма")
    decimal_rate = _as_decimal(rate, "Ставка НДС")
    if decimal_rate < 0:
        raise ValueError("Ставка НДС не может быть отрицательной")
    with localcontext() as context:
        context.prec = _CALC_PRECISION
        multiplier = Decimal(1) + decimal_rate / Decimal(100)
        base = decimal_amount / multiplier
        included_tax = decimal_amount - base
        added_tax = decimal_amount * decimal_rate / Decimal(100)
        total_with_tax = decimal_amount + added_tax
    return {
        "base": base,
        "included_tax": included_tax,
        "added_tax": added_tax,
        "total_with_tax": total_with_tax,
    }
