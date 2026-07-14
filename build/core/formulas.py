"""Formula encoding authority for cells and defined LAMBDA names.

``F`` owns function prefixes, LET/LAMBDA parameter namespaces and spill
encoding. ``LAM`` emits fully encoded defined names. Optional LAMBDA arguments
are represented by omitted trailing arguments and ``ISOMITTED`` checks.

Excel stores supported dynamic functions with ``_xlfn`` prefixes, FILTER and
SORT with the additional ``_xlws`` namespace, LET/LAMBDA variables with
``_xlpm``, and spill references as ``ANCHORARRAY``. Every formula is encoded
here because XlsxWriter does not apply these rules consistently to defined
names and manually prefixed formulas.
"""

import re
from collections.abc import Collection, Iterator, Sequence
from enum import Enum

# Functions stored with the ``_xlfn`` prefix. This set covers XlsxWriter's
# supported future functions plus workbook functions absent from its table.
MODERN = {
    "ACOT",
    "ACOTH",
    "AGGREGATE",
    "ANCHORARRAY",
    "ARABIC",
    "ARRAYTOTEXT",
    "BASE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.DIST.RANGE",
    "BINOM.INV",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "BYCOL",
    "BYROW",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "COMBINA",
    "CONCAT",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COT",
    "COTH",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "DROP",
    "ERF.PRECISE",
    "ERFC.PRECISE",
    "EXPAND",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FILTERXML",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "FORMULATEXT",
    "GAMMA",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "GAUSS",
    "HSTACK",
    "HYPGEOM.DIST",
    "IFNA",
    "IFS",
    "IMAGE",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMSEC",
    "IMSECH",
    "IMSINH",
    "IMTAN",
    "ISFORMULA",
    "ISOMITTED",
    "ISOWEEKNUM",
    "LAMBDA",
    "LET",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MAKEARRAY",
    "MAP",
    "MAXIFS",
    "MINIFS",
    "MODE.MULT",
    "MODE.SNGL",
    "MUNIT",
    "NEGBINOM.DIST",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "NUMBERVALUE",
    "PDURATION",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "PERMUTATIONA",
    "PHI",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "QUERYSTRING",
    "RANDARRAY",
    "RANK.AVG",
    "RANK.EQ",
    "REDUCE",
    "RRI",
    "SCAN",
    "SEC",
    "SECH",
    "SEQUENCE",
    "SHEET",
    "SHEETS",
    "SKEW.P",
    "SORTBY",
    "STDEV.P",
    "STDEV.S",
    "SWITCH",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTJOIN",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "UNICHAR",
    "UNICODE",
    "UNIQUE",
    "VALUETOTEXT",
    "VAR.P",
    "VAR.S",
    "VSTACK",
    "WEBSERVICE",
    "WEIBULL.DIST",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
    "XOR",
    "Z.TEST",
    "SINGLE",
    # Functions supported by the target Excel release.
    "GROUPBY",
    "PIVOTBY",
    "PERCENTOF",
    "REGEXTEST",
    "REGEXEXTRACT",
    "REGEXREPLACE",
}
# modern functions that need the extra _xlws. namespace (worksheet functions)
MODERN_XLWS = {"FILTER", "SORT"}

# classic functions: stored bare (whitelist for the linter)
CLASSIC = {
    "SUM",
    "SUMIF",
    "SUMIFS",
    "SUMPRODUCT",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "COUNTIF",
    "COUNTIFS",
    "AVERAGE",
    "AVERAGEIF",
    "AVERAGEIFS",
    "MIN",
    "MAX",
    "MEDIAN",
    "LARGE",
    "SMALL",
    "RANK",
    "IF",
    "AND",
    "OR",
    "NOT",
    "IFERROR",
    "ISNUMBER",
    "ISBLANK",
    "ISTEXT",
    "ISERROR",
    "ISLOGICAL",
    "N",
    "T",
    "NA",
    "TYPE",
    "INDEX",
    "MATCH",
    "ROWS",
    "COLUMNS",
    "ROW",
    "COLUMN",
    "CHOOSE",
    "LOOKUP",
    "TRANSPOSE",
    "HYPERLINK",
    "ADDRESS",
    "TODAY",
    "DATE",
    "DATEVALUE",
    "YEAR",
    "MONTH",
    "DAY",
    "WEEKDAY",
    "EDATE",
    "EOMONTH",
    "NETWORKDAYS",
    "NETWORKDAYS.INTL",
    "WORKDAY",
    "WORKDAY.INTL",
    "DATEDIF",
    "YEARFRAC",
    "TEXT",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "FIND",
    "SEARCH",
    "SUBSTITUTE",
    "TRIM",
    "UPPER",
    "LOWER",
    "PROPER",
    "REPT",
    "CHAR",
    "CODE",
    "VALUE",
    "EXACT",
    "ABS",
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "INT",
    "MOD",
    "MROUND",
    "CEILING",
    "FLOOR",
    "SIGN",
    "SQRT",
    "SUBTOTAL",
}
BANNED = {"INDIRECT", "OFFSET", "NOW", "RAND", "RANDBETWEEN", "TRIMRANGE", "VLOOKUP", "HLOOKUP"}

_TOKEN = re.compile(r"(?<![A-Za-z0-9_.])([A-Z][A-Z0-9.]*)\s*\(")
_STRING = re.compile(r'"(?:[^"]|"")*"')
_SPILL = re.compile(r"\b([A-Za-z]{1,3}\$?\d+)#")


class _FormulaProblem(Enum):
    UNEXPECTED_CLOSING = "unexpected closing parenthesis in {!r}"
    UNTERMINATED_STRING = "unterminated string literal in {!r}"
    UNTERMINATED_NAME = "unterminated quoted sheet or workbook name in {!r}"
    UNCLOSED_PARENTHESIS = "unclosed parenthesis in {!r}"
    BANNED_FUNCTION = "banned function {} in {!r}"
    UNCLASSIFIED_FUNCTION = (
        "unclassified function {!r} in {!r} - add it to MODERN or CLASSIC in formulas.py"
    )
    DOUBLE_PREFIX = "double-prefixed param in {!r}"
    INVALID_VARIABLE_CASE = "LAMBDA/LET vars must be lowercase: {!r}"


class _FormulaError(ValueError):
    def __init__(self, problem: _FormulaProblem, *details: object) -> None:
        super().__init__(problem.value.format(*details))


def _split_strings(formula: str) -> Iterator[tuple[bool, str]]:
    """Yield string-literal and formula-token chunks for targeted rewriting.

    Yields:
        Tuples containing a string-literal flag and the corresponding chunk.

    """
    pos = 0
    for match in _STRING.finditer(formula):
        if match.start() > pos:
            yield False, formula[pos : match.start()]
        yield True, match.group(0)
        pos = match.end()
    if pos < len(formula):
        yield False, formula[pos:]


def _prefix_functions(chunk: str) -> str:
    def replace_token(match: re.Match[str]) -> str:
        name = match.group(1)
        if name in MODERN_XLWS:
            return f"_xlfn._xlws.{name}("
        if name in MODERN:
            return f"_xlfn.{name}("
        return match.group(0)

    return _TOKEN.sub(replace_token, chunk)


def _sub_vars(chunk: str, variables: Collection[str]) -> str:
    rewritten = chunk
    for variable in sorted(variables, key=len, reverse=True):
        rewritten = re.sub(
            rf"(?<![\w.\[])\b{re.escape(variable)}\b(?!\s*\()",
            f"_xlpm.{variable}",
            rewritten,
        )
    return rewritten


def _quoted_segment_end(formula: str, start: int, quote: str) -> int:
    index = start + 1
    while index < len(formula):
        if formula[index] != quote:
            index += 1
            continue
        if index + 1 < len(formula) and formula[index + 1] == quote:
            index += 2
            continue
        return index + 1

    problem = (
        _FormulaProblem.UNTERMINATED_STRING if quote == '"' else _FormulaProblem.UNTERMINATED_NAME
    )
    raise _FormulaError(problem, formula)


def _validate_syntax(formula: str) -> None:
    """Validate quote termination and parenthesis balance before encoding.

    Raises:
        _FormulaError: If a quote is unterminated or parentheses are unbalanced.

    """
    depth = 0
    index = 0
    while index < len(formula):
        char = formula[index]
        if char in {'"', "'"}:
            index = _quoted_segment_end(formula, index, char)
            continue
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth < 0:
                raise _FormulaError(_FormulaProblem.UNEXPECTED_CLOSING, formula)
        index += 1
    if depth:
        raise _FormulaError(_FormulaProblem.UNCLOSED_PARENTHESIS, formula)


def _lint_formula(stored: str, original: str) -> str:
    """Reject banned or unclassified function tokens.

    Returns:
        The validated stored formula.

    Raises:
        _FormulaError: If the formula uses a banned or unclassified function or a
            parameter has already been prefixed.

    """
    for is_str, chunk in _split_strings(original):
        if is_str:
            continue
        for match in _TOKEN.finditer(chunk):
            name = match.group(1)
            if name in BANNED:
                raise _FormulaError(_FormulaProblem.BANNED_FUNCTION, name, original)
            if (
                name not in MODERN
                and name not in MODERN_XLWS
                and name not in CLASSIC
                and not name.startswith("_xl")
                and not name.startswith("fn")
            ):
                raise _FormulaError(
                    _FormulaProblem.UNCLASSIFIED_FUNCTION,
                    name,
                    original,
                )
    if "_xlfn._xlpm." in stored:
        raise _FormulaError(_FormulaProblem.DOUBLE_PREFIX, stored)
    return stored


def encode_formula(formula: str, variables: Collection[str] = ()) -> str:
    """Encode a cell formula for storage in an OOXML workbook.

    Invalid syntax and banned or unclassified function tokens raise
    :class:`ValueError`.

    Returns:
        The formula with future-function, parameter and spill prefixes encoded.

    """
    _validate_syntax(formula)
    out: list[str] = []
    for is_str, chunk in _split_strings(formula):
        if is_str:
            out.append(chunk)
            continue
        rewritten = _SPILL.sub(r"_xlfn.ANCHORARRAY(\1)", chunk)
        rewritten = _prefix_functions(rewritten)
        if variables:
            rewritten = _sub_vars(rewritten, variables)
        out.append(rewritten)
    stored = "".join(out)
    return _lint_formula(stored, formula)


def _validate_variable_case(variables: Collection[str]) -> None:
    """Require lowercase LET and LAMBDA variable names.

    Raises:
        _FormulaError: If a variable contains uppercase characters.

    """
    for variable in variables:
        if variable != variable.lower():
            raise _FormulaError(_FormulaProblem.INVALID_VARIABLE_CASE, variable)


def encode_lambda(
    params: Sequence[str],
    body: str,
    extra_variables: Sequence[str] = (),
) -> str:
    """Encode a LAMBDA defined name and its parameter namespace.

    A trailing '?' marks a parameter the caller may omit. Excel LAMBDA has
    NO optional-parameter *declaration* syntax (no brackets) — every
    parameter is declared plainly; a caller simply passes fewer arguments
    and ISOMITTED() detects the omission. So '?' only affects our bookkeeping,
    never the emitted text. Invalid variable names and formula bodies raise
    :class:`ValueError`.

    Returns:
        The fully encoded OOXML defined-name formula.

    """
    names = [parameter.removesuffix("?") for parameter in params]
    variables = tuple(names) + tuple(extra_variables)
    _validate_variable_case(variables)
    body_enc = encode_formula(body, variables=variables)
    if body_enc.startswith("="):
        body_enc = body_enc.removeprefix("=")
    plist = ",".join(f"_xlpm.{n}" for n in names)
    return f"=_xlfn.LAMBDA({plist},{body_enc})"
