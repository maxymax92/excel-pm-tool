"""Workbook capacity contract.

These are technical package limits, not business settings. Every bounded
validation range, Calc helper and view spill uses the same constants so every
accepted record remains represented throughout the workbook.
"""

# Items and RAID records supported by all validation, calculation and view
# layers. The data tables may physically grow further, so visible sheet-level
# warnings must identify a breached limit before a user trusts a partial view.
DATA_ROWS = 2000

# Editable Config list rows covered by dropdown sources, paste-safety rules and
# role-helper spills. The capacity accommodates substantial taxonomy growth.
CONFIG_ROWS = 500

# Plan renders every supported operational record across its 52-week grid.
PLAN_ROWS = DATA_ROWS
PLAN_WEEKS = 52
