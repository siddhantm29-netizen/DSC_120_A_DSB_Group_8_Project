"""
dsb_utils — shared utilities for the DSB Group 8 assignment.

A small, dependency-light module so every notebook stays DRY and consistent:
    * robust loaders for the commodity CSVs and the SWIFT tracker,
    * one matplotlib style applied everywhere so charts look uniform,
    * save_fig / save_table helpers that write to a tidy figures/ + tables/ layout.

Design notes
------------
* Uses **relative paths only**. ``repo_root()`` walks up the tree from this
  file until it finds the data files, so notebooks reproduce on any machine
  regardless of the working directory.
* The commodity files use *different* date separators — Brent uses ``MM-DD-YYYY``
  (e.g. ``03-01-2026``) while Gold/Silver use ``MM/DD/YYYY``. The first field is
  always the **month**, the day component is always ``01`` (monthly series).
  ``load_commodity`` normalises both before parsing so no file is misread.
"""

from __future__ import annotations

import os
import pandas as pd
import matplotlib.pyplot as plt

# --------------------------------------------------------------------------- #
# Repository / data discovery (relative paths only)
# --------------------------------------------------------------------------- #
COMMODITY_FILES = {
    "Brent Oil": "Brent Oil.csv",
    "Gold": "Gold 100years.csv",
    "Silver": "silver 100 years.csv",
}
SWIFT_FILE = "swift_currency_tracker_all_reports.csv"

#: The most recent SWIFT report month in the dataset (single source of truth).
RECENT_REPORT = "April 2026"

#: Marker file used to locate the repository root.
_ROOT_MARKER = "Brent Oil.csv"


def repo_root(start: str | None = None) -> str:
    """Return the repository root by walking up until the data files appear.

    Falls back to the parent of this file's directory (``src/``) so the module
    still resolves sensibly even if the marker is renamed.
    """
    here = os.path.abspath(start or os.path.dirname(__file__))
    d = here
    while True:
        if os.path.exists(os.path.join(d, _ROOT_MARKER)):
            return d
        parent = os.path.dirname(d)
        if parent == d:  # reached filesystem root without a hit
            return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        d = parent


#: Resolved once at import time; the directory that holds the CSV datasets.
ROOT = repo_root()


def data_path(filename: str) -> str:
    """Absolute path to a data file living in the repository root."""
    return os.path.join(ROOT, filename)


# --------------------------------------------------------------------------- #
# Loaders
# --------------------------------------------------------------------------- #
def _parse_monthly_dates(raw: pd.Series) -> pd.Series:
    """Parse a month-first ``MM?DD?YYYY`` series regardless of separator.

    Both ``01/01/1915`` and ``03-01-2026`` are normalised to ``MM/DD/YYYY``
    and parsed with an explicit format so pandas never swaps day and month.
    """
    cleaned = (
        raw.astype(str)
        .str.strip()
        .str.strip('"')
        .str.replace("-", "/", regex=False)
    )
    return pd.to_datetime(cleaned, format="%m/%d/%Y")


def load_commodity(path_or_name: str) -> pd.DataFrame:
    """Load a commodity CSV → cleaned, dated, sorted DataFrame.

    Accepts either a friendly name (``"Brent Oil"``, ``"Gold"``, ``"Silver"``),
    a bare filename, or a full path. Adds ``Year`` and ``Month`` helper columns.
    """
    if os.path.exists(path_or_name):
        path = path_or_name
    elif path_or_name in COMMODITY_FILES:
        path = data_path(COMMODITY_FILES[path_or_name])
    else:
        path = data_path(path_or_name)

    df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = [c.strip().strip('"') for c in df.columns]
    df["Date"] = _parse_monthly_dates(df["Date"])
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month
    return df.sort_values("Date").reset_index(drop=True)


def load_all_commodities() -> dict[str, pd.DataFrame]:
    """Return ``{name: DataFrame}`` for Brent Oil, Gold and Silver."""
    return {name: load_commodity(name) for name in COMMODITY_FILES}


def load_swift(path: str | None = None) -> pd.DataFrame:
    """Load the SWIFT currency tracker with ``value`` coerced to numeric."""
    df = pd.read_csv(path or data_path(SWIFT_FILE))
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    return df


# --------------------------------------------------------------------------- #
# SWIFT convenience helpers
# --------------------------------------------------------------------------- #
def _metric(swift: pd.DataFrame, metric: str, report_month: str) -> pd.DataFrame:
    d = swift[(swift["metric"] == metric) & (swift["report_month"] == report_month)]
    return d.sort_values("value", ascending=False).reset_index(drop=True)


def global_payment_share(swift: pd.DataFrame, report_month: str = RECENT_REPORT) -> pd.DataFrame:
    """Global Payment Share ranking for a given report month (descending)."""
    return _metric(swift, "Global Payment Share", report_month)


def trade_finance_share(swift: pd.DataFrame, report_month: str = RECENT_REPORT) -> pd.DataFrame:
    """Trade Finance Share ranking for a given report month (descending)."""
    return _metric(swift, "Trade Finance Share", report_month)


def offshore_rmb(swift: pd.DataFrame, report_month: str = RECENT_REPORT) -> pd.DataFrame:
    """Offshore RMB by Economy ranking for a given report month (descending)."""
    return _metric(swift, "Offshore RMB by Economy", report_month)


# --------------------------------------------------------------------------- #
# Plot style + IO helpers
# --------------------------------------------------------------------------- #
#: One palette reused across notebooks so commodities/currencies keep one colour.
COLORS = {
    "oil": "#1f77b4",      # blue
    "gold": "#d4a017",     # goldenrod
    "silver": "#7f8c9b",   # slate
    "usd": "#2e8b57",      # sea green
    "eur": "#7b68ee",      # medium slate
    "cny": "#d9534f",      # red
    "accent": "#17a2b8",   # teal
}

_STYLE = {
    "figure.figsize": (10, 4.5),
    "figure.dpi": 110,
    "savefig.dpi": 150,
    "savefig.bbox": "tight",
    "axes.grid": True,
    "grid.alpha": 0.35,
    "grid.linestyle": "--",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.titlesize": 13,
    "axes.titleweight": "bold",
    "axes.labelsize": 11,
    "axes.axisbelow": True,
    "font.size": 10,
    "legend.frameon": False,
}


def apply_style() -> None:
    """Apply the shared matplotlib style (call once near the top of a notebook)."""
    plt.rcParams.update(_STYLE)


def _ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


def save_fig(fig, name: str, qdir: str) -> str:
    """Save ``fig`` to ``<qdir>/figures/<name>.png`` and return the path."""
    out = os.path.join(_ensure_dir(os.path.join(qdir, "figures")), f"{name}.png")
    fig.savefig(out)
    print(f"saved figure -> {os.path.relpath(out, ROOT)}")
    return out


def save_table(df: pd.DataFrame, name: str, qdir: str, index: bool = False) -> str:
    """Save ``df`` to ``<qdir>/tables/<name>.csv`` and return the path."""
    out = os.path.join(_ensure_dir(os.path.join(qdir, "tables")), f"{name}.csv")
    df.to_csv(out, index=index)
    print(f"saved table  -> {os.path.relpath(out, ROOT)}")
    return out
