"""Pure, tested rules for the pharmacy attribution model.

This module contains no source-file or network side effects.
"""
from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
import re
import unicodedata

import numpy as np
import pandas as pd

AGE_COLUMNS = ["pop_65_69", "pop_70_74", "pop_75_79", "pop_80_84", "pop_85_mas"]


def code(value, width: int) -> str | None:
    """Normalize genuinely numeric codes; missing/NULL keys never become zero."""
    if value is None or pd.isna(value):
        return None
    s = str(value).strip()
    if not re.fullmatch(r"\d+(?:\.0+)?", s):
        return None
    s = s.split(".")[0].lstrip("0") or "0"
    return s.zfill(width) if s != "0" and len(s) <= width else None


def text_key(value) -> str:
    if value is None or pd.isna(value):
        return ""
    s = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]+", " ", s).strip()


def cents(value) -> int:
    return int((Decimal(str(value)) * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def allocate_cents(budget: int, weights: np.ndarray) -> np.ndarray:
    """Largest remainders with stable input-order ties, preserving every cent."""
    weights = np.asarray(weights, dtype=float)
    if budget < 0 or weights.ndim != 1 or len(weights) == 0:
        raise ValueError("Nonnegative budget and nonempty one-dimensional weights required")
    if not np.isfinite(weights).all() or (weights < 0).any() or weights.sum() <= 0:
        raise ValueError("Weights must be finite, nonnegative and have positive total")
    raw = budget * (weights / weights.sum())
    result = np.floor(raw).astype(np.int64)
    remainder = budget - int(result.sum())
    order = np.argsort(-(raw - result), kind="stable")
    result[order[:remainder]] += 1
    assert result.sum() == budget
    return result


def age_demand(population: np.ndarray, weights=(1., 2., 3.)) -> np.ndarray:
    a, b, c = weights
    return np.asarray(population) @ np.array([a, a, b, b, c])


def haversine_matrix(demand_lonlat: np.ndarray, store_lonlat: np.ndarray) -> np.ndarray:
    """Great-circle metres, including stable coincident points."""
    x = np.radians(demand_lonlat)
    y = np.radians(store_lonlat)
    dlat = x[:, 1, None] - y[None, :, 1]
    dlon = x[:, 0, None] - y[None, :, 0]
    a = np.sin(dlat / 2.) ** 2 + np.cos(x[:, 1, None]) * np.cos(y[None, :, 1]) * np.sin(dlon / 2.) ** 2
    return 2. * 6371008.8 * np.arcsin(np.sqrt(np.clip(a, 0., 1.)))


def distance_shares(distances: np.ndarray, offsets: np.ndarray,
                    placement_probabilities: np.ndarray, demand_probabilities: np.ndarray,
                    exponent=2., distance_offset=100.) -> np.ndarray:
    """Average attraction over uncertain store locations, then normalize per origin.

    offsets mark the first placement of each unique candidate establishment.
    A store with eight possible locations remains ONE competitor.
    """
    attraction = (distances + distance_offset) ** (-exponent)
    attraction *= placement_probabilities[None, :]
    expected = np.add.reduceat(attraction, offsets, axis=1)
    expected /= expected.sum(axis=1, keepdims=True)
    shares = np.asarray(demand_probabilities) @ expected
    return shares / shares.sum()


def categories(stores: pd.DataFrame, budgets: pd.Series, high_quantile=.75,
               cutoffs=(.2, .5, .75)) -> pd.DataFrame:
    """Categories on physical establishments, with unchanged underlying euros."""
    out = stores.copy()
    out["categoria_base"] = "D"
    out["pct_acumulado_previo_ccaa"] = 0.
    for _, group in out.groupby("ccaa_modelo", dropna=False, sort=True):
        ordered = group.sort_values(["potential_cents", "establishment_id"], ascending=[False, True], kind="stable")
        total = int(ordered.potential_cents.sum())
        previous = (ordered.potential_cents.cumsum() - ordered.potential_cents) / total if total else pd.Series(1., index=ordered.index)
        values = np.select([previous < x for x in cutoffs], ["A", "B", "C"], default="D")
        out.loc[ordered.index, "categoria_base"] = values
        out.loc[ordered.index, "pct_acumulado_previo_ccaa"] = previous
    threshold = float(budgets.quantile(high_quantile))
    out["sb_alto"] = out.sb_modelo.map(budgets).ge(threshold)
    out["categoria_final"] = out.categoria_base
    out["motivo_ajuste_categoria"] = ""
    upgraded = out.sb_alto & out.categoria_final.eq("D")
    out.loc[upgraded, "categoria_final"] = "C"
    out.loc[upgraded, "motivo_ajuste_categoria"] = "Minimo C: SB de ubicacion >= P75 nacional"
    for sb, group in out[out.sb_alto].groupby("sb_modelo", sort=True):
        if group.categoria_final.eq("A").all():
            selected = group.sort_values(["potential_cents", "establishment_id"], kind="stable").index[0]
            out.loc[selected, "categoria_final"] = "B"
            out.loc[selected, "motivo_ajuste_categoria"] = "SB alto integramente A: menor potencial pasa a B"
    assert not (out.sb_alto & out.categoria_final.eq("D")).any()
    assert not out[out.sb_alto].groupby("sb_modelo").categoria_final.apply(lambda g: g.eq("A").all()).any()
    out.attrs["high_sb_threshold_cents"] = threshold
    return out
