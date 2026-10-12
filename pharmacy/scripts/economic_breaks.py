"""Exact, deterministic one-dimensional four-class segmentation for the Madrid pilot.

No quotas, SB protections or regional production rules are applied here.
Divide-and-conquer DP minimizes within-class squared deviations of positive EUR
values (or of their logarithm in a explicitly separate sensitivity scenario).
"""
from __future__ import annotations

import numpy as np


def natural_breaks(values, classes=4, logarithmic=False):
    money = np.asarray(values, dtype=float)
    if not np.isfinite(money).all() or (money <= 0).any():
        raise ValueError("Pilot requires strictly positive, finite potentials")
    unique, counts = np.unique(money, return_counts=True)
    if len(unique) < classes:
        raise ValueError("Fewer distinct values than requested classes")
    x = np.log(unique) if logarithmic else unique
    # Scale for numerical conditioning; affine scaling preserves the optimum.
    x = (x - x.mean()) / max(x.std(), np.finfo(float).eps)
    n = len(x)
    weight = np.r_[0., np.cumsum(counts)]
    weighted_x = np.r_[0., np.cumsum(counts * x)]
    weighted_sq = np.r_[0., np.cumsum(counts * x * x)]
    previous = np.full(n + 1, np.inf)
    previous[0] = 0.
    path = np.zeros((classes + 1, n + 1), dtype=np.int64)
    for k in range(1, classes + 1):
        current = np.full(n + 1, np.inf)

        def solve(left, right, opt_left, opt_right):
            if left > right:
                return
            mid = (left + right) // 2
            indices = np.arange(max(k - 1, opt_left), min(mid - 1, opt_right) + 1)
            w = weight[mid] - weight[indices]
            sx = weighted_x[mid] - weighted_x[indices]
            cost = weighted_sq[mid] - weighted_sq[indices] - sx * sx / w
            scores = previous[indices] + np.maximum(cost, 0.)
            best = int(indices[int(np.argmin(scores))])
            current[mid] = float(scores[best - indices[0]])
            path[k, mid] = best
            solve(left, mid - 1, opt_left, best)
            solve(mid + 1, right, best, opt_right)

        solve(k, n, k - 1, n - 1)
        previous = current
    splits, end = [], n
    for k in range(classes, 1, -1):
        end = int(path[k, end])
        splits.append(end)
    splits = sorted(splits)
    # Any cut in the empirical gap gives the optimum. Use its midpoint,
    # rounded UP to a cent: inclusive lower bounds reproduce every membership.
    midpoint = [float(np.sqrt(unique[t-1] * unique[t])) if logarithmic
                else float((unique[t-1] + unique[t]) / 2) for t in splits]
    cuts = np.ceil(np.asarray(midpoint) * 100 - 1e-7) / 100
    total_variance = weighted_sq[-1] - weighted_x[-1] ** 2 / weight[-1]
    return {"cuts_ascending_eur": cuts.tolist(),
            "class_counts_D_C_B_A": np.bincount(np.searchsorted(cuts, money, side="right"), minlength=classes).tolist(),
            "goodness_variance_fit": float(1 - previous[n] / total_variance),
            "boundary_gaps_eur": [[float(unique[t-1]), float(unique[t])] for t in splits],
            "logarithmic": logarithmic}


def assign(values, cuts):
    positions = np.searchsorted(np.asarray(cuts), np.asarray(values, dtype=float), side="right")
    return np.asarray(list("DCBA"))[positions]
