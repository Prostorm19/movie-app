"""Utilities to load and cache the CSV datasets."""

import os
import pandas as pd

# Resolve paths relative to this file so the server can be started from anywhere.
_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_MOVIES_PATH = os.path.join(_BASE, "movies.csv")
_RATINGS_PATH = os.path.join(_BASE, "ratings.csv")

# Module-level caches (loaded once on first call)
_movies_df: pd.DataFrame | None = None
_ratings_df: pd.DataFrame | None = None


def load_movies() -> pd.DataFrame:
    """Return the movies DataFrame (cached after first load)."""
    global _movies_df
    if _movies_df is None:
        _movies_df = pd.read_csv(_MOVIES_PATH)
    return _movies_df


def load_ratings() -> pd.DataFrame:
    """Return the ratings DataFrame (cached after first load)."""
    global _ratings_df
    if _ratings_df is None:
        _ratings_df = pd.read_csv(_RATINGS_PATH)
    return _ratings_df
