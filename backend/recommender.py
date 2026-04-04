"""Hybrid Movie Recommendation Engine.

Combines:
1. Collaborative Filtering  – cosine similarity on a user-movie rating matrix
2. Matrix Factorization (SVD) – cosine similarity on latent feature embeddings
3. Content-Based Filtering – cosine similarity on genre vectors

The three scores are combined with configurable weights to produce a final
ranked list of recommendations.
"""

import numpy as np
import pandas as pd
from functools import lru_cache
from sklearn.decomposition import TruncatedSVD
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.preprocessing import MultiLabelBinarizer

from data_loader import load_movies, load_ratings

# ---------------------------------------------------------------------------
# Module-level caches – built once via build_models()
# ---------------------------------------------------------------------------
_collab_sim: pd.DataFrame | None = None       # movie×movie collaborative similarity
_svd_sim: pd.DataFrame | None = None          # movie×movie SVD latent-feature similarity
_content_sim: pd.DataFrame | None = None      # movie×movie content similarity
_movies_df: pd.DataFrame | None = None
_title_to_id: dict[str, int] = {}
_id_to_title: dict[int, str] = {}


# ---------------------------------------------------------------------------
# 1. Collaborative Filtering
# ---------------------------------------------------------------------------

def build_user_movie_matrix(ratings: pd.DataFrame) -> pd.DataFrame:
    """Pivot ratings into a user×movie matrix, fill NaN with 0."""
    return ratings.pivot_table(
        index="userId", columns="movieId", values="rating"
    ).fillna(0)


def compute_collab_similarity(user_movie_matrix: pd.DataFrame) -> pd.DataFrame:
    """Compute movie-movie cosine similarity from the user-movie matrix."""
    # Transpose so each column becomes a movie vector over users
    sim_matrix = cosine_similarity(user_movie_matrix.T)
    return pd.DataFrame(
        sim_matrix,
        index=user_movie_matrix.columns,
        columns=user_movie_matrix.columns,
    )


# ---------------------------------------------------------------------------
# 2. Matrix Factorization (SVD)
# ---------------------------------------------------------------------------

def compute_svd_similarity(
    user_movie_matrix: pd.DataFrame,
    n_components: int = 50,
) -> pd.DataFrame:
    """Use TruncatedSVD to extract latent features, then compute similarity.

    Steps:
        1. Transpose user-movie matrix so rows = movies, cols = users.
        2. Apply TruncatedSVD to reduce dimensionality to n_components.
        3. The transformed matrix gives each movie an embedding vector.
        4. Compute cosine similarity between all movie embeddings.
    """
    movie_user_matrix = user_movie_matrix.T  # rows = movies

    # Cap n_components at the smaller dimension minus 1 (SVD requirement)
    max_components = min(movie_user_matrix.shape) - 1
    n_components = min(n_components, max_components)

    svd = TruncatedSVD(n_components=n_components, random_state=42)
    movie_embeddings = svd.fit_transform(movie_user_matrix)

    sim_matrix = cosine_similarity(movie_embeddings)
    return pd.DataFrame(
        sim_matrix,
        index=user_movie_matrix.columns,
        columns=user_movie_matrix.columns,
    )


# ---------------------------------------------------------------------------
# 3. Content-Based Filtering
# ---------------------------------------------------------------------------

def build_genre_similarity(movies: pd.DataFrame) -> pd.DataFrame:
    """Build a movie-movie cosine similarity matrix from genre vectors."""
    genre_lists = movies["genres"].str.split("|")
    mlb = MultiLabelBinarizer()
    genre_matrix = mlb.fit_transform(genre_lists)
    sim_matrix = cosine_similarity(genre_matrix)
    return pd.DataFrame(
        sim_matrix,
        index=movies["movieId"].values,
        columns=movies["movieId"].values,
    )


# ---------------------------------------------------------------------------
# Initialisation helper
# ---------------------------------------------------------------------------

def build_models() -> None:
    """Pre-compute all similarity matrices (called once at server startup)."""
    global _collab_sim, _svd_sim, _content_sim, _movies_df
    global _title_to_id, _id_to_title

    _movies_df = load_movies()
    ratings_df = load_ratings()

    # Lookups
    _title_to_id = dict(zip(_movies_df["title"], _movies_df["movieId"]))
    _id_to_title = dict(zip(_movies_df["movieId"], _movies_df["title"]))

    # Collaborative similarity
    user_movie = build_user_movie_matrix(ratings_df)
    _collab_sim = compute_collab_similarity(user_movie)

    # SVD latent-feature similarity
    _svd_sim = compute_svd_similarity(user_movie)

    # Content similarity
    _content_sim = build_genre_similarity(_movies_df)


# ---------------------------------------------------------------------------
# Recommendation API
# ---------------------------------------------------------------------------

# Cache repeated calls – avoids recomputing similarity lookups for the same
# movie+parameters combination.  All arguments are hashable (str, int, float)
# so lru_cache works directly.
@lru_cache(maxsize=1000)
def get_recommendations(
    movie_title: str,
    top_n: int = 10,
    collab_weight: float = 0.5,
    svd_weight: float = 0.3,
    content_weight: float = 0.2,
) -> tuple[dict, ...]:
    """Return top-N recommended movies for the given title.

    Hybrid scoring formula:
        final_score = collab_weight * collaborative_similarity
                    + svd_weight    * svd_similarity
                    + content_weight * content_similarity

    Returns a tuple of dicts (tuple instead of list so lru_cache can work).
    """
    if _movies_df is None:
        build_models()

    movie_id = _title_to_id.get(movie_title)
    if movie_id is None:
        return ()

    # --- Collaborative scores ---
    collab_scores = pd.Series(dtype="float64")
    if _collab_sim is not None and movie_id in _collab_sim.columns:
        collab_scores = _collab_sim[movie_id]

    # --- SVD scores ---
    svd_scores = pd.Series(dtype="float64")
    if _svd_sim is not None and movie_id in _svd_sim.columns:
        svd_scores = _svd_sim[movie_id]

    # --- Content scores ---
    content_scores = pd.Series(dtype="float64")
    if _content_sim is not None and movie_id in _content_sim.columns:
        content_scores = _content_sim[movie_id]

    # Align indexes (some movies may only appear in one matrix)
    all_ids = set(collab_scores.index) | set(svd_scores.index) | set(content_scores.index)
    combined = pd.Series(0.0, index=list(all_ids))

    for mid in all_ids:
        c = collab_scores.get(mid, 0.0)
        s = svd_scores.get(mid, 0.0)
        g = content_scores.get(mid, 0.0)
        combined[mid] = collab_weight * c + svd_weight * s + content_weight * g

    # Remove the query movie itself
    combined = combined.drop(movie_id, errors="ignore")

    # Sort and take top-N
    top_ids = combined.sort_values(ascending=False).head(top_n)

    # --- Build explanations for each recommendation ---
    query_genres = set()
    if _movies_df is not None:
        row = _movies_df.loc[_movies_df["movieId"] == movie_id, "genres"]
        if not row.empty:
            query_genres = set(row.iloc[0].split("|"))

    genres_map = dict(zip(_movies_df["movieId"], _movies_df["genres"]))
    results = []
    for mid, score in top_ids.items():
        explanation = _build_explanation(
            mid, movie_title, query_genres, genres_map,
            collab_scores.get(mid, 0.0),
            svd_scores.get(mid, 0.0),
        )
        results.append(
            {
                "movieId": int(mid),
                "title": _id_to_title.get(mid, "Unknown"),
                "genres": genres_map.get(mid, ""),
                "score": round(float(score), 4),
                "explanation": explanation,
            }
        )
    return tuple(results)


def _build_explanation(
    rec_movie_id: int,
    query_title: str,
    query_genres: set[str],
    genres_map: dict,
    collab_score: float,
    svd_score: float,
) -> list[str]:
    """Generate human-readable explanation strings for a recommendation.

    This powers the Explainable AI feature – showing *why* a movie was
    recommended makes the system more transparent and trustworthy.
    """
    explanation: list[str] = []

    # 1. Shared genres
    rec_genres = set(genres_map.get(rec_movie_id, "").split("|"))
    shared = query_genres & rec_genres - {"(no genres listed)"}
    for genre in sorted(shared):
        explanation.append(f"Similar genre: {genre}")

    # 2. High collaborative score → users who liked the query also liked this
    if collab_score > 0.3:
        explanation.append(f"Highly rated by users who liked {query_title}")

    # 3. High SVD score → similar viewing pattern in latent features
    if svd_score > 0.3:
        explanation.append("Similar viewing pattern in latent features")

    # Fallback if nothing matched
    if not explanation:
        explanation.append("Recommended based on combined similarity analysis")

    return explanation


# ---------------------------------------------------------------------------
# Personalized Recommendations
# ---------------------------------------------------------------------------

def get_personalized_recommendations(
    liked_movies: list[str],
    top_n: int = 10,
) -> list[dict]:
    """Recommend movies based on multiple liked movies.

    Build a combined preference vector by averaging the hybrid similarity
    scores across all liked movies, then return the top-N results.
    """
    if _movies_df is None:
        build_models()

    liked_ids = [_title_to_id[t] for t in liked_movies if t in _title_to_id]
    if not liked_ids:
        return []

    # Collect similarity vectors for each liked movie, then average them
    all_ids: set = set()
    collab_vectors = []
    svd_vectors = []
    content_vectors = []

    for mid in liked_ids:
        if _collab_sim is not None and mid in _collab_sim.columns:
            collab_vectors.append(_collab_sim[mid])
            all_ids |= set(_collab_sim[mid].index)
        if _svd_sim is not None and mid in _svd_sim.columns:
            svd_vectors.append(_svd_sim[mid])
            all_ids |= set(_svd_sim[mid].index)
        if _content_sim is not None and mid in _content_sim.columns:
            content_vectors.append(_content_sim[mid])
            all_ids |= set(_content_sim[mid].index)

    if not all_ids:
        return []

    # Average each similarity type across liked movies
    avg_collab = pd.concat(collab_vectors, axis=1).mean(axis=1) if collab_vectors else pd.Series(dtype="float64")
    avg_svd = pd.concat(svd_vectors, axis=1).mean(axis=1) if svd_vectors else pd.Series(dtype="float64")
    avg_content = pd.concat(content_vectors, axis=1).mean(axis=1) if content_vectors else pd.Series(dtype="float64")

    # Hybrid formula: 0.5 collab + 0.3 SVD + 0.2 content
    combined = pd.Series(0.0, index=list(all_ids))
    for mid in all_ids:
        c = avg_collab.get(mid, 0.0)
        s = avg_svd.get(mid, 0.0)
        g = avg_content.get(mid, 0.0)
        combined[mid] = 0.5 * c + 0.3 * s + 0.2 * g

    # Exclude the liked movies themselves
    combined = combined.drop(liked_ids, errors="ignore")

    top_ids = combined.sort_values(ascending=False).head(top_n)

    # Build explanations based on the union of liked genres
    liked_genres: set[str] = set()
    for mid in liked_ids:
        row = _movies_df.loc[_movies_df["movieId"] == mid, "genres"]
        if not row.empty:
            liked_genres |= set(row.iloc[0].split("|"))

    genres_map = dict(zip(_movies_df["movieId"], _movies_df["genres"]))
    results = []
    for mid, score in top_ids.items():
        rec_genres = set(genres_map.get(mid, "").split("|"))
        shared = liked_genres & rec_genres - {"(no genres listed)"}
        explanation = [f"Similar genre: {g}" for g in sorted(shared)]
        explanation.append("Matches your combined movie preferences")
        results.append(
            {
                "movieId": int(mid),
                "title": _id_to_title.get(mid, "Unknown"),
                "genres": genres_map.get(mid, ""),
                "score": round(float(score), 4),
                "explanation": explanation,
            }
        )
    return results


# ---------------------------------------------------------------------------
# Similarity Graph
# ---------------------------------------------------------------------------

def get_similarity_graph(movie_title: str, top_n: int = 8) -> dict | None:
    """Return a graph structure of the most similar movies.

    Returns:
        {
            "nodes": [{"id": "Movie Title"}, ...],
            "links": [{"source": "A", "target": "B", "weight": 0.91}, ...]
        }
    """
    recs = get_recommendations(movie_title, top_n=top_n)
    if not recs:
        return None

    nodes = [{"id": movie_title}]
    links = []
    for rec in recs:
        nodes.append({"id": rec["title"]})
        links.append({
            "source": movie_title,
            "target": rec["title"],
            "weight": rec["score"],
        })

    return {"nodes": nodes, "links": links}


def get_popular_movies(top_n: int = 20) -> list[dict]:
    """Return the top-N movies by average rating (min 50 ratings)."""
    ratings = load_ratings()
    movies = load_movies()

    stats = ratings.groupby("movieId")["rating"].agg(["mean", "count"])
    # Only consider movies with enough ratings to be meaningful
    stats = stats[stats["count"] >= 50].sort_values("mean", ascending=False)
    top = stats.head(top_n)

    genres_map = dict(zip(movies["movieId"], movies["genres"]))
    title_map = dict(zip(movies["movieId"], movies["title"]))

    results = []
    for mid, row in top.iterrows():
        results.append(
            {
                "movieId": int(mid),
                "title": title_map.get(mid, "Unknown"),
                "genres": genres_map.get(mid, ""),
                "score": round(float(row["mean"]), 2),
            }
        )
    return results


def search_titles(query: str, limit: int = 10) -> list[dict]:
    """Return movies whose title contains the query (case-insensitive)."""
    movies = load_movies()
    mask = movies["title"].str.contains(query, case=False, na=False)
    hits = movies[mask].head(limit)
    return hits[["movieId", "title", "genres"]].to_dict(orient="records")
