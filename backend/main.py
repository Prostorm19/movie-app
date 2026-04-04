"""FastAPI backend for the Movie Recommendation System."""

import re
from contextlib import asynccontextmanager
from functools import lru_cache

import httpx
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware

from models import (
    GraphOut,
    MetricsOut,
    MovieOut,
    PersonalizedRequest,
    PosterOut,
    RecommendationOut,
)
from recommender import (
    build_models,
    get_recommendations,
    get_popular_movies,
    get_personalized_recommendations,
    get_similarity_graph,
    search_titles,
)
from evaluation import compute_metrics


# ---------------------------------------------------------------------------
# Startup / Shutdown
# ---------------------------------------------------------------------------

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Build ML models once at startup
    build_models()
    yield


app = FastAPI(
    title="Movie Recommendation API",
    version="2.0.0",
    lifespan=lifespan,
)

# Allow the React dev-server to call the API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Wikipedia Poster helper  (no API key required)
# ---------------------------------------------------------------------------

def _clean_title(movie_title: str) -> str:
    """Prepare a movie title for Wikipedia lookup.

    Two transformations:
    1. Strip the year suffix: "Toy Story (1995)" → "Toy Story"
    2. Move trailing article to front: "Princess Bride, The" → "The Princess Bride"
    """
    title = re.sub(r"\s*\(\d{4}\)\s*$", "", movie_title).strip()
    title = re.sub(r"^(.*),\s+(The|A|An)$", r"\2 \1", title)
    return title


@lru_cache(maxsize=2000)
def _fetch_poster_url(movie_title: str) -> str:
    """Query the Wikipedia REST API for a movie thumbnail. No API key needed.

    Wikipedia's /page/summary/{title} endpoint returns a JSON object that
    includes a `thumbnail.source` URL when the article has a lead image.
    Results are cached so repeated lookups are instant.
    """
    clean = _clean_title(movie_title)
    encoded = clean.replace(" ", "_")
    try:
        resp = httpx.get(
            f"https://en.wikipedia.org/api/rest_v1/page/summary/{encoded}",
            timeout=5.0,
            # Wikipedia asks that clients identify themselves via User-Agent
            headers={"User-Agent": "MovieRecommender/2.0 (student project; httpx)"},
            follow_redirects=True,
        )
        if resp.status_code == 200:
            source = resp.json().get("thumbnail", {}).get("source", "")
            if source:
                return source
    except Exception:
        pass
    return ""  # empty string → frontend shows gradient fallback


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/movies", response_model=list[MovieOut])
def list_movies(q: str = Query("", description="Optional title search")):
    """Return movies, optionally filtered by a search query."""
    return search_titles(q, limit=50) if q else search_titles("", limit=100)


@app.get("/recommend/{movie_title}", response_model=list[RecommendationOut])
def recommend(movie_title: str):
    """Return top-10 recommended movies for the given title."""
    results = get_recommendations(movie_title, top_n=10)
    if not results:
        raise HTTPException(
            status_code=404,
            detail=f"Movie '{movie_title}' not found. Try the /movies endpoint to search.",
        )
    return list(results)


@app.get("/popular", response_model=list[RecommendationOut])
def popular():
    """Return top-20 highest-rated movies (min 50 ratings)."""
    return get_popular_movies(top_n=20)


@app.get("/metrics", response_model=MetricsOut)
def metrics():
    """Return evaluation metrics (RMSE, MAE, Precision@10, Recall@10)."""
    return compute_metrics()


@app.post("/personalized-recommendations", response_model=list[RecommendationOut])
def personalized(body: PersonalizedRequest):
    """Return top-10 recommendations based on multiple liked movies."""
    if not body.liked_movies:
        raise HTTPException(status_code=400, detail="Provide at least one movie title.")
    results = get_personalized_recommendations(body.liked_movies, top_n=10)
    if not results:
        raise HTTPException(
            status_code=404,
            detail="None of the provided movies were found in the dataset.",
        )
    return results


@app.get("/graph/{movie_title}", response_model=GraphOut)
def graph(movie_title: str):
    """Return the similarity graph for a movie."""
    data = get_similarity_graph(movie_title, top_n=8)
    if data is None:
        raise HTTPException(
            status_code=404,
            detail=f"Movie '{movie_title}' not found.",
        )
    return data


@app.get("/poster/{movie_title}", response_model=PosterOut)
def poster(movie_title: str):
    """Return the TMDB poster URL for a movie title."""
    url = _fetch_poster_url(movie_title)
    return {"poster_url": url}
