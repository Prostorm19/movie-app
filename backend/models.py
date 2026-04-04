"""Pydantic models for API requests and responses."""

from pydantic import BaseModel


class MovieOut(BaseModel):
    """Schema returned when listing movies."""
    movieId: int
    title: str
    genres: str


class RecommendationOut(BaseModel):
    """Schema returned for a single recommendation, now with explanations."""
    movieId: int
    title: str
    genres: str
    score: float
    explanation: list[str] = []


# ---- Personalized Recommendations ----

class PersonalizedRequest(BaseModel):
    """Request body for personalized recommendations."""
    liked_movies: list[str]


# ---- Similarity Graph ----

class GraphNode(BaseModel):
    id: str

class GraphLink(BaseModel):
    source: str
    target: str
    weight: float

class GraphOut(BaseModel):
    nodes: list[GraphNode]
    links: list[GraphLink]


# ---- Evaluation Metrics ----

class MetricsOut(BaseModel):
    rmse: float
    mae: float
    precision_at_10: float
    recall_at_10: float


# ---- TMDB Poster ----

class PosterOut(BaseModel):
    poster_url: str
