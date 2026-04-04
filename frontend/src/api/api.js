import axios from "axios";

const API = axios.create({
    baseURL: "http://localhost:8000",
});

/** Search / list movies */
export const fetchMovies = (query = "") =>
    API.get("/movies", { params: { q: query } }).then((res) => res.data);

/** Get recommendations for a movie title */
export const fetchRecommendations = (title) =>
    API.get(`/recommend/${encodeURIComponent(title)}`).then((res) => res.data);

/** Get popular movies */
export const fetchPopular = () =>
    API.get("/popular").then((res) => res.data);

/** Get evaluation metrics (RMSE, MAE, Precision@10, Recall@10) */
export const fetchMetrics = () =>
    API.get("/metrics").then((res) => res.data);

/** Get personalized recommendations based on multiple liked movies */
export const fetchPersonalized = (likedMovies) =>
    API.post("/personalized-recommendations", { liked_movies: likedMovies }).then(
        (res) => res.data
    );

/** Get the similarity graph for a movie */
export const fetchGraph = (title) =>
    API.get(`/graph/${encodeURIComponent(title)}`).then((res) => res.data);

/** Get the poster URL for a movie (backend queries Wikipedia, no API key needed) */
export const fetchPoster = (title) =>
    API.get(`/poster/${encodeURIComponent(title)}`).then((res) => res.data);
