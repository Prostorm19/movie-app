import { useEffect, useState } from "react";
import { fetchRecommendations } from "../api/api";
import MovieCard from "../components/MovieCard";
import MovieGraph from "../components/MovieGraph";

export default function Recommendations({ movieTitle, onMovieChange }) {
    const [recs, setRecs] = useState([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState(null);

    useEffect(() => {
        if (!movieTitle) return;
        setLoading(true);
        setError(null);
        fetchRecommendations(movieTitle)
            .then((data) => setRecs(data))
            .catch((err) => {
                const msg =
                    err.response?.data?.detail || "Failed to fetch recommendations.";
                setError(msg);
                setRecs([]);
            })
            .finally(() => setLoading(false));
    }, [movieTitle]);

    // When a node in the graph is clicked, navigate to that movie
    const handleGraphNodeClick = (title) => {
        if (onMovieChange) onMovieChange(title);
    };

    if (loading)
        return (
            <div className="flex justify-center py-16">
                <div className="h-10 w-10 rounded-full border-4 border-indigo-500 border-t-transparent animate-spin" />
            </div>
        );

    if (error)
        return (
            <p className="text-center text-red-500 dark:text-red-400 py-8">{error}</p>
        );

    if (recs.length === 0) return null;

    return (
        <>
            <h3 className="text-2xl font-bold mb-6">
                Because you liked{" "}
                <span className="text-indigo-500">{movieTitle}</span>
            </h3>
            <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
                {recs.map((r) => (
                    <MovieCard
                        key={r.movieId}
                        title={r.title}
                        genres={r.genres}
                        score={r.score}
                        explanation={r.explanation || []}
                    />
                ))}
            </div>

            {/* Similarity Network Graph */}
            <section className="mt-12">
                <MovieGraph
                    movieTitle={movieTitle}
                    onNodeClick={handleGraphNodeClick}
                />
            </section>
        </>
    );
}
