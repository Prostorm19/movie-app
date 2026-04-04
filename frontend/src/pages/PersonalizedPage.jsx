import { useState } from "react";
import SearchBar from "../components/SearchBar";
import MovieCard from "../components/MovieCard";
import { fetchPersonalized } from "../api/api";

/**
 * PersonalizedPage — Users select multiple favourite movies and receive
 * personalised recommendations built from a combined preference vector.
 */
export default function PersonalizedPage() {
    const [liked, setLiked] = useState([]);       // list of selected titles
    const [recs, setRecs] = useState([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState(null);

    // Add a movie to the liked list (avoid duplicates)
    const handleSelect = (title) => {
        setLiked((prev) => (prev.includes(title) ? prev : [...prev, title]));
    };

    // Remove a movie from the liked list
    const handleRemove = (title) => {
        setLiked((prev) => prev.filter((t) => t !== title));
    };

    // Fetch personalised recommendations from the backend
    const handleGetRecs = () => {
        if (liked.length === 0) return;
        setLoading(true);
        setError(null);
        fetchPersonalized(liked)
            .then(setRecs)
            .catch((err) => {
                setError(err.response?.data?.detail || "Failed to fetch recommendations.");
                setRecs([]);
            })
            .finally(() => setLoading(false));
    };

    return (
        <div className="max-w-6xl mx-auto px-4 py-10">
            {/* Header */}
            <section className="text-center mb-10">
                <h2 className="text-3xl font-extrabold tracking-tight mb-2 bg-gradient-to-r from-indigo-500 to-purple-500 bg-clip-text text-transparent">
                    Personalised Recommendations
                </h2>
                <p className="text-gray-500 dark:text-gray-400 max-w-lg mx-auto">
                    Select movies you enjoy, and we'll build a combined preference
                    vector to recommend movies tailored just for you.
                </p>
            </section>

            {/* Movie search / selector */}
            <SearchBar onSelect={handleSelect} />

            {/* Selected movies */}
            {liked.length > 0 && (
                <div className="mt-6 flex flex-wrap gap-2 justify-center">
                    {liked.map((title) => (
                        <span
                            key={title}
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-indigo-100 dark:bg-indigo-900/40 text-indigo-700 dark:text-indigo-300 text-sm font-medium"
                        >
                            {title}
                            <button
                                onClick={() => handleRemove(title)}
                                className="ml-1 hover:text-red-500 transition-colors"
                                aria-label={`Remove ${title}`}
                            >
                                ×
                            </button>
                        </span>
                    ))}
                </div>
            )}

            {/* Get Recommendations button */}
            <div className="flex justify-center mt-6">
                <button
                    onClick={handleGetRecs}
                    disabled={liked.length === 0 || loading}
                    className="px-6 py-2.5 rounded-xl bg-indigo-600 text-white font-semibold hover:bg-indigo-700 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                >
                    {loading ? "Loading..." : "Get Recommendations"}
                </button>
            </div>

            {/* Error */}
            {error && (
                <p className="text-center text-red-500 dark:text-red-400 mt-6">{error}</p>
            )}

            {/* Results */}
            {recs.length > 0 && (
                <section className="mt-10">
                    <h3 className="text-2xl font-bold mb-6">
                        Your Personalised Picks
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
                </section>
            )}
        </div>
    );
}
