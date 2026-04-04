import { useEffect, useState } from "react";
import { fetchPopular } from "../api/api";
import MovieCard from "../components/MovieCard";

export default function Popular() {
    const [movies, setMovies] = useState([]);
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        fetchPopular()
            .then(setMovies)
            .catch(() => setMovies([]))
            .finally(() => setLoading(false));
    }, []);

    if (loading)
        return (
            <div className="flex justify-center py-16">
                <div className="h-10 w-10 rounded-full border-4 border-indigo-500 border-t-transparent animate-spin" />
            </div>
        );

    return (
        <div className="max-w-6xl mx-auto px-4 py-10">
            <h2 className="text-3xl font-extrabold mb-2">🔥 Popular Movies</h2>
            <p className="text-gray-500 dark:text-gray-400 mb-8">
                Top-rated movies with at least 50 ratings.
            </p>
            <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
                {movies.map((m) => (
                    <MovieCard
                        key={m.movieId}
                        title={m.title}
                        genres={m.genres}
                        score={m.score}
                    />
                ))}
            </div>
        </div>
    );
}
