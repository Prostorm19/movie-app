import { useState } from "react";
import SearchBar from "../components/SearchBar";
import Recommendations from "./Recommendations";

export default function Home() {
    const [selectedMovie, setSelectedMovie] = useState(null);

    return (
        <div className="max-w-6xl mx-auto px-4 py-10">
            {/* Hero Section */}
            <section className="text-center mb-12">
                <h2 className="text-4xl font-extrabold tracking-tight mb-3 bg-gradient-to-r from-indigo-500 to-purple-500 bg-clip-text text-transparent">
                    Find Your Next Favourite Movie
                </h2>
                <p className="text-gray-500 dark:text-gray-400 max-w-md mx-auto">
                    Search for a movie you love and we'll recommend similar ones using
                    collaborative &amp; content-based filtering.
                </p>
            </section>

            {/* Search */}
            <SearchBar onSelect={(title) => setSelectedMovie(title)} />

            {/* Recommendations */}
            {selectedMovie && (
                <section className="mt-12">
                    <Recommendations
                        movieTitle={selectedMovie}
                        onMovieChange={(title) => setSelectedMovie(title)}
                    />
                </section>
            )}
        </div>
    );
}
