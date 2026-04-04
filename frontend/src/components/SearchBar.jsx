import { useState, useEffect, useRef } from "react";
import { fetchMovies } from "../api/api";

export default function SearchBar({ onSelect }) {
    const [query, setQuery] = useState("");
    const [results, setResults] = useState([]);
    const [open, setOpen] = useState(false);
    const wrapperRef = useRef(null);

    // Debounced autocomplete search
    useEffect(() => {
        if (query.length < 2) {
            setResults([]);
            setOpen(false);
            return;
        }
        const timer = setTimeout(() => {
            fetchMovies(query)
                .then((data) => {
                    setResults(data);
                    setOpen(true);
                })
                .catch(() => setResults([]));
        }, 300);
        return () => clearTimeout(timer);
    }, [query]);

    // Close dropdown on outside click
    useEffect(() => {
        const handler = (e) => {
            if (wrapperRef.current && !wrapperRef.current.contains(e.target)) {
                setOpen(false);
            }
        };
        document.addEventListener("mousedown", handler);
        return () => document.removeEventListener("mousedown", handler);
    }, []);

    const handleSelect = (movie) => {
        setQuery(movie.title);
        setOpen(false);
        onSelect(movie.title);
    };

    return (
        <div ref={wrapperRef} className="relative w-full max-w-xl mx-auto">
            <input
                type="text"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Search for a movie…"
                className="w-full px-5 py-3 rounded-xl border border-gray-300 dark:border-gray-700 bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 shadow-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 transition"
            />

            {open && results.length > 0 && (
                <ul className="absolute z-40 mt-1 w-full bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-xl shadow-lg max-h-72 overflow-y-auto">
                    {results.map((m) => (
                        <li
                            key={m.movieId}
                            onClick={() => handleSelect(m)}
                            className="px-4 py-2 cursor-pointer hover:bg-indigo-50 dark:hover:bg-gray-700 transition-colors"
                        >
                            <span className="font-medium">{m.title}</span>
                            <span className="ml-2 text-xs text-gray-400">{m.genres}</span>
                        </li>
                    ))}
                </ul>
            )}
        </div>
    );
}
