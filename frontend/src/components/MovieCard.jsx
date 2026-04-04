import { useEffect, useState } from "react";
import { fetchPoster } from "../api/api";

// Deterministic gradient — hashes the title to one of 8 colour combinations.
// The gradient always renders instantly; the real poster overlays it once loaded.
const GRADIENTS = [
    "from-indigo-500 to-purple-600",
    "from-blue-500 to-cyan-500",
    "from-violet-500 to-fuchsia-500",
    "from-rose-500 to-orange-500",
    "from-emerald-500 to-teal-500",
    "from-amber-500 to-orange-600",
    "from-sky-500 to-indigo-600",
    "from-pink-500 to-rose-600",
];

function titleGradient(title) {
    let hash = 0;
    for (let i = 0; i < title.length; i++) {
        hash = title.charCodeAt(i) + ((hash << 5) - hash);
    }
    return GRADIENTS[Math.abs(hash) % GRADIENTS.length];
}

/**
 * MovieCard
 *
 * Props:
 *   title       – movie title (string)
 *   genres      – pipe-separated genre string, e.g. "Action|Comedy"
 *   score       – similarity / rating score (number, optional)
 *   explanation – array of explanation strings for Explainable AI (optional)
 *   posterUrl   – optional real poster image URL; if provided it replaces the gradient
 */
export default function MovieCard({ title, genres, score, explanation = [] }) {
    const genreList = genres ? genres.split("|") : [];
    const gradient = titleGradient(title);
    const [poster, setPoster] = useState(null);

    // Fetch the real poster in the background.
    // The gradient renders immediately — the poster overlays it once loaded.
    // If the fetch fails or returns empty, the gradient remains (no broken images).
    useEffect(() => {
        let cancelled = false;
        fetchPoster(title)
            .then((data) => {
                if (!cancelled && data.poster_url) setPoster(data.poster_url);
            })
            .catch(() => { });
        return () => { cancelled = true; };
    }, [title]);

    return (
        <div className="group rounded-2xl border border-gray-200 dark:border-gray-800 bg-white dark:bg-gray-900 shadow-sm hover:shadow-xl hover:scale-[1.02] transition-all duration-200 overflow-hidden flex flex-col">
            {/* Poster area: gradient always visible; real poster loads on top */}
            <div className={`relative w-full aspect-[2/3] bg-gradient-to-br ${gradient} flex items-center justify-center overflow-hidden`}>
                {/* Film-reel placeholder icon, hidden once poster loads */}
                {!poster && (
                    <svg
                        className="w-14 h-14 text-white/30"
                        fill="none"
                        viewBox="0 0 24 24"
                        stroke="currentColor"
                        strokeWidth={1}
                    >
                        <path strokeLinecap="round" strokeLinejoin="round"
                            d="M7 4v16M17 4v16M3 8h4m10 0h4M3 16h4m10 0h4M4 20h16a1 1 0 001-1V5a1 1 0 00-1-1H4a1 1 0 00-1 1v14a1 1 0 001 1z" />
                    </svg>
                )}
                {/* Real poster — fades in; onError clears src so gradient shows again */}
                {poster && (
                    <img
                        src={poster}
                        alt={title}
                        className="absolute inset-0 w-full h-full object-cover transition-opacity duration-500"
                        onError={() => setPoster(null)}
                    />
                )}
            </div>

            {/* Info */}
            <div className="p-4 flex flex-col gap-2 flex-1">
                {/* Title */}
                <h3 className="text-base font-semibold leading-snug group-hover:text-indigo-500 transition-colors line-clamp-2">
                    {title}
                </h3>

                {/* Genre badges */}
                <div className="flex flex-wrap gap-1.5">
                    {genreList.map((g) => (
                        <span
                            key={g}
                            className="inline-block text-xs font-medium px-2 py-0.5 rounded-full bg-indigo-100 dark:bg-indigo-900/40 text-indigo-700 dark:text-indigo-300"
                        >
                            {g}
                        </span>
                    ))}
                </div>

                {/* Score */}
                {score !== undefined && (
                    <p className="text-sm text-gray-500 dark:text-gray-400">
                        Score:{" "}
                        <span className="font-bold text-indigo-600 dark:text-indigo-400">
                            {score}
                        </span>
                    </p>
                )}

                {/* Explanation — Explainable AI */}
                {explanation.length > 0 && (
                    <ul className="mt-1 space-y-0.5">
                        {explanation.map((text, i) => (
                            <li
                                key={i}
                                className="text-xs italic text-gray-400 dark:text-gray-500 flex items-start gap-1"
                            >
                                <span className="text-indigo-400">•</span>
                                {text}
                            </li>
                        ))}
                    </ul>
                )}
            </div>
        </div>
    );
}
