import { useState, useEffect } from "react";

const sunIcon = (
    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none"
        viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
        <path strokeLinecap="round" strokeLinejoin="round"
            d="M12 3v1m0 16v1m8.66-13.66l-.71.71M4.05 19.95l-.71.71M21 12h-1M4 12H3m16.66 7.66l-.71-.71M4.05 4.05l-.71-.71M16 12a4 4 0 11-8 0 4 4 0 018 0z" />
    </svg>
);

const moonIcon = (
    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none"
        viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
        <path strokeLinecap="round" strokeLinejoin="round"
            d="M21 12.79A9 9 0 1111.21 3a7 7 0 009.79 9.79z" />
    </svg>
);

export default function Navbar({ page, setPage }) {
    const [dark, setDark] = useState(true);

    useEffect(() => {
        document.documentElement.classList.toggle("dark", dark);
    }, [dark]);

    const link =
        "px-3 py-1 rounded-lg text-sm font-medium transition-colors cursor-pointer ";
    const active = "bg-indigo-600 text-white";
    const inactive =
        "text-gray-600 dark:text-gray-300 hover:bg-gray-200 dark:hover:bg-gray-800";

    return (
        <nav className="sticky top-0 z-50 backdrop-blur bg-white/80 dark:bg-gray-900/80 border-b border-gray-200 dark:border-gray-800">
            <div className="max-w-6xl mx-auto flex items-center justify-between px-4 py-3">
                {/* Brand */}
                <h1
                    className="text-xl font-bold tracking-tight cursor-pointer bg-gradient-to-r from-indigo-500 to-purple-500 bg-clip-text text-transparent"
                    onClick={() => setPage("home")}
                >
                    🎬 MovieRec
                </h1>

                {/* Links */}
                <div className="flex items-center gap-2">
                    <button className={link + (page === "home" ? active : inactive)} onClick={() => setPage("home")}>
                        Home
                    </button>
                    <button className={link + (page === "popular" ? active : inactive)} onClick={() => setPage("popular")}>
                        Popular
                    </button>
                    <button className={link + (page === "personalized" ? active : inactive)} onClick={() => setPage("personalized")}>
                        Personalised
                    </button>
                    <button
                        onClick={() => setDark((d) => !d)}
                        className="ml-2 p-2 rounded-lg hover:bg-gray-200 dark:hover:bg-gray-800 transition-colors"
                        aria-label="Toggle dark mode"
                    >
                        {dark ? sunIcon : moonIcon}
                    </button>
                </div>
            </div>
        </nav>
    );
}
