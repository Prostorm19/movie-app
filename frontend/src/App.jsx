import { useState } from "react";
import Navbar from "./components/Navbar";
import Home from "./pages/Home";
import Popular from "./pages/Popular";
import PersonalizedPage from "./pages/PersonalizedPage";

export default function App() {
    const [page, setPage] = useState("home");

    return (
        <>
            <Navbar page={page} setPage={setPage} />
            {page === "home" && <Home />}
            {page === "popular" && <Popular />}
            {page === "personalized" && <PersonalizedPage />}
        </>
    );
}
