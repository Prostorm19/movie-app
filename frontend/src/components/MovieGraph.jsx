import { useEffect, useState, useCallback, useRef } from "react";
import ForceGraph2D from "react-force-graph-2d";
import { forceCollide } from "d3-force";
import { fetchGraph } from "../api/api";

/**
 * MovieGraph — Interactive movie similarity network graph.
 *
 * Features:
 * - Nodes represent movies (center node = queried movie).
 * - Edges connect similar movies; thickness ∝ similarity weight.
 * - Hover tooltip shows similarity value.
 * - Click a node to navigate to that movie's recommendations.
 */
export default function MovieGraph({ movieTitle, onNodeClick }) {
    const [graphData, setGraphData] = useState(null);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState(null);
    const [tooltip, setTooltip] = useState(null);
    const containerRef = useRef(null);
    const [dimensions, setDimensions] = useState({ width: 600, height: 400 });
    const fgRef = useRef();

    // Fetch graph data when the movie changes
    useEffect(() => {
        if (!movieTitle) return;
        setLoading(true);
        setError(null);
        fetchGraph(movieTitle)
            .then((data) => {
                // Defensive guard against malformed payloads
                if (!data || !Array.isArray(data.nodes) || !Array.isArray(data.links)) {
                    setError("Graph data is unavailable for this movie.");
                    setGraphData(null);
                    return;
                }
                setGraphData(data);
            })
            .catch(() => {
                setError("Could not load similarity graph for this movie.");
                setGraphData(null);
            })
            .finally(() => setLoading(false));
    }, [movieTitle]);

    // Responsive width — measure the container
    useEffect(() => {
        if (!containerRef.current) return;
        const observer = new ResizeObserver((entries) => {
            for (const entry of entries) {
                setDimensions({
                    width: entry.contentRect.width,
                    height: 400,
                });
            }
        });
        observer.observe(containerRef.current);
        return () => observer.disconnect();
    }, []);

    // Configure force graph physics
    useEffect(() => {
        if (fgRef.current) {
            fgRef.current.d3Force("charge").strength(-1200); // Increase repulsion
            fgRef.current.d3Force("link").distance(180); // Increase link distance
            fgRef.current.d3Force("collide", forceCollide(30)); // Collision detection to prevent overlapping nodes
        }
    }, [graphData]);

    // Node styling — highlight the center movie
    const paintNode = useCallback(
        (node, ctx) => {
            const isCenter = node.id === movieTitle;
            const radius = isCenter ? 8 : 6;

            ctx.beginPath();
            ctx.arc(node.x, node.y, radius, 0, 2 * Math.PI);
            ctx.fillStyle = isCenter ? "#6366f1" : "#818cf8";
            ctx.fill();
            ctx.strokeStyle = "#312e81";
            ctx.lineWidth = 1.5;
            ctx.stroke();

            // Label
            const fontSize = isCenter ? 12 : 10;
            ctx.font = `${isCenter ? "bold " : ""}${fontSize}px sans-serif`;
            const label = node.id;
            const textWidth = ctx.measureText(label).width;
            const bckgDimensions = [textWidth, fontSize].map(n => n + fontSize * 0.2);

            // Draw text background
            ctx.fillStyle = "rgba(10, 10, 10, 0.7)";
            ctx.fillRect(
                node.x - bckgDimensions[0] / 2,
                node.y + radius + 2,
                bckgDimensions[0],
                bckgDimensions[1]
            );

            // Draw text
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";
            ctx.fillStyle = isCenter ? "#ffffff" : "#a5b4fc";
            ctx.fillText(label, node.x, node.y + radius + 2 + bckgDimensions[1] / 2);
        },
        [movieTitle]
    );

    // Link width proportional to similarity weight
    const linkWidth = useCallback((link) => {
        const w = typeof link.weight === "number" ? link.weight : 0;
        // Use a gentle non-linear scale so small differences are still visible
        const scaled = Math.pow(w, 0.7) * 6;
        return Math.max(0.5, scaled);
    }, []);

    const linkColor = useCallback(() => "#6366f180", []);

    // Tooltip on link hover
    const handleLinkHover = useCallback((link) => {
        if (link) {
            const weight =
                typeof link.weight === "number" ? link.weight.toFixed(4) : "n/a";
            setTooltip(
                `${link.source.id || link.source} ↔ ${
                    link.target.id || link.target
                }: ${weight}`
            );
        } else {
            setTooltip(null);
        }
    }, []);

    // Click node → load that movie's recommendations
    const handleNodeClick = useCallback(
        (node) => {
            if (onNodeClick) onNodeClick(node.id);
        },
        [onNodeClick]
    );

    if (loading) {
        return (
            <div className="flex justify-center py-8">
                <div className="h-8 w-8 rounded-full border-4 border-indigo-500 border-t-transparent animate-spin" />
            </div>
        );
    }

    if (error) {
        return (
            <div className="mt-4 rounded-2xl border border-red-500/40 bg-red-950/40 px-4 py-3 text-sm text-red-200">
                {error}
            </div>
        );
    }

    if (!graphData) return null;

    const hasUsableGraph =
        Array.isArray(graphData.nodes) &&
        Array.isArray(graphData.links) &&
        graphData.nodes.length > 1 &&
        graphData.links.length > 0;

    return (
        <div
            ref={containerRef}
            className="relative w-full rounded-2xl border border-gray-200 dark:border-gray-800 bg-gray-950 overflow-hidden"
        >
            <h4 className="text-lg font-bold px-4 pt-4 text-gray-200">
                Similarity Network
            </h4>
            <p className="text-xs text-gray-400 px-4 pb-2">
                Click a node to explore that movie&apos;s recommendations
            </p>

            {/* Tooltip */}
            {tooltip && (
                <div className="absolute top-3 right-3 bg-gray-800/90 text-gray-200 text-xs px-3 py-1.5 rounded-lg shadow z-10">
                    {tooltip}
                </div>
            )}

            {hasUsableGraph ? (
                <ForceGraph2D
                    ref={fgRef}
                    graphData={graphData}
                    width={dimensions.width}
                    height={dimensions.height}
                    nodeCanvasObject={paintNode}
                    nodePointerAreaPaint={(node, color, ctx) => {
                        ctx.beginPath();
                        ctx.arc(node.x, node.y, 10, 0, 2 * Math.PI);
                        ctx.fillStyle = color;
                        ctx.fill();
                    }}
                    linkWidth={linkWidth}
                    linkColor={linkColor}
                    onLinkHover={handleLinkHover}
                    onNodeClick={handleNodeClick}
                    cooldownTicks={90}
                    backgroundColor="transparent"
                />
            ) : (
                <div className="flex h-[260px] items-center justify-center text-xs text-gray-500">
                    Not enough similarity data to draw a meaningful graph for this
                    movie.
                </div>
            )}
        </div>
    );
}
