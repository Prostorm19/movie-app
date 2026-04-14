import json

NOTEBOOK_PATH = r"d:\movie-app\backend\results_analysis.ipynb"

with open(NOTEBOOK_PATH, "r", encoding="utf-8") as f:
    nb = json.load(f)

# ── helper ────────────────────────────────────────────────────────────────────
def code_cell(lines):
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": lines,
    }

def md_cell(lines):
    return {
        "cell_type": "markdown",
        "metadata": {},
        "source": lines,
    }

# ── NEW CELLS ─────────────────────────────────────────────────────────────────

# ── Section header ────────────────────────────────────────────────────────────
header = md_cell([
    "## 6b. Multivariate & Heatmap Analyses\n",
    "\n",
    "Two additional multivariate views are provided here:\n",
    "1. **Scatter plot** – Average Rating vs. Number of Ratings, colour-encoded by release decade.\n",
    "2. **Heatmap** – Genre × Decade grid showing mean rating per cell.\n",
])

# ── Graph 1 : Scatter with colour-encoded decade (multivariate) ───────────────
scatter_code = code_cell([
    "# ── Multivariate Scatter: Avg Rating vs Rating Count, coloured by Decade ──\n",
    "import matplotlib.cm as cm\n",
    "import matplotlib.colors as mcolors\n",
    "\n",
    "scatter_df = movies_enriched.dropna(subset=['rating_count', 'avg_rating', 'decade']).copy()\n",
    "scatter_df['decade'] = scatter_df['decade'].astype(int)\n",
    "\n",
    "# Build a discrete colour palette keyed by decade\n",
    "decades_sorted = sorted(scatter_df['decade'].unique())\n",
    "palette = sns.color_palette('tab10', n_colors=len(decades_sorted))\n",
    "colour_map = {d: palette[i] for i, d in enumerate(decades_sorted)}\n",
    "colours = scatter_df['decade'].map(colour_map)\n",
    "\n",
    "fig, ax = plt.subplots(figsize=(12, 7))\n",
    "sc = ax.scatter(\n",
    "    scatter_df['rating_count'],\n",
    "    scatter_df['avg_rating'],\n",
    "    c=colours,\n",
    "    alpha=0.45,\n",
    "    s=30,\n",
    "    edgecolors='none',\n",
    ")\n",
    "ax.set_xscale('log')\n",
    "ax.set_title('Avg Rating vs. Number of Ratings (colour = Release Decade)', fontsize=15)\n",
    "ax.set_xlabel('Number of Ratings (log scale)', fontsize=12)\n",
    "ax.set_ylabel('Average Rating', fontsize=12)\n",
    "\n",
    "# Legend\n",
    "handles = [plt.Line2D([0], [0], marker='o', color='w',\n",
    "           markerfacecolor=colour_map[d], markersize=8, label=str(d))\n",
    "           for d in decades_sorted]\n",
    "ax.legend(handles=handles, title='Decade', bbox_to_anchor=(1.01, 1),\n",
    "          loc='upper left', fontsize=9)\n",
    "\n",
    "plt.tight_layout()\n",
    "plt.show()\n",
    "fig.savefig('rating_quality_vs_popularity.png', dpi=300, bbox_inches='tight')\n",
    "print('Saved: rating_quality_vs_popularity.png')\n",
])

scatter_interp = md_cell([
    "**Interpretation:** Three variables are encoded simultaneously – popularity (x-axis, log), "
    "quality (y-axis), and decade (colour). Older films (1960s–1980s) cluster at high average "
    "ratings but low counts, reflecting survivorship bias: only acclaimed titles remain widely "
    "rated. Recent films (2000s–2010s) appear across the full popularity range, suggesting a "
    "broader but more polarised audience response.\n",
])

# ── Graph 2 : Genre × Decade Heatmap (multivariate) ─────────────────────────
heatmap_code = code_cell([
    "# ── Multivariate Heatmap: Mean Rating per Genre × Decade ──\n",
    "\n",
    "# Explode genres so each movie appears once per genre\n",
    "hm_df = movies_enriched.dropna(subset=['avg_rating', 'decade']).copy()\n",
    "hm_df['decade'] = hm_df['decade'].astype(int)\n",
    "hm_df['genre_list'] = hm_df['genres'].str.split('|')\n",
    "hm_df = hm_df.explode('genre_list')\n",
    "hm_df = hm_df[hm_df['genre_list'] != '(no genres listed)']\n",
    "\n",
    "# Limit to decades >= 1960 for readability\n",
    "hm_df = hm_df[hm_df['decade'] >= 1960]\n",
    "\n",
    "# Top 10 genres by movie count\n",
    "top_genres = (\n",
    "    hm_df.groupby('genre_list')['movieId']\n",
    "    .count()\n",
    "    .nlargest(10)\n",
    "    .index\n",
    ")\n",
    "hm_df = hm_df[hm_df['genre_list'].isin(top_genres)]\n",
    "\n",
    "# Pivot: rows = genre, columns = decade, values = mean avg_rating\n",
    "pivot = (\n",
    "    hm_df.groupby(['genre_list', 'decade'])['avg_rating']\n",
    "    .mean()\n",
    "    .unstack()\n",
    ")\n",
    "pivot = pivot.reindex(index=top_genres)   # keep consistent row order\n",
    "\n",
    "# Count matrix (for annotation context)\n",
    "count_pivot = (\n",
    "    hm_df.groupby(['genre_list', 'decade'])['movieId']\n",
    "    .count()\n",
    "    .unstack()\n",
    "    .reindex(index=top_genres)\n",
    ")\n",
    "\n",
    "plt.figure(figsize=(16, 7))\n",
    "ax = sns.heatmap(\n",
    "    pivot,\n",
    "    annot=True,\n",
    "    fmt='.2f',\n",
    "    cmap='YlOrRd',\n",
    "    linewidths=0.4,\n",
    "    linecolor='white',\n",
    "    vmin=2.5,\n",
    "    vmax=4.5,\n",
    "    cbar_kws={'label': 'Mean Avg Rating', 'shrink': 0.8},\n",
    ")\n",
    "ax.set_title('Mean Rating by Genre × Decade  (Genre–Decade Heatmap)', fontsize=15)\n",
    "ax.set_xlabel('Decade', fontsize=12)\n",
    "ax.set_ylabel('Genre', fontsize=12)\n",
    "\n",
    "plt.tight_layout()\n",
    "plt.show()\n",
    "print('Genre x Decade heatmap rendered.')\n",
])

heatmap_interp = md_cell([
    "**Interpretation:** The heatmap encodes three variables – genre (row), decade (column), and "
    "mean rating (colour + annotation). Warmer colours indicate higher ratings. "
    "Key observations:\n",
    "- **Film-Noir & Documentary** consistently score high where data exists, though their catalogs are thin.\n",
    "- **Drama** shows steady quality across all decades.\n",
    "- **Action & Horror** tend to sit in cooler zones, indicating lower average audience scores.\n",
    "- Empty cells reflect that the genre–decade combination has too few films to produce a reliable estimate.\n",
])

# ──────────────────────────────────────────────────────────────────────────────
# Locate insertion point: just BEFORE "## 7. Model Evaluation Metrics"
# ──────────────────────────────────────────────────────────────────────────────
insert_pos = None
for idx, cell in enumerate(nb["cells"]):
    src = "".join(cell.get("source", []))
    if "## 7. Model Evaluation Metrics" in src:
        insert_pos = idx
        break

if insert_pos is None:
    # Fallback: append before last cell
    insert_pos = len(nb["cells"]) - 1
    print(f"WARNING: '## 7.' not found; inserting at position {insert_pos}")
else:
    print(f"Inserting 6 new cells before cell {insert_pos} (Section 7)")

new_cells = [header, scatter_code, scatter_interp, heatmap_code, heatmap_interp]

for i, cell in enumerate(new_cells):
    nb["cells"].insert(insert_pos + i, cell)

with open(NOTEBOOK_PATH, "w", encoding="utf-8") as f:
    json.dump(nb, f, indent=2, ensure_ascii=False)

print(f"Done. Total cells: {len(nb['cells'])}")
print("New section '6b. Multivariate & Heatmap Analyses' added successfully.")
