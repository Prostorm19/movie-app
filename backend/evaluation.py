"""Evaluation module for the recommendation system.

Implements standard ML evaluation metrics using a train/test split:
    - RMSE  (Root Mean Squared Error)  – measures average prediction error
    - MAE   (Mean Absolute Error)      – measures average absolute error
    - Precision@K – fraction of top-K recommendations that are relevant
    - Recall@K    – fraction of relevant items that appear in top-K

A "relevant" item is a movie the user rated >= 4.0 in the test set.
"""

import numpy as np
import pandas as pd
from sklearn.decomposition import TruncatedSVD
from sklearn.model_selection import train_test_split

from data_loader import load_ratings

# Cache the metrics after the first computation to avoid
# repeating the expensive train/test evaluation on every request.
_cached_metrics: dict | None = None


def compute_metrics(k: int = 10) -> dict:
    """Compute evaluation metrics and cache the result.

    Returns:
        dict with keys: rmse, mae, precision_at_10, recall_at_10
    """
    global _cached_metrics
    if _cached_metrics is not None:
        return _cached_metrics

    ratings = load_ratings()

    # ----- Step 1: Train/test split (80/20) -----
    # We split randomly by rows.  Each row is one (user, movie, rating).
    train, test = train_test_split(ratings, test_size=0.2, random_state=42)

    # ----- Step 2: Build user-movie matrix from TRAINING data -----
    train_matrix = train.pivot_table(
        index="userId", columns="movieId", values="rating"
    ).fillna(0)

    # ----- Step 3: Use SVD to create a low-rank approximation -----
    # This approximation acts as our predicted ratings.
    n_components = min(50, min(train_matrix.shape) - 1)
    svd = TruncatedSVD(n_components=n_components, random_state=42)
    user_features = svd.fit_transform(train_matrix)          # users × latent
    movie_features = svd.components_                          # latent × movies

    # Reconstruct the full predicted matrix (users × movies)
    predicted_matrix = pd.DataFrame(
        np.dot(user_features, movie_features),
        index=train_matrix.index,
        columns=train_matrix.columns,
    )

    # ----- Step 4: Evaluate on the TEST set -----
    # Collect actual vs predicted for every (user, movie) pair in the test set
    # that exists in our training matrix dimensions.
    actuals = []
    predictions = []
    for _, row in test.iterrows():
        uid, mid, rating = int(row["userId"]), int(row["movieId"]), row["rating"]
        if uid in predicted_matrix.index and mid in predicted_matrix.columns:
            actuals.append(rating)
            predictions.append(predicted_matrix.loc[uid, mid])

    actuals = np.array(actuals)
    predictions = np.array(predictions)

    # RMSE: sqrt(mean((predicted - actual)^2))
    rmse = float(np.sqrt(np.mean((predictions - actuals) ** 2)))

    # MAE: mean(|predicted - actual|)
    mae = float(np.mean(np.abs(predictions - actuals)))

    # ----- Step 5: Precision@K and Recall@K -----
    # For each user in the test set:
    #   - "Relevant" = movies the user rated >= 4.0 in the test set
    #   - "Recommended" = top-K movies by predicted score (excluding train movies)
    #   - Precision@K = |relevant ∩ recommended| / K
    #   - Recall@K    = |relevant ∩ recommended| / |relevant|
    test_users = test.groupby("userId")
    precisions = []
    recalls = []

    for uid, group in test_users:
        if uid not in predicted_matrix.index:
            continue

        # Movies rated >= 4.0 in test → these are the "relevant" items
        relevant = set(group.loc[group["rating"] >= 4.0, "movieId"].values)
        if not relevant:
            continue

        # Predicted scores for this user (only movies in our matrix)
        user_preds = predicted_matrix.loc[uid]

        # Exclude movies the user already saw in training
        train_seen = set(train.loc[train["userId"] == uid, "movieId"].values)
        candidates = user_preds.drop(
            labels=[m for m in train_seen if m in user_preds.index], errors="ignore"
        )

        # Top-K by predicted score
        top_k = set(candidates.sort_values(ascending=False).head(k).index)

        # Precision@K and Recall@K
        hits = relevant & top_k
        precisions.append(len(hits) / k)
        recalls.append(len(hits) / len(relevant))

    precision_at_k = float(np.mean(precisions)) if precisions else 0.0
    recall_at_k = float(np.mean(recalls)) if recalls else 0.0

    _cached_metrics = {
        "rmse": round(rmse, 4),
        "mae": round(mae, 4),
        "precision_at_10": round(precision_at_k, 4),
        "recall_at_10": round(recall_at_k, 4),
    }
    return _cached_metrics
