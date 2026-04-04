# Movie Recommendation System 🎬

A full-stack movie recommendation application featuring content-based filtering, personalized recommendations, and an interactive similarity graph visualization.

![Movie App Preview](https://img.shields.io/badge/Status-Active-brightgreen)
![React](https://img.shields.io/badge/Frontend-React-61DAFB?logo=react)
![FastAPI](https://img.shields.io/badge/Backend-FastAPI-009688?logo=fastapi)
![Tailwind CSS](https://img.shields.io/badge/Styling-Tailwind_CSS-38B2AC?logo=tailwind-css)

## ✨ Features

- **Search & Explore**: Quickly find movies from a dataset of thousands.
- **Smart Recommendations**: Get top-10 similar movies based on content similarity.
- **Personalized Recommendations**: Input multiple liked movies to get a tailored list of suggestions.
- **Interactive Similarity Graph**: Visualize how movies are connected through a dynamic D3-powered force graph.
- **Auto-Fetching Posters**: Automatically retrieves movie poster thumbnails via the Wikipedia API.
- **Metrics Dashboard**: View performance metrics like RMSE, MAE, and Precision@10.

## 🚀 Tech Stack

- **Frontend**:
  - React (Vite)
  - Tailwind CSS for premium responsive styling
  - `react-force-graph-2d` for interactive network visualizations
  - Axios for API communication
- **Backend**:
  - FastAPI (Python)
  - Pandas for data manipulation
  - Scikit-learn for similarity computations (Cosine Similarity)
  - Wikipedia API integration for poster fetching

## 🛠️ Installation & Setup

### Prerequisites
- [Python 3.10+](https://www.python.org/downloads/)
- [Node.js & npm](https://nodejs.org/)

### 1. Clone the Repository
```bash
git clone https://github.com/your-username/movie-recommender.git
cd movie-recommender
```

### 2. Backend Setup
```bash
cd backend

# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Start the FastAPI server
uvicorn main:app --reload
```
The backend will be running at `http://localhost:8000`.

### 3. Frontend Setup
Open a new terminal window:
```bash
cd frontend

# Install dependencies
npm install

# Start the development server
npm run dev
```
The frontend will be running at `http://localhost:5173`.

## 📊 Dataset
This project uses a movie dataset (included as `movies.csv` and `ratings.csv` in the root directory) which contains movie metadata and user ratings.

## 📝 License
Distributed under the MIT License. See `LICENSE` for more information.

---

*Made with ❤️ for movie lovers.*
