import os
try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("Please install python-docx using: pip install python-docx")
    exit()

def create_report():
    doc = Document()
    
    # --- Title Page ---
    title = doc.add_heading('Hybrid Movie Recommendation Engine: Integrating Collaborative Filtering, Matrix Factorization, and Content-Based Methods', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    author_p = doc.add_paragraph()
    author_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    author_p.add_run('[Your Name(s) Here]\n').bold = True
    author_p.add_run('[Name of Institution]\n')
    author_p.add_run('[Date]')
    
    doc.add_page_break()

    # --- Abstract ---
    doc.add_heading('Abstract', level=1)
    doc.add_paragraph(
        "With the rapid growth of digital media, recommendation systems have become crucial for filtering "
        "information and personalizing user experiences. This project presents a full-stack movie recommendation "
        "application that employs a hybrid algorithm integrating Collaborative Filtering, Matrix Factorization (Truncated SVD), "
        "and Content-Based Filtering. By leveraging cosine similarity across user-movie ratings, latent feature embeddings, "
        "and genre vectors, the system delivers highly tailored recommendations. The application also introduces an "
        "Explainable AI feature and an interactive network visualization of movie similarities using D3. Evaluation relies on "
        "RMSE, MAE, Precision@10, and Recall@10. This report outlines the development, methodology, and empirical results "
        "of the proposed system."
    )

    # --- Table of Contents ---
    doc.add_heading('Table of Contents', level=1)
    toc_items = [
        "1. Introduction", 
        "2. Literature Review", 
        "3. Methodology", 
        "4. Results and Discussion", 
        "5. Conclusion", 
        "6. References", 
        "7. Appendices"
    ]
    for item in toc_items:
        doc.add_paragraph(item)
    doc.add_paragraph("\n(Note: Please use Word's native Table of Contents generator for accurate page numbers.)")
    
    doc.add_page_break()

    # --- 1. Introduction ---
    doc.add_heading('1. Introduction', level=1)
    doc.add_paragraph(
        "The abundance of movies available on modern streaming platforms often leads to information overload, "
        "making it difficult for users to select content aligned with their preferences. The primary objective "
        "of this project is to construct a scalable and accurate movie recommendation engine. The specific aims are:"
    )
    obj_list = [
        "To develop a hybrid recommendation model avoiding the cold-start and sparsity problems inherent in singular models.",
        "To provide interpretable explanations for why a movie is recommended (Explainable AI).",
        "To visualize the proximity of unstructured movie data using a force-directed interactive graph.",
        "To construct a responsive full-stack application (FastAPI backend and React frontend)."
    ]
    for obj in obj_list:
        doc.add_paragraph(obj, style='List Bullet')
    doc.add_paragraph(
        "The scope of this project covers algorithm design, model evaluation, and software engineering. It holds "
        "significance by demonstrating how multiple mathematical similarity measures can be harmonized to yield "
        "a superior User Experience."
    )

    # --- 2. Literature Review ---
    doc.add_heading('2. Literature Review', level=1)
    doc.add_paragraph(
        "Traditional recommendation systems are primarily bifurcated into Content-Based Filtering (CBF) and "
        "Collaborative Filtering (CF). CBF relies on item metadata, such as genres, but often suffers from over-specialization. "
        "CF leverages historical user interactions (user-item matrices) but struggles with high sparsity and the 'cold-start' issue. "
        "Matrix Factorization techniques, notably Singular Value Decomposition (SVD), have been popularized since the Netflix Prize "
        "for their ability to deduce latent features and lower dimensionality. "
    )
    doc.add_paragraph(
        "This project builds upon existing research by implementing a weighted Hybrid System. Furthermore, it incorporates Explainable "
        "AI principles. While many models operate as 'black boxes', this system correlates high similarity scores back to human-readable "
        "reasons, establishing trust while mitigating the limitations of standalone algorithms."
    )

    # --- 3. Methodology ---
    doc.add_heading('3. Methodology', level=1)
    doc.add_paragraph(
        "The system architecture follows a client-server paradigm, utilizing React for the frontend and FastAPI (Python) "
        "for the backend. The recommendation engine computes scores across three distinct dimensions before aggregating them:"
    )
    
    doc.add_heading('3.1 Algorithmic Approach', level=2)
    doc.add_paragraph(
        "1. Collaborative Similarity: A user-movie rating matrix is constructed, pivoting user interactions. The transposed "
        "matrix is then passed through a Cosine Similarity function to yield a movie-by-movie similarity matrix.\n"
        "2. Matrix Factorization (SVD): To capture latent features, Truncated Singular Value Decomposition (TruncatedSVD) is "
        "applied to the matrix, capped at 50 components. Embeddings are transformed, and cosine similarity is extracted.\n"
        "3. Content-Based Similarity: MultiLabel Binarizer is used to transform movie genres into vectors, followed by "
        "cosine similarity computations."
    )
    
    doc.add_heading('3.2 Hybrid Weighting Formulation', level=2)
    doc.add_paragraph(
        "The final recommendation list is generated by combining the normalized similarity scores. The weights configuration is dynamically "
        "set as: Final Score = (0.5 * Collaborative) + (0.3 * SVD) + (0.2 * Content). This formulation ensures user behaviors "
        "have the strongest impact, supported mathematically by latent relationships, and grounded by actual genre semantics."
    )
    
    doc.add_heading('3.3 Experimental Setup & Tools', level=2)
    doc.add_paragraph(
        "The engine operates on a dataset composed of 'movies.csv' and 'ratings.csv'. Backend calculations heavily utilize "
        "pandas for data manipulation and scikit-learn for mathematical modeling. The frontend visualizer integrates 'react-force-graph-2d' "
        "to cast spatial relationships of top-K elements."
    )

    # --- 4. Results and Discussion ---
    doc.add_heading('4. Results and Discussion', level=1)
    doc.add_heading('4.1 Exploratory Data Analysis (EDA)', level=2)
    if os.path.exists('backend/rating_distribution.png'):
        doc.add_picture('backend/rating_distribution.png', width=Inches(5.5))
    if os.path.exists('backend/top_genres.png'):
        doc.add_picture('backend/top_genres.png', width=Inches(5.5))

    doc.add_heading('4.2 Predictive Performance', level=2)
    doc.add_paragraph(
        "Model performance is verified using an 80/20 train-test split executed over the ratings dataset. The performance matrix encapsulates "
        "statistical error rates and ranking accuracy algorithms (Precision and Recall at K=10). "
        "Metrics analyzed:"
    )
    eval_list = [
        "Root Mean Squared Error (RMSE): Indicates the standard deviation of prediction errors.",
        "Mean Absolute Error (MAE): Indicates the average absolute disparity between predicted user ratings (derived via SVD) and actual test ratings.",
        "Precision@10: Shows the fraction of the top 10 recommended movies that the user rated favorably (>=4.0).",
        "Recall@10: Reflects the proportion of total favorable items successfully retrieved within the top 10."
    ]
    for eval_i in eval_list:
        doc.add_paragraph(eval_i, style='List Bullet')
        
    if os.path.exists('backend/error_metrics.png'):
        doc.add_picture('backend/error_metrics.png', width=Inches(5.5))
    if os.path.exists('backend/retrieval_metrics.png'):
        doc.add_picture('backend/retrieval_metrics.png', width=Inches(5.5))
        
    doc.add_paragraph(
        "The system robustly handles edge cases, effectively merging the three matrices. The similarity thresholds "
        "reliably triggered the Explainability modules, successfully returning explanations such as 'Similar viewing pattern in latent features' "
        "when SVD scores exceeded 0.3. A limitation discovered during data processing was the matrix sparsity, which justified our reliance "
        "on Truncated SVD to predict null values before raw collaborative similarities dropped to zero."
    )

    # --- 5. Conclusion ---
    doc.add_heading('5. Conclusion', level=1)
    doc.add_paragraph(
        "In conclusion, the developed full-stack movie recommendation system successfully satisfies the primary objectives of "
        "delivering high-quality, personalized, and interpretable movie recommendations. By engineering a hybrid algorithm weighting Collaborative "
        "Filtering, latent feature extraction via SVD, and Content-based metrics, the model bypasses standard limitations of single-node recommenders. "
        "Furthermore, by wrapping the algorithm in a modern UI with active data visualization, the project bridges the gap between raw data science and user-centric software engineering."
    )
    doc.add_paragraph(
        "Future work may include integrating Deep Learning architectures, such as Neural Collaborative Filtering (NCF) or Autoencoders, "
        "to substitute the SVD layer. Additionally, migrating the application to real-time streaming data via Apache Kafka would significantly enhance scalability."
    )

    # --- 6. References ---
    doc.add_heading('6. References', level=1)
    refs = [
        "[1] F. Maxwell Harper and Joseph A. Konstan. 2015. The MovieLens Datasets: History and Context. ACM Transactions on Interactive Intelligent Systems.",
        "[2] Yehuda Koren, Robert Bell, and Chris Volinsky. 2009. Matrix Factorization Techniques for Recommender Systems. Computer 42, 8 (Aug. 2009), 30-37.",
        "[3] Pedregosa et al. 2011. Scikit-learn: Machine Learning in Python. JMLR 12, pp. 2825-2830.",
        "[4] FastAPI Documentation. [Online]. Available: https://fastapi.tiangolo.com/"
    ]
    for ref in refs:
        doc.add_paragraph(ref)

    # --- 7. Appendices ---
    doc.add_heading('7. Appendices', level=1)
    doc.add_paragraph("Appendix A: Source Code Structure")
    doc.add_paragraph(
        "- backend/recommender.py (Contains the hybrid weighting mechanism and Explainable AI functions)\n"
        "- backend/evaluation.py (Contains metric computation for RMSE, MAE)\n"
        "- frontend/src (Contains the React UI and Force Graph integration)"
    )

    doc_name = 'Project_Report.docx'
    doc.save(doc_name)
    print(f"Report successfully generated as '{doc_name}' in the current directory.")

if __name__ == "__main__":
    create_report()
