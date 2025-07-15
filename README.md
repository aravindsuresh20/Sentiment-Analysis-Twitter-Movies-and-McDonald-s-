# Sentiment-Analysis-Twitter-Movies-and-McDonald-s-
# 💬 Individual Sentiment Analysis Scripts

This project consists of a collection of **independent Jupyter Notebooks**, each dedicated to performing **sentiment analysis** on a specific dataset (McDonald’s reviews, movie ratings, or tweets). Each notebook processes an Excel file, computes sentiment, and exports a new Excel file with sentiment results and **conditional formatting** for intuitive visualization.

---

## 🚀 Features

- **📁 Self-Contained Notebooks**
  - `mcdonaldsdashbaord.ipynb`: Analyzes McDonald’s text reviews.
  - `movies.ipynb`: Applies logic to classify sentiment from numeric ratings.
  - `twitter_dashboard.ipynb`: Analyzes tweet sentiment using NLP.

- **🧠 Sentiment Classification**
  - `TextBlob` used in McDonald’s and Twitter notebooks for polarity-based sentiment.
  - Custom logic in the movie notebook for converting ratings to sentiment.

- **📊 Sentiment Output Columns**
  - `sentiment`: -1 (negative), 0 (neutral), 1 (positive)
  - `sentiment_percentage`: Normalized score
  - `sentiment_score`: Raw score or polarity

- **🎨 Excel Export with Conditional Formatting**
  - Green for positive
  - Yellow for negative
  - White for neutral

- **🔎 Quick Dataset Summary**
  Each notebook prints:
  - Dataset shape
  - Data types
  - Missing values
  - Basic statistics

---

## 🧰 Requirements

> ⚠️ There is **no built-in `requirements.txt`** in this project. You must **create it manually**.

Required Python libraries:

- `pandas`
- `textblob`
- `openpyxl`
- `jupyter`

To install manually:

```bash
pip install pandas textblob openpyxl jupyter
python -m textblob.download_corpora
```

If you'd like to create a `requirements.txt`, here's the content:

```txt
pandas
textblob
openpyxl
jupyter
```


---

## 📁 Project Structure

```
/your-project-directory/
│── mcdonaldsdashbaord.ipynb            # McDonald's review sentiment analysis
│── movies.ipynb                        # Movie ratings sentiment classification
│── twitter_dashboard.ipynb             # Twitter data sentiment analysis
│── README.md                           # This documentation
│
└── /dashbaord/
    ├── McDonald_s_Reviews.xlsx         # Input for McDonald’s analysis
    ├── n_movies.xlsx                   # Input for movie ratings
    ├── twitter_dataset.xlsx            # Input for tweet data
    ├── McDonald_s_coloured.xlsx        # Output with sentiment + formatting
    ├── n_movies_coloured.xlsx          # Output with sentiment + formatting
    └── twitter_dataset_1.xlsx          # Output with sentiment + formatting
```

---

## ⚙️ Installation & Setup

1. **Clone the Repository**

```bash
git clone https://github.com/yourusername/sentiment-analysis-scripts.git
cd sentiment-analysis-scripts
```

2. **Add Input Files**  
Place the following Excel files inside the `/dashbaord/` folder:
- `McDonald_s_Reviews.xlsx` — must contain a `review` column
- `n_movies.xlsx` — must contain a `rating` column
- `twitter_dataset.xlsx` — must contain a `Text` column

3. **Set Up Environment (Optional but Recommended)**

```bash
python -m venv venv
# Activate:
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate
```

4. **Install Dependencies**

```bash
pip install pandas textblob openpyxl jupyter
python -m textblob.download_corpora
```

---

## 💡 Usage Instructions

Launch the Jupyter Notebook environment:

```bash
jupyter notebook
```

### ▶️ To Analyze McDonald’s Reviews

- Open: `mcdonaldsdashbaord.ipynb`
- Input file: `/dashbaord/McDonald_s_Reviews.xlsx`
- Output: `/dashbaord/McDonald_s_coloured.xlsx`

### 🎬 To Analyze Movie Ratings

- Open: `movies.ipynb`
- Input file: `/dashbaord/n_movies.xlsx`
- Output: `/dashbaord/n_movies_coloured.xlsx`

### 🐦 To Analyze Twitter Data

- Open: `twitter_dashboard.ipynb`
- Input file: `/dashbaord/twitter_dataset.xlsx`
- Output: `/dashbaord/twitter_dataset_1.xlsx`

---

## 📦 Output

Each notebook will generate a new Excel file with:
- New sentiment-related columns
- Color-coded sentiment results (green/yellow/white)
- File saved in the same `/dashbaord/` directory

---
