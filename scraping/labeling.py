from transformers import pipeline
import pandas as pd

def load_reviews(file_path):
    """
    Membaca ulasan dari file Excel.
    
    Args:
    - file_path (str): Path ke file Excel yang berisi ulasan.
    
    Returns:
    - List[str]: Daftar ulasan.
    """
    df = pd.read_excel(file_path)
    return df['Ulasan'].tolist()  # Pastikan kolom ulasan bernama 'Ulasan'

def label_reviews(reviews):
    """
    Melabeli sentimen ulasan menggunakan model IndoBERT.
    
    Args:
    - reviews (List[str]): Daftar ulasan.
    
    Returns:
    - List[str]: Daftar label sentimen ('positif', 'negatif', 'netral').
    """
    sentiment_model = pipeline("sentiment-analysis", model="w11wo/indonesian-roberta-base-sentiment-classifier")
    
    labels = []
    for review in reviews:
        result = sentiment_model(review)
        labels.append(result[0]['label'].lower())  # Label: 'positive', 'negative', 'neutral'
    return labels

def save_labeled_reviews(reviews, labels, output_file):
    """
    Menyimpan ulasan dan label ke file Excel.
    
    Args:
    - reviews (List[str]): Daftar ulasan.
    - labels (List[str]): Daftar label sentimen.
    - output_file (str): Path file Excel untuk menyimpan hasil.
    """
    df = pd.DataFrame({
        'Ulasan': reviews,
        'Sentimen': labels
    })
    df.to_excel(output_file, index=False)
    print(f"Data telah disimpan ke file: {output_file}")

# Contoh penggunaan
input_file = "mobile_legends_reviews_combined.xlsx"  # File input ulasan
output_file = "mobile_legends_reviews_labeled.xlsx"  # File output ulasan berlabel

# Langkah-langkah
print("Membaca ulasan dari file...")
reviews = load_reviews(input_file)

print("Melabeli ulasan...")
labels = label_reviews(reviews)

print("Menyimpan ulasan berlabel ke file Excel...")
save_labeled_reviews(reviews, labels, output_file)
