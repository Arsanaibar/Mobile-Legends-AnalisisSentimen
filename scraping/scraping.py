from google_play_scraper import reviews
import openpyxl

def fetch_all_reviews(app_id, num_reviews):
    """
    Args:
    - app_id (str): ID aplikasi (contoh: com.mobile.legends).
    - num_reviews (int): Jumlah total ulasan yang ingin diambil.
    
    Returns:
    - List[str]: Daftar teks ulasan.
    """
    all_reviews = []
    for rating in range(5, 0, -1):  
        print(f"Mengambil ulasan dengan bintang {rating}...")
        result, _ = reviews(
            app_id,
            lang='id',  
            country='id',  
            count=num_reviews,
            filter_score_with=rating  
        )
        all_reviews.extend([r['content'] for r in result])
    return all_reviews

def save_reviews_to_excel(reviews, output_file):
    """    
    Args:
    - reviews (List[str]): Daftar ulasan.
    - output_file (str): Nama file Excel untuk menyimpan hasil.
    """
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Reviews"

    for review in reviews:
        sheet.append([review])

    workbook.save(output_file)
    print(f"Ulasan telah disimpan di file: {output_file}")

app_id = "com.mobile.legends"  
num_reviews_per_rating = 2000 
output_file = "data\mobile_legends_reviews_combined.xlsx"  

all_reviews = fetch_all_reviews(app_id, num_reviews_per_rating)

save_reviews_to_excel(all_reviews, output_file)
