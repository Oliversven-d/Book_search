import pandas as pd  # データ処理のためのPandasライブラリをインポート
import tkinter as tk  # GUI作成のためのTkinterライブラリをインポート
from tkinter import ttk, messagebox  # ttkとmessageboxをインポート
import random  # ランダムな選択のためのrandomライブラリをインポート
import webbrowser  # Webブラウザを制御するためのwebbrowserライブラリをインポート
import nltk  # 自然言語処理のためのNLTKライブラリをインポート
from janome.tokenizer import Tokenizer  # 日本語の形態素解析のためのJanomeライブラリからTokenizerをインポート 
from sklearn.feature_extraction.text import TfidfVectorizer  # TF-IDFベクトル化のためのライブラリをインポート
from sklearn.metrics.pairwise import linear_kernel  # 線形カーネルを使用するためのライブラリをインポート
from nltk.corpus import wordnet  # WordNetを使用するためのライブラリをインポート
import openpyxl
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk  
import requests
from io import BytesIO
from sklearn.cluster import KMeans
from sklearn.neighbors import NearestNeighbors
import os

API_KEY = 'AIzaSyCDVgGl2LEGeivkhAJQRjTlNg0-FpcWCvQ' 

# Download NLTK data # NLTKのデータをダウンロード

nltk.download('wordnet')

file_path = os.path.join(os.path.dirname(__file__), 'PS＿６階＿本棚.xlsx')

# Google Sheetsに接続するための認証情報を設定
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']


# Initialize tokenizers for the languages I want to search for the words
# 異なる言語用のトークナイザーを初期化
tokenizer_jp = Tokenizer()
tokenizer_en = nltk.word_tokenize
tokenizer_de = nltk.word_tokenize

# Set to keep track of previously recommended books
previously_recommended_books = set()

 
def update_search_box_with_associated_word(word):
    search_var.set(word)

def fetch_book_description(book_title):
    try:
        url = f"https://www.googleapis.com/books/v1/volumes?q={book_title}&key={API_KEY}"
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            if "items" in data and len(data["items"]) > 0:
                description = data["items"][0]["volumeInfo"].get("description", "No description available.")
                return description
        return "No description available."
    except Exception as e:
        return f"Error fetching description: {str(e)}"


def ai_recommended_books(query, books_df, num_recommendations=5):
    # Tokenize the query and create the TF-IDF vector
    tfidf_vectorizer = TfidfVectorizer()
    text_columns = (
        books_df['タイトル'].astype(str).fillna('') + ' ' +
        books_df['ジャンル'].astype(str).fillna('') + ' ' +
        books_df['著者'].astype(str).fillna('')
    )
    tfidf_matrix = tfidf_vectorizer.fit_transform(text_columns)
    
    # Fit a NearestNeighbors model to the TF-IDF matrix
    nn = NearestNeighbors(n_neighbors=num_recommendations + 1, metric='cosine')
    nn.fit(tfidf_matrix)
    
    # Tokenize and vectorize the query
    query_vector = tfidf_vectorizer.transform([query])
    
    # Find the nearest neighbors
    distances, indices = nn.kneighbors(query_vector)
    
    # Exclude the query itself from the recommendations
    recommended_indices = indices.flatten()[1:]
    
    # Get recommended book titles
    recommended_books = books_df.iloc[recommended_indices]['タイトル'].tolist()
    
    return recommended_books[:num_recommendations]


def show_book_description(book_title):
    description = fetch_book_description(book_title)
    
    description_window = tk.Toplevel(root)
    description_window.title("Book Description")
    
    description_label = tk.Label(description_window, text=description, wraplength=400, justify="left")
    description_label.pack(padx=10, pady=10)

def change_cell_content(cell_id, new_content):   
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        sheet[cell_id] = new_content
        wb.save(file_path)
        print(f"Cell {cell_id} successfully changed to '{new_content}'.")
        return True
    except KeyError as e:
        print(f"Error: cell-id '{cell_id}' is not right.")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False

def update_cell():
    cell_id = cell_entry.get()
    
    # Falls cell_id bereits mit 'Y' beginnt, keine Änderung vornehmen
    if not cell_id.startswith('Z'):
        cell_id = cell_id  # 'Y' an die Zellnummer anhängen
    
    # Inhalt der Eingabebox aktualisieren
    cell_entry.delete(0, tk.END)  # Eingabebox leeren
    cell_entry.insert(0, cell_id)  # Neue Zellen-ID mit 'Y' einfügen
    
    new_content = content_entry.get()
    print(f"Ändere Inhalt der Zelle {cell_id} zu: {new_content}")  # Debugging-Ausgabe
    
    success = change_cell_content(cell_id, new_content)
    if success:
        messagebox.showinfo("Success", f"The Cell {cell_id} was changed.")
    else:
        messagebox.showerror("Error", f"The {cell_id} was not changed.")


def tokenize_text(text, lang):
    # Tokenize text based on language / 言語に基づいてテキストをトークン化する
    if lang == 'jpn':
        return tokenizer_jp.tokenize(text, wakati=True)  # Tokenize with Janome for Japanese / 日本語の場合はJanomeで単語に分割する
    elif lang == 'eng':
        return tokenizer_en(text)  # Tokenize with NLTK for English / 英語の場合はNLTKで単語に分割する
    elif lang == 'deu':
        return tokenizer_de(text)  # Tokenize with NLTK for German / ドイツ語の場合はNLTKで単語に分割する
    else:
        return []

def open_table(book_info):  # Open the Google Sheets URL / Google SheetsのURLを開く
    # Load book data from an Excel file / Excelファイルから書籍データを読み込む
    google_sheets_url = 'https://docs.google.com/spreadsheets/d/10B8Wpk1M8iuGvI5vryIHMIG2SO_Qj-Kjn0M02Y3fgeg/edit?gid=0#gid=0'
    try:
        webbrowser.open_new(google_sheets_url)
        messagebox.showinfo("Information", "The Google Sheets-Document was opened.")
    except Exception as e:
        messagebox.showerror("Error", f"Error for: {str(e)}")

def load_books(excel_file):
    df = pd.read_excel(excel_file, sheet_name=1)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    return df

def select_random_books(books_list, num_books=5):   # Select random books from a list / リストからランダムに書籍を選択する
    return random.sample(books_list, min(num_books, len(books_list)))


def cluster_books(books_df, num_clusters=10):
    tfidf_vectorizer = TfidfVectorizer()
    text_columns = (
        books_df['タイトル'].astype(str).fillna('') + ' ' +
        books_df['ジャンル'].astype(str).fillna('') + ' ' +
        books_df['著者'].astype(str).fillna('')
    )
    tfidf_matrix = tfidf_vectorizer.fit_transform(text_columns)
    
    
    kmeans = KMeans(n_clusters=num_clusters, random_state=42)
    kmeans.fit(tfidf_matrix)
    books_df['Cluster'] = kmeans.labels_
    
    return books_df

def recommend_books_in_cluster(book_title, books_df):
    try:
        book_cluster = books_df[books_df['タイトル'] == book_title]['Cluster'].values[0]
        similar_books = books_df[books_df['Cluster'] == book_cluster]['タイトル'].tolist()
        similar_books.remove(book_title)
        return similar_books[:10]
    except IndexError:
        return []

def recommend_books_tfidf(query, lang):
    tfidf_vectorizer = TfidfVectorizer()
    text_columns = (
        books['タイトル'].astype(str).fillna('') + ' ' +
        books['ジャンル'].astype(str).fillna('') + ' ' +
        books['著者'].astype(str).fillna('')
    )
    tfidf_matrix = tfidf_vectorizer.fit_transform(text_columns)
    cosine_similarities = linear_kernel(tfidf_matrix, tfidf_matrix)

    queries = query.split()  # Split the query into individual words
    indices = []

    for q in queries:
        if lang == 'jpn':
            matching_indices = books[books['タイトル'].str.contains(q, case=False, na=False, regex=False)].index.tolist()
        elif lang == 'eng':
            matching_indices = books[books['タイトル'].str.contains(q, case=False, na=False, regex=False)].index.tolist()
        elif lang == 'deu':
            matching_indices = books[books['タイトル'].str.contains(q, case=False, na=False, regex=False)].index.tolist()

        if matching_indices:
            indices.extend(matching_indices)

    indices = list(set(indices))  # Remove duplicates

    listbox.delete(0, tk.END)
    recommended_books = []
    if indices:
        for idx in indices:
            similar_books = list(enumerate(cosine_similarities[idx]))
            similar_books = sorted(similar_books, key=lambda x: x[1], reverse=True)
            recommended_books += [books.iloc[book_index]['タイトル'] for book_index, _ in similar_books[1:11]]

        recommended_books = list(set(recommended_books))  # Remove duplicates

        if recommended_books:
            listbox.insert(tk.END, "TF-IDFベクトルに基づくおすすめの本:")
            random_recommended_books = select_random_books(recommended_books)
            for book in random_recommended_books:
                listbox.insert(tk.END, book)
            listbox.insert(tk.END, "")
        else:
            listbox.insert(tk.END, "見つかりませんでした。")

        ai_recommended = ai_recommended_books(query, books)
        if ai_recommended:
            listbox.insert(tk.END, "AIが推奨する本:")
            for book in ai_recommended:
                listbox.insert(tk.END, book)
            listbox.insert(tk.END, "")
    else:
        listbox.insert(tk.END, "見つかりませんでした。")

    associated_words = find_associated_words(query, lang, topn=20)
    recommended_associated_words = set()
    if associated_words:
        listbox.insert(tk.END, "推奨される関連単語:")
        for word in associated_words:
            if isinstance(word, str) and word.lower() not in recommended_associated_words and word.lower() not in query.lower():
                listbox.insert(tk.END, word)
                recommended_associated_words.add(word.lower())
        listbox.insert(tk.END, "")



    # Use the recommended_books from TF-IDF as exclude_titles
    recommend_additional_books_tfidf(query, lang, recommended_books)


def search_books():
    search_term = search_var.get().strip()
    if not search_term:
        messagebox.showwarning("Input Error", "Please enter a search term.")
        return

    lang = 'jpn'
    if search_term.isascii():
        if search_term.isalpha():
            lang = 'eng'
        else:
            lang = 'deu'

    listbox.delete(0, tk.END)
    recommend_books_tfidf(search_term, lang)


def detect_language(query):
    if all(ord(char) < 128 for char in query):
        return 'eng'  # Assume English if all characters are ASCII
    elif any(char.isdigit() for char in query):
        return 'deu'  # Assume German if query contains digits
    else:
        return 'jpn'  # Default to Japanese if not English or German

def find_associated_words_and_search(query, lang, pos=None, topn=20): # Find associated words and perform search / 関連する単語を見つけて検索を実行する
    associated_words = []
    if lang == 'jpn':
        synsets = wordnet.synsets(query, pos=pos, lang='jpn')
        for synset in synsets:
            for lemma in synset.lemmas(lang='jpn'):
                associated_words.append(lemma.name())
    elif lang == 'eng':
        synsets = wordnet.synsets(query)
        for synset in synsets:
            for lemma in synset.lemmas():
                associated_words.append(lemma.name())
    elif lang == 'deu':
        synsets = wordnet.synsets(query, lang='deu')
        for synset in synsets:
            for lemma in synset.lemmas(lang='deu'):
                associated_words.append(lemma.name())

    if associated_words:
        search_books_with_associated_word(associated_words[0].lower())
    return associated_words[:topn]
    

def search_books_with_associated_word(query):  # Search for books using associated words / 関連する単語を使用して書籍を検索する
    lang = detect_language(query)
    recommend_books_tfidf(query, lang)


def recommend_books(query, lang, exclude_titles): # Recommend books based on associated words / 関連する単語に基づいて書籍を推薦する
    recommended_books = []
    associated_words = find_associated_words(query, lang, topn=20)
    if associated_words:
        for word in associated_words:
            if isinstance(word, str):
                results = books[
                    books['タイトル'].str.contains(word, case=False, na=False, regex=False) |
                    books['ジャンル'].str.contains(word, case=False, na=False, regex=False) |
                    books['著者'].str.contains(word, case=False, na=False, regex=False) |
                    books['翻訳者'].str.contains(word, case=False, na=False, regex=False) |
                    books['言語'].str.contains(word, case=False, na=False, regex=False) |
                    books['出版年'].astype(str).str.contains(word, case=False, na=False, regex=False) |
                    books['出版社'].str.contains(word, case=False, na=False, regex=False)
                ]
                if not results.empty:
                    for title in results['タイトル'].tolist():
                        if title.lower() not in exclude_titles:
                            recommended_books.append(title)
                            if len(recommended_books) >= 20:
                                break
                if len(recommended_books) >= 20:
                    break
    return recommended_books[:5]



def recommend_additional_books_tfidf(query, lang, exclude_titles):  
    # Recommend additional books using TF-IDF Vectors
    additional_books = recommend_books(query, lang, exclude_titles)
    unique_books = set(additional_books)  # Set to keep unique book titles
    
    # Ensure additional books are not in the excluded titles
    filtered_books = [book for book in unique_books if book not in exclude_titles]
    
    if filtered_books:
        if lang == 'jpn':
            listbox.insert(tk.END, "以下の本に興味があるかもしれません:")
        elif lang == 'eng':
            listbox.insert(tk.END, "以下の本に興味があるかもしれません:")
        elif lang == 'deu':
            listbox.insert(tk.END, "以下の本に興味があるかもしれません:")
        
        for book in filtered_books:
            listbox.insert(tk.END, book)
        listbox.insert(tk.END, "")




def find_associated_words(query, lang, pos=None, topn=100): # Find associated words / 関連する単語を見つける
     
    associated_words = []
    if lang == 'jpn':
        synsets = wordnet.synsets(query, pos=pos, lang='jpn')
        for synset in synsets:
            for lemma in synset.lemmas(lang='jpn'):
                associated_words.append(lemma.name())
    elif lang == 'eng':
        synsets = wordnet.synsets(query, pos=pos, lang='eng')
        for synset in synsets:
            for lemma in synset.lemmas(lang='eng'):
                associated_words.append(lemma.name())
    elif lang == 'deu':
        synsets = wordnet.synsets(query, pos=pos, lang='deu')
        for synset in synsets:
            for lemma in synset.lemmas(lang='deu'):
                associated_words.append(lemma.name())
    
    return associated_words[:topn]


def show_book_cell(book_info):
    title = search_var.get().strip().lower()  # Get the search query and normalize it / 検索クエリを取得して正規化する
    if title:
        try:
            found = False
            for index, row in books.iterrows():
                if isinstance(row['タイトル'], str):  # Check if the title is a string
                    book_title = row['タイトル'].strip().lower()  # Normalize case and strip spaces
                    print(f"Searching for: '{title}', Checking: '{book_title}'")  # Debug print

                    if book_title == title:
                        # Get the book number to navigate to the corresponding row
                        book_number = row['ばんごう']

                        # Find the row index based on the book number
                        row_index = None
                        for idx, row in books.iterrows():
                            if row['ばんごう'] == book_number:
                                row_index = idx + 2  # Adjust for Excel row index (assuming 0-based index)

                        if row_index is not None:
                            # Assuming column F for demonstration (you may adjust this as needed)
                            col_index = 'G'

                            # Show message with cell information
                            messagebox.showinfo("Cell Information", f"The Cell '{row['タイトル']}' is in {row_index} and {col_index}.")

                            # Construct the URL with the specific cell range
                            google_sheets_url = 'https://docs.google.com/spreadsheets/d/10B8Wpk1M8iuGvI5vryIHMIG2SO_Qj-Kjn0M02Y3fgeg/edit?gid=0#gid=0'
                            cell_range = f'{col_index}{row_index}'
                            google_sheets_url_with_cell = f'{google_sheets_url}#gid=0&range={cell_range}'

                            try:
                                webbrowser.open_new(google_sheets_url_with_cell)
                                messagebox.showinfo("Information", "Google Sheets-Dokument was opened and you navigate to it.")
                                found = True
                                break
                            except Exception as e:
                                messagebox.showerror("Error", f"There was a problem when opening the Cell: {str(e)}")

            if not found:
                messagebox.showwarning("Warning", f"The book '{title}' was not found.")

        except Exception as e:
            messagebox.showerror("Error", f"Error for Book Search: {str(e)}")

    else:
        messagebox.showwarning("Warning", "Please enter a book title.")


def show_book_info(event):
    index = listbox.nearest(event.y)
    if index != -1:
        title_author = listbox.get(index)
        clicked_associated_word = listbox.get(index)
        search_var.set(clicked_associated_word) 
        

        if "Empfohlene Bücher basierend auf TF-IDF-Vektoren:" in title_author:
            word_to_copy = title_author.replace("おすすめの本:", "").strip()
            search_var.set(word_to_copy)
        elif "おすすめのキーワード:" in title_author:
            # Update search box with the associated word
            associated_word = title_author.replace("おすすめのキーワード:", "").strip()
            update_search_box_with_associated_word(associated_word)
        else:
            book_title = title_author.split(' von ')[0]
            book_info = books[books['タイトル'] == book_title]
            if not book_info.empty:
                info_window = tk.Toplevel(root)
                info_window.title("Book Information")

                cell_number = book_info.index.values[0] + 2
                cell_id = f"Z{cell_number}"

                # Ensure search_var is empty before setting new value
                search_var.set("")

                # Set the new cell_id
                search_var.set(cell_id)

                # Fetch book cover from Google Books API
                cover_url = fetch_book_cover(book_title)
                cover_image = None
                if cover_url:
                    try:
                        cover_response = requests.get(cover_url)
                        if cover_response.status_code == 200:
                            cover_image = Image.open(BytesIO(cover_response.content))
                            cover_image = cover_image.resize((150, 225), Image.LANCZOS)
                        else:
                            cover_url = ''
                    except Exception as e:
                        print(f"Error loading cover image: {e}")
                        cover_url = ''

                info_text = f"""
                タイトル: {book_info['タイトル'].values[0]}
                著者: {book_info['著者'].values[0]}
                ジャンル: {book_info['ジャンル'].values[0]}
                言語: {book_info['言語'].values[0]}
                出版年: {book_info['出版年'].values[0]}
                出版社: {book_info['出版社'].values[0]}
                リスト入力者のメモ（リスト上に記入）: {book_info['リスト入力者のメモ（リスト上に記入）'].values[0]}
                Cell Number: {'Z' + str(cell_number)}
                """

                info_label = tk.Label(info_window, text=info_text, justify="left")
                info_label.pack(padx=10, pady=10)

                if cover_image:
                    cover_photo = ImageTk.PhotoImage(cover_image)
                    cover_label = tk.Label(info_window, image=cover_photo)
                    cover_label.image = cover_photo
                    cover_label.pack(pady=5)
                else:
                    cover_label = tk.Label(info_window, text="No cover image available.")
                    cover_label.pack(pady=5)

                open_table_button = ttk.Button(info_window, text="リストを開く", command=lambda: open_table(book_info.iloc[0]))
                open_table_button.pack(pady=5)

                show_author_cell_button = ttk.Button(info_window, text="著者を検索する", command=lambda: show_author_cell(book_info.iloc[0]))
                show_author_cell_button.pack(pady=5)

              
                show_author_cell_button2 = ttk.Button(info_window, text="本のタイトルを検索する", command=lambda: show_author_cell2(book_info.iloc[0], books))
                show_author_cell_button2.pack(pady=5)

                global content_entry

                cell_entry.delete(0, tk.END)  # Clear the cell entry field
                cell_entry.insert(0, cell_id)  # Automatically fill the cell number

                content_label = tk.Label(info_window, text="コメントを書く:")
                content_label.pack(pady=5)

                content_entry = tk.Entry(info_window, width=30)
                content_entry.pack()

                button = tk.Button(info_window, text="概要を表示", command=update_cell)
                button.pack(pady=10)

                summary_button = ttk.Button(info_window, text="概要を表示", command=lambda: show_book_description(book_title))
                summary_button.pack(pady=5)

                info_window.mainloop()

 
def fetch_book_cover(title):
    url = f"https://www.googleapis.com/books/v1/volumes?q=intitle:{title}&key={API_KEY}"
    try:
        response = requests.get(url).json()
        if 'items' in response:
            book = response['items'][0]['volumeInfo']
            cover_url = book.get('imageLinks', {}).get('thumbnail', '')
            return cover_url
    except Exception as e:
        print(f"Error fetching book cover: {e}")
    return None

def show_author_cell(book_info, books):
    author = book_info['著者'].strip().lower()  # Get the author name and normalize it
    if author:
        try:
            found = False
            for index, row in books.iterrows():
                if isinstance(row['著者'], str):  # Check if the author is a string
                    book_author = row['著者'].strip().lower()  # Normalize case and strip spaces
                    print(f"Searching for: '{author}', Checking: '{book_author}'")  # Debug print

                    if book_author == author:
                        # Adjust for Excel row index (assuming 0-based index)
                        row_index = index + 2  # Adjust this if needed

                        # Assuming column J for author cell (you may adjust this as needed)
                        col_index = 'J'

                        # Show message with cell information
                        messagebox.showinfo("Cell Information", f"The cell for '{row['著者']}' is in row {row_index}, column {col_index}.")

                        # Construct the URL with the specific cell range
                        google_sheets_url = 'https://docs.google.com/spreadsheets/d/10B8Wpk1M8iuGvI5vryIHMIG2SO_Qj-Kjn0M02Y3fgeg/edit'
                        cell_range = f'{col_index}{row_index}'
                        google_sheets_url_with_cell = f'{google_sheets_url}#gid=0&range={cell_range}'

                        try:
                            webbrowser.open_new(google_sheets_url_with_cell)
                            messagebox.showinfo("Information", "The Google Sheet was opened and you navigated to the cell.")
                            found = True
                            break
                        except Exception as e:
                            messagebox.showerror("Error", f"There was a problem with opening the Google Sheet: {str(e)}")

            if not found:
                messagebox.showwarning("Warning", f"The author '{author}' was not found.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while searching for the author: {str(e)}")

    else:
        messagebox.showwarning("Warning", "Please enter an author name.")


def show_author_cell2(book_info, books):
    try:
        author_name = book_info['著者'].strip().lower()  # Get the author name and normalize it
        if author_name:
            found = False
            for index, row in books.iterrows():
                if isinstance(row['著者'], str):  # Check if the author is a string
                    book_author = row['著者'].strip().lower()  # Normalize case and strip spaces
                    print(f"Searching for: '{author_name}', Checking: '{book_author}'")  # Debug print

                    if book_author == author_name:
                        # Adjust for Excel row index (assuming 0-based index)
                        row_index = index + 2  # Adjust this if needed

                        # Assuming column G for author cell (you may adjust this as needed)
                        col_index = 'G'

                        # Show message with cell information
                        messagebox.showinfo("Cell Information", f"The cell for '{row['著者']}' is in row {row_index}, column {col_index}.")

                        # Construct the URL with the specific cell range
                        google_sheets_url = 'https://docs.google.com/spreadsheets/d/10B8Wpk1M8iuGvI5vryIHMIG2SO_Qj-Kjn0M02Y3fgeg/edit'
                        cell_range = f'{col_index}{row_index}'
                        google_sheets_url_with_cell = f'{google_sheets_url}#gid=0&range={cell_range}'

                        try:
                            webbrowser.open_new(google_sheets_url_with_cell)
                            messagebox.showinfo("Information", "The Google Sheet was opened and you navigated to the cell.")
                            found = True
                            break
                        except Exception as e:
                            messagebox.showerror("Error", f"There was a problem with opening the Google Sheet: {str(e)}")

            if not found:
                messagebox.showwarning("Warning", "Author not found.")

        else:
            messagebox.showwarning("Warning", "Please provide an author name.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while searching for the author: {str(e)}")

pass

# Load books data / 書籍データを読み込む
excel_file = 'PS＿６階＿本棚.xlsx'
books = load_books(excel_file)

# GUI setup / GUIの設定
root = tk.Tk()
root.title("Book Search")

search_var = tk.StringVar()
search_label = ttk.Label(root, text="本を探す:")
search_label.pack(pady=5)
search_entry = ttk.Entry(root, textvariable=search_var)
search_entry.pack(pady=5)

search_button = ttk.Button(root, text="検索する", command=search_books)
search_button.pack(pady=5)

cell_label = tk.Label(root, text="セルIDを入力する:")
cell_label.pack(pady=5)

cell_entry = tk.Entry(root, width=30)
cell_entry.pack()

content_label = tk.Label(root, text="セルの新しい内容:")
content_label.pack(pady=5)

content_entry = tk.Entry(root, width=30)
content_entry.pack()

button = tk.Button(root, text="概要を表示", command=update_cell)
button.pack(pady=10)
 
listbox = tk.Listbox(root, width=50, height=15)
listbox.pack(pady=10)
listbox.bind('<Button-1>', show_book_info)
listbox.bind("<Double-1>", show_book_info)

root.mainloop()
