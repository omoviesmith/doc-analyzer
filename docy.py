import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox
from textblob import TextBlob  # You can choose any other sentiment analysis library
import csv
import openpyxl
import PyPDF2
import re
import docx2txt


# Global variables
uploaded_document = None  # Define the global variable to hold the uploaded document content
positive_words = set()
negative_words = set()
highlighted_words = []

# UI Functions


def upload_document(uploaded_document):
    file_path = filedialog.askopenfilename(title="Upload Document to Analyze")
    if file_path:
        if file_path.lower().endswith('.pdf'):
            document_text = extract_pdf_text(file_path)
        elif file_path.lower().endswith('.docx'):
            document_text = docx2txt.process(file_path)  # Convert DOCX to plain text
        else:
            messagebox.showerror("Unsupported File", "Unsupported file format. Please select a plain text, PDF, or DOCX file.")
            return
        
        uploaded_document.delete("1.0", tk.END)
        uploaded_document.insert("1.0", document_text)


def extract_pdf_text(file_path):
    with open(file_path, "rb") as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
        return text
    
def analyze_document(uploaded_document, sentiment_label):
    global positive_words, negative_words
    if not uploaded_document:
        messagebox.showwarning("No Document", "Please upload a document first.")
        return
    
    doc_text = uploaded_document.get("1.0", tk.END)
    blob = TextBlob(doc_text)
    positive_hits = [word for word in blob.words if word in positive_words]
    negative_hits = [word for word in blob.words if word in negative_words]
    
    # Clear previous highlighting
    uploaded_document.tag_remove("positive", "1.0", tk.END)
    uploaded_document.tag_remove("negative", "1.0", tk.END)
    
    # Highlight positive and negative words in the document
    highlight_words(uploaded_document, positive_hits, "positive")
    highlight_words(uploaded_document, negative_hits, "negative")
    
    # Calculate sentiment score
    total_words = len(blob.words)
    net_positivity_score = (len(positive_hits) - len(negative_hits)) / total_words
    update_sentiment_score(sentiment_label, net_positivity_score, len(positive_hits), len(negative_hits))

def highlight_words(text_widget, word_list, tag):
    search_terms = set()
    for word in word_list:
        pattern = rf"\b{word}\b"
        search_results = re.finditer(pattern, text_widget.get("1.0", tk.END))
        for match in search_results:
            start = match.start()
            end = match.end()
            start_index = f"1.0+{start}c"
            end_index = f"1.0+{end}c"
            text_widget.tag_add(tag, start_index, end_index)
            search_terms.add(word)  # Collect the search terms
            highlighted_words.append(word)
    return search_terms

def find_and_replace(text_widget):
    def find_next():
        search_term = highlighted_words_combobox.get()
        start_index = text_widget.search(search_term, "insert", stopindex=tk.END, count=tk.StringVar())
        if start_index:
            end_index = f"{start_index}+{len(search_term)}c"
            text_widget.tag_add("search", start_index, end_index)
            text_widget.see(start_index)
            text_widget.focus()
            # Scroll to the word
            line_num, col_num = map(int, start_index.split('.'))
            text_widget.see(f"{line_num}.{col_num+1}")

    def replace():
        search_term = highlighted_words_combobox.get()
        replace_text = replace_entry.get()
        start_index = text_widget.search(search_term, "insert", stopindex=tk.END)
        if start_index:
            end_index = f"{start_index}+{len(search_term)}c"
            text_widget.delete(start_index, end_index)
            text_widget.insert(start_index, replace_text)
            text_widget.see(start_index)
            text_widget.focus()

    def replace_all():
        search_term = highlighted_words_combobox.get()
        replace_text = replace_entry.get()
        start_index = "1.0"
        while True:
            start_index = text_widget.search(search_term, start_index, stopindex=tk.END)
            if not start_index:
                break
            end_index = f"{start_index}+{len(search_term)}c"
            text_widget.delete(start_index, end_index)
            text_widget.insert(start_index, replace_text)
            text_widget.see(start_index)
            start_index = end_index
        text_widget.focus()

    def cancel():
        popup.destroy()

    popup = tk.Toplevel()
    popup.title("Find and Replace")
    popup.geometry("400x150")

    find_label = tk.Label(popup, text="Find what:")
    find_label.pack()
    
    highlighted_words_combobox = Combobox(popup, values=highlighted_words, state="readonly")
    highlighted_words_combobox.pack()

    replace_label = tk.Label(popup, text="Replace with:")
    replace_label.pack()

    replace_var = tk.StringVar()
    replace_var.set("")
    replace_entry = tk.Entry(popup, textvariable=replace_var)
    replace_entry.pack()

    find_button = tk.Button(popup, text="Find Next", command=find_next)
    find_button.pack(side="left")

    replace_button = tk.Button(popup, text="Replace", command=replace)
    replace_button.pack(side="left")

    replace_all_button = tk.Button(popup, text="Replace All", command=replace_all)
    replace_all_button.pack(side="left")

    cancel_button = tk.Button(popup, text="Cancel", command=cancel)
    cancel_button.pack(side="left")

    text_widget.tag_remove("search", "1.0", tk.END)

def clear_document(uploaded_document, sentiment_label):
    uploaded_document.delete("1.0", tk.END)
    uploaded_document.tag_remove("positive", "1.0", tk.END)
    uploaded_document.tag_remove("negative", "1.0", tk.END)
    update_sentiment_score(sentiment_label, 0.0, 0, 0)



def update_sentiment_score(label, score, pos_count, neg_count):
    label.config(text=f"Net Positivity Score: {score:.2f}\nPositive Words: {pos_count}\nNegative Words: {neg_count}")
    

def upload_dictionary():
    file_path = filedialog.askopenfilename(title="Select Sentiment Dictionary")
    if file_path:
        if file_path.lower().endswith('.csv'):
            process_csv(file_path)
        elif file_path.lower().endswith('.xlsx'):
            process_xlsx(file_path)
        else:
            messagebox.showerror("Unsupported File", "Unsupported file format. Please select a CSV or XLSX file.")

def process_csv(file_path):
    with open(file_path, "r", newline='', encoding='utf-8') as file:
        # csv_reader = csv.reader(file, delimiter='\t')
        csv_reader = csv.reader(file, delimiter=',') 
        try:
            process_data(csv_reader)
        except csv.Error as e:
            print(f"CSV Error: {e}")


def process_xlsx(file_path):
    workbook = openpyxl.load_workbook(file_path)
    worksheet = workbook.active
    rows = worksheet.iter_rows(values_only=True)
    process_data(rows)

def process_data(data):
    header_skipped = False  # Flag to track if the header row has been skipped
    for row in data:
        if not header_skipped:
            header_skipped = True
            continue  # Skip the header row
        
        print(row)  # For debugging
        word = row[0].lower()
        negative_value = int(row[7])
        positive_value = int(row[8])
        if negative_value > 0:
            negative_words.add(word)
        if positive_value > 0:
            positive_words.add(word)
    messagebox.showinfo("Dictionary Uploaded", "Sentiment dictionary uploaded successfully.")


# Main Function

def main():
    root = tk.Tk()
    root.title("Document Sentiment Analysis Tool")
    
    menu_bar = tk.Menu(root)
    root.config(menu=menu_bar)
    
    file_menu = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Upload Dictionary", command=upload_dictionary)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)

    edit_menu = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Edit", menu=edit_menu)
    edit_menu.add_command(label="Find and Replace", command=lambda: find_and_replace(doc_text))
    
    doc_frame = tk.Frame(root)
    doc_frame.pack(pady=10)

    # Create a scroll bar
    scroll_bar = tk.Scrollbar(root)
    scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)

    # Create a text widget and attach the scroll bar
    doc_text = tk.Text(root, height=15, width=80)
    doc_text.pack(pady=10)

    # Configure the text widget to use the scroll bar
    doc_text.config(yscrollcommand=scroll_bar.set)
    scroll_bar.config(command=doc_text.yview)

    doc_text.tag_config("positive", background="lightgreen")
    doc_text.tag_config("negative", background="lightcoral")
    
    upload_button = tk.Button(root, text="Upload Document", command=lambda: upload_document(doc_text))
    upload_button.pack(pady=10)

    # analyze_button = tk.Button(root, text="Analyze Document", command=lambda: analyze_document(doc_text))
    analyze_button = tk.Button(root, text="Analyze Document", command=lambda: analyze_document(doc_text, sentiment_label))
    analyze_button.pack()
    
    # clear_button = tk.Button(root, text="Clear Document", command=lambda: clear_document(doc_text))
    clear_button = tk.Button(root, text="Clear Document", command=lambda: clear_document(doc_text, sentiment_label))
    clear_button.pack()
    
    sentiment_label = tk.Label(root, text="Net Positivity Score: 0.00\nPositive Words: 0\nNegative Words: 0")
    sentiment_label.pack(pady=10)

    # find_replace_button = tk.Button(root, text="Find and Replace", command=lambda: find_and_replace(doc_text))
    # find_replace_button.pack()
    
    tk.mainloop()

if __name__ == "__main__":
    main()
