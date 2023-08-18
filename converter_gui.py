import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tabula
import pandas as pd
 
class PDFtoExcelConverter:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PDF to Excel Converter")
 
        self.pdf_label = ttk.Label(root, text="PDF Path:")
        self.pdf_label.pack()
 
        self.pdf_path = ttk.Entry(root, width=40)
        self.pdf_path.pack()
 
        self.pages_label = ttk.Label(root, text="Page Range (e.g., 3-26):")
        self.pages_label.pack()
 
        self.pages = ttk.Entry(root, width=20)
        self.pages.pack()
 
        self.excel_label = ttk.Label(root, text="Excel Output Path:")
        self.excel_label.pack()
 
        self.excel_path = ttk.Entry(root, width=40)
        self.excel_path.pack()
 
        self.convert_button = ttk.Button(root, text="Convert", command=self.convert)
        self.convert_button.pack()
 
    def convert(self) -> None:
        pdf_path = self.pdf_path.get()
        pages = self.pages.get()
        excel_output_path = self.excel_path.get()
 
        try:
            tables = tabula.read_pdf(pdf_path, pages=pages, multiple_tables=True)
            table_df = pd.concat(tables, ignore_index=True)
            table_df.to_excel(excel_output_path, index=False)
            self.show_message("Conversion complete!", "PDF table has been converted and saved to Excel.")
        except Exception as e:
            self.show_message("Error", f"An error occurred: {e}")
 
    def show_message(self, title: str | None, message: str | None) -> None:
        messagebox.showinfo(title, message)
 
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoExcelConverter(root)
    root.mainloop()