import tkinter as tk
from tkinter import filedialog
import pandas as pd

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to XLSX Converter")
        self.root.geometry("700x200")

        self.file_path = ""
        self.status_var = tk.StringVar()
        self.conversion_status_var = tk.StringVar()

        self.path_frame = tk.Frame(self.root)
        self.path_frame.pack(fill=tk.X, padx=20, pady=(20, 10))

        self.path_label = tk.Label(self.path_frame, text="Folder:")
        self.path_label.pack(side=tk.LEFT, padx=(0, 10))

        self.path_entry = tk.Entry(self.path_frame, textvariable=self.status_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.choose_button = tk.Button(self.path_frame, text="Search File", command=self.choose_csv_file)
        self.choose_button.pack(side=tk.LEFT, padx=(10, 0))

        self.convert_button = tk.Button(self.root, text="Convert to XLSX", command=self.convert_csv_to_xlsx)
        self.convert_button.pack(padx=20, pady=(10, 0))

        self.conversion_status_label = tk.Label(self.root, textvariable=self.conversion_status_var)
        self.conversion_status_label.pack(fill=tk.X, padx=20, pady=(0, 20))

    def choose_csv_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        self.status_var.set(self.file_path)
        self.conversion_status_var.set("")  # Limpar a mensagem de status anterior

    def convert_csv_to_xlsx(self):
        if self.file_path:
            df = pd.read_csv(self.file_path)
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if output_path:
                df.to_excel(output_path, index=False)
                self.conversion_status_var.set("Conversion Complete!")

if __name__ == "__main__":
    app_root = tk.Tk()
    app = ConverterApp(app_root)
    app_root.mainloop()
