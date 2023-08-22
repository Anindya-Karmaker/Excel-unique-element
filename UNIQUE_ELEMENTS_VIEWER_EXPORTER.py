import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Viewer")

        self.load_button = ttk.Button(root, text="Load Excel File", command=self.load_excel)
        self.load_button.pack(pady=10)
        
        self.unique_elements_label0 = ttk.Label(root, text="Select the Sheet:")
        self.unique_elements_label0.pack(pady=5)
        
        self.sheets_listbox = tk.Listbox(root)
        self.sheets_listbox.pack(pady=10)
        self.sheets_listbox.bind("<<ListboxSelect>>", self.select_sheet)
        
        self.unique_elements_label0 = ttk.Label(root, text="Select the Column:")
        self.unique_elements_label0.pack(pady=5)

        self.columns_listbox = tk.Listbox(root)
        self.columns_listbox.pack(pady=10)
        self.columns_listbox.bind("<<ListboxSelect>>", self.select_column)

        self.unique_elements_label = ttk.Label(root, text="Unique Elements:")
        self.unique_elements_label.pack(pady=5)

        self.unique_elements_text = tk.Text(root, height=10, width=40)
        self.unique_elements_text.pack(pady=5)

        self.export_button = ttk.Button(root, text="Export Unique Elements", command=self.export_elements)
        self.export_button.pack(pady=10)

        self.total_elements_label = ttk.Label(root, text="Total Elements:")
        self.total_elements_label.pack(pady=5)

        self.total_elements_var = tk.StringVar()
        self.total_elements_label = ttk.Label(root, textvariable=self.total_elements_var)
        self.total_elements_label.pack(pady=5)

        self.df = None
        self.selected_sheet = None
        self.selected_column = None

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                self.df = pd.read_excel(file_path, sheet_name=None)
                self.update_sheets_listbox()
                self.update_columns_listbox()
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def update_sheets_listbox(self):
        self.sheets_listbox.delete(0, tk.END)
        for sheet_name in self.df.keys():
            self.sheets_listbox.insert(tk.END, sheet_name)

    def update_columns_listbox(self):
        self.columns_listbox.delete(0, tk.END)
        if self.selected_sheet:
            columns = self.df[self.selected_sheet].columns
            for column in columns:
                self.columns_listbox.insert(tk.END, column)

    def select_sheet(self, event):
        selected_index = self.sheets_listbox.curselection()
        if selected_index:
            self.selected_sheet = self.sheets_listbox.get(selected_index[0])
            self.update_columns_listbox()
            self.update_total_elements()

    def select_column(self, event):
        selected_index = self.columns_listbox.curselection()
        if selected_index:
            self.selected_column = self.columns_listbox.get(selected_index[0])
            unique_elements = self.df[self.selected_sheet][self.selected_column].unique()
            self.unique_elements_text.delete(1.0, tk.END)
            for element in unique_elements:
                self.unique_elements_text.insert(tk.END, f"{element}\n")
            self.update_total_elements()

    def update_total_elements(self):
        if self.selected_column:
            total_elements = len(self.df[self.selected_sheet][self.selected_column])
            self.total_elements_var.set(total_elements)

    def export_elements(self):
        if self.selected_column:
            elements = self.df[self.selected_sheet][self.selected_column].unique()
            export_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if export_path:
                export_df = pd.DataFrame(elements, columns=[self.selected_column])
                export_df.to_excel(export_path, index=False)
                messagebox.showinfo("Export Successful", "Unique elements exported to Excel.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
