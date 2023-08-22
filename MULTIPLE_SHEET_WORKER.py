import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import threading
import xlsxwriter

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EXCEL COMMON ELEMENTS v2.1 by Anindya Karmaker")

        self.load_button = ttk.Button(root, text="Load Excel File", command=self.load_excel)
        self.load_button.pack(pady=10)
                
        self.unique_elements_label0 = ttk.Label(root, text="Select the Sheet:")
        self.unique_elements_label0.pack(pady=5)

        self.sheets_listbox = tk.Listbox(root)
        self.sheets_listbox.pack(pady=10, fill=tk.BOTH)
        
        self.sheets_listbox.bind("<<ListboxSelect>>", self.select_sheet)
        
        
                        
        self.column_list_frame = ttk.Frame(root)
        self.column_list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.columns_list_frame = ttk.LabelFrame(self.column_list_frame, text="Available Columns")
        self.columns_list_frame.pack(side=tk.LEFT, padx=5, fill=tk.BOTH, expand=True)

        self.columns_listbox = tk.Listbox(self.columns_list_frame)
        self.columns_listbox.pack(pady=5, fill=tk.BOTH, expand=True)

        self.columns_label_frame = ttk.LabelFrame(self.column_list_frame, text="Selected Columns")
        self.columns_label_frame.pack(side=tk.LEFT, padx=5, fill=tk.BOTH, expand=True)


        self.selected_columns_listbox = tk.Listbox(self.columns_label_frame)
        self.selected_columns_listbox.pack(pady=5, fill=tk.BOTH, expand=True)

        
        
        self.column_buttons_frame = ttk.Frame(root)
        self.column_buttons_frame.pack(pady=5)

        self.add_columns_button = ttk.Button(self.column_buttons_frame, text="Select Column", command=self.add_columns_to_list)
        self.add_columns_button.pack(side=tk.LEFT, padx=5)

        self.remove_columns_button = ttk.Button(self.column_buttons_frame, text="Remove Selected Column", command=self.remove_selected_column)
        self.remove_columns_button.pack(side=tk.LEFT, padx=5)
        self.selected_sheets = {}
        self.selected_columns = []
        
        self.unique_export_frame = ttk.Frame(root)
        self.unique_export_frame.pack(pady=5)

        self.get_unique_elements_button = ttk.Button(self.unique_export_frame, text="Get Common Unique Elements", command=self.get_common_unique_elements)
        self.get_unique_elements_button.pack(side=tk.LEFT, padx=5)

        self.export_button = ttk.Button(self.unique_export_frame, text="Export Common Unique Elements", command=self.export_elements)
        self.export_button.pack(side=tk.LEFT, padx=5)
        
        self.unique_elements_label = ttk.Label(root, text="Unique Elements:")
        self.unique_elements_label.pack(pady=5)

        self.unique_elements_text = tk.Text(root, height=10, width=40)
        self.unique_elements_text.pack(pady=5)

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
            self.progress_window = tk.Toplevel(self.root)
            self.progress_window.title("Loading Excel File")
            self.progress_window.geometry("400x50")  # Set the size of the progress window

            self.progress_bar = ttk.Progressbar(self.progress_window, mode="indeterminate")
            self.progress_bar.pack(pady=10)

            self.load_thread = threading.Thread(target=self.load_excel_thread, args=(file_path,))
            self.load_thread.start()

            self.progress_window.protocol("WM_DELETE_WINDOW", self.on_progress_window_close)

    def on_progress_window_close(self):
        if self.load_thread.is_alive():
            self.load_thread.join()
        self.progress_window.destroy()

    def load_excel_thread(self, file_path):
        try:
            self.df = pd.read_excel(file_path, sheet_name=None)
            self.update_sheets_listbox()
            self.update_columns_listbox()
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.progress_bar.stop()
            self.progress_bar.destroy()
            self.progress_window.destroy()

  
    def select_sheet(self, event):
        selected_index = self.sheets_listbox.curselection()
        if selected_index:
            self.selected_sheet = self.sheets_listbox.get(selected_index[0])
            self.update_columns_listbox()
            self.update_total_elements()
            
    def get_unique_elements(self):
        self.unique_elements_text.delete(1.0, tk.END)
        for column_name in self.selected_columns:
            unique_elements = self.df[self.selected_sheet][column_name].unique()
            self.unique_elements_text.insert(tk.END, f"Unique elements in {column_name}:\n")
            for element in unique_elements:
                self.unique_elements_text.insert(tk.END, f"{element}\n")
            self.unique_elements_text.insert(tk.END, "\n")

        self.update_total_elements()
    
    def get_common_unique_elements(self):
        self.unique_elements_text.delete(1.0, tk.END)
        common_unique_elements = None

        for sheet, columns in self.selected_sheets.items():
            for column_name in columns:
                unique_elements = set(self.df[sheet][column_name].unique())
                if common_unique_elements is None:
                    common_unique_elements = unique_elements
                else:
                    common_unique_elements &= unique_elements  # Intersection of sets

        if common_unique_elements:
            self.unique_elements_text.insert(tk.END, "Common Unique Elements:\n")
            for element in common_unique_elements:
                self.unique_elements_text.insert(tk.END, f"{element}\n")
            self.total_elements_var.set(len(common_unique_elements))
        else:
            self.unique_elements_text.insert(tk.END, "No common unique elements found.\n")
            self.total_elements_var.set(0)    
    
    def add_columns_to_list(self):
        selected_indices = self.columns_listbox.curselection()
        if selected_indices and self.selected_sheet:
            if self.selected_sheet not in self.selected_sheets:
                self.selected_sheets[self.selected_sheet] = []
            for index in selected_indices:
                column_name = self.columns_listbox.get(index)
                if column_name not in self.selected_sheets[self.selected_sheet]:
                    self.selected_sheets[self.selected_sheet].append(column_name)
                    self.selected_columns_listbox.insert(tk.END, column_name)
    
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
                
    def remove_selected_column(self):
        selected_indices = self.selected_columns_listbox.curselection()
        if selected_indices and self.selected_sheet:
            for index in selected_indices:
                column_name = self.selected_columns_listbox.get(index)
                if column_name in self.selected_sheets.get(self.selected_sheet, []):
                    self.selected_sheets[self.selected_sheet].remove(column_name)
                    self.selected_columns_listbox.delete(index)
                else:
                    messagebox.showerror("Error", "Column not found in selected columns.")
        else:
            messagebox.showerror("Error", "No column selected.")
    
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
        if self.selected_sheets:
            export_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if export_path:
                writer = pd.ExcelWriter(export_path, engine='xlsxwriter')
    
                for sheet, columns in self.selected_sheets.items():
                    data = []
                    for column_name in columns:
                        unique_elements = self.df[sheet][column_name].dropna().unique()  # Drop rows with missing values
                        data.extend([(f"{sheet} - {column_name}", element) for element in unique_elements])
                    
                    df_export = pd.DataFrame(data, columns=['Sheet-Column', 'Element'])
                    df_export.to_excel(writer, sheet_name=sheet, index=False)
                
                writer.save()
                writer.close()
                messagebox.showinfo("Export Successful", "Common unique elements exported to an Excel file.")
        else:
            messagebox.showerror("Error", "No columns selected.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
