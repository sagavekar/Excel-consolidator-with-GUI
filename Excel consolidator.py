import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

class ExcelConsolidator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('Excel / CSV Consolidator')
        self.root.geometry("300x300")

        # Input file selection button
        self.input_files_button = tk.Button(
            self.root,
            text='Select Input Files',
            command=self.select_input_files,bg='#ffffff', activebackground='#00ff00'
        )
        self.input_files_button.pack(padx=10, pady=10)

        
        # Output file selection button
        self.output_file_button = tk.Button(
            self.root,
            text='Select Output File',
            command=self.select_output_file,
            bg='#ffffff', activebackground='#00ff00'
        )
        self.output_file_button.pack(padx=10, pady=10)

        # Sheet number input box
        self.sheet_number_label = tk.Label(
            self.root,
            text='Sheet Number:'
        )
        self.sheet_number_label.pack(padx=10, pady=10)

        self.sheet_number_entry = tk.Entry(
            self.root
        )
        self.sheet_number_entry.pack(padx=10, pady=10)

        # Consolidate button
        self.consolidate_button = tk.Button(
            self.root,
            text='Consolidate',
            command=self.consolidate,bg='#ffffff', activebackground='#00ff00'
        )
        self.consolidate_button.pack(padx=10, pady=10)
        
        label3 = tk.Label(self.root, text="Designed and developed by Omkar Sagavekar", font="Tahoma 10", bg ="black",fg="white")
        label3.pack(anchor="s",fill="x",side = "bottom")


        
        self.input_file_paths = []
        self.output_file_path = None
        self.consolidated_data = None

        self.root.mainloop()

    def select_input_files(self):
        self.input_file_paths = filedialog.askopenfilenames(
            title='Select Input Files',
            filetypes=[('Excel Files', '*.xlsx;*.xls'), ('CSV Files', '*.csv')]
        )
        if self.input_file_paths:
            self.input_files_button.configure(text='Selected {} input file(s)'.format(len(self.input_file_paths)))


    def select_output_file(self):
        self.output_file_path = filedialog.asksaveasfilename(
            title='Save Consolidated Data',
            defaultextension='.xlsx',
            filetypes=[('Excel Files', '*.xlsx')]
        )
        if self.output_file_path:
            self.output_file_button.configure(text='Selected output file: {}'.format(self.output_file_path))

    def consolidate(self):
        if not self.input_file_paths:
            messagebox.showerror('Error', 'Please select input files')
            return

        if not self.output_file_path:
            messagebox.showerror('Error', 'Please select output file')
            return

        try:
            sheet_number = int(self.sheet_number_entry.get())
        except ValueError:
            messagebox.showerror('Error', 'Invalid sheet number')
            return

        data_list = []
        for file_path in self.input_file_paths:
            try:
                if file_path.endswith('.csv'):
                    data = pd.read_csv(file_path).assign(source_file=file_path)
                else:
                    data = pd.read_excel(file_path, sheet_name=sheet_number-1).assign(source_file=file_path)
                data_list.append(data)
            except Exception as e:
                messagebox.showwarning('Warning', 'Error reading file {}: {}'.format(file_path, e))
                continue

        if not data_list:
            messagebox.showwarning('Warning', 'No valid input files selected')
            return

        self.consolidated_data = pd.concat(data_list)

        try:
            self.consolidated_data.to_excel(self.output_file_path, index=False)
        except Exception as e:
            messagebox.showerror('Error', 'Failed to save consolidated data to output file')
            return

        messagebox.showinfo('Success', f'Total {len(data_list)} files consolidated!')

if __name__ == '__main__':
    app = ExcelConsolidator()   