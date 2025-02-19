import pandas as pd
import requests
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime, timedelta
import tkinterdnd2 as tkdnd  # 需要先安裝：pip install tkinterdnd2

class DragDropApp(tkdnd.Tk):
    def __init__(self):
        super().__init__()
        self.title("PubMed Disease-Related Article Query Tool")
        self.current_file = None
        self.df = None  # 新增 DataFrame 變數
        self.setup_ui()
        
    def fetch_pmids_with_term(self, gene, term, date_ranges):
        try:
            url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
            results = {}
            current_year = datetime.now().year

            # Initialize results dictionary for each date range
            for range_name in date_ranges:
                results[range_name] = []

            for range_name in date_ranges:
                pmids = []
                retmax = 100
                retstart = 0

                # Calculate date range
                date_filter = ""
                if range_name != "no_limit":  # Skip date filter for "no limit" option
                    if range_name == "10_years":
                        start_date = f"{current_year-10}"
                    elif range_name == "5_years":
                        start_date = f"{current_year-5}"
                    elif range_name == "3_years":
                        start_date = f"{current_year-3}"
                    date_filter = f" AND {start_date}:3000[PDAT]"

                while True:
                    params = {
                        "db": "pubmed",
                        "term": f"({term}) AND ({gene}){date_filter}",
                        "retmode": "xml",
                        "retmax": retmax,
                        "retstart": retstart
                    }
                    response = requests.get(url, params=params)
                    if response.status_code == 200:
                        root = ET.fromstring(response.text)
                        batch_pmids = [id_elem.text for id_elem in root.findall(".//Id")]
                        pmids.extend(batch_pmids)
                        if len(batch_pmids) < retmax:
                            break
                        retstart += retmax
                    else:
                        break
                results[range_name] = pmids

            return results
        except Exception as e:
            return {range_name: [] for range_name in date_ranges}

    def setup_ui(self):
        # Drop zone frame
        self.drop_frame = tk.LabelFrame(self, text="Drop Excel File Here", padx=10, pady=10)
        self.drop_frame.pack(pady=10, padx=10, fill="x")
        
        # File display label
        self.file_label = tk.Label(self.drop_frame, text="No file selected", wraplength=350)
        self.file_label.pack(pady=20)
        
        # Configure drop zone
        self.drop_frame.drop_target_register(tkdnd.DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.handle_drop)
        
        # Browse button
        self.browse_button = tk.Button(self.drop_frame, text="Browse File", command=self.browse_file)
        self.browse_button.pack(pady=5)

        # Title
        title_label = tk.Label(self, text="PubMed Disease-Related Article Query Tool", font=("Helvetica", 16))
        title_label.pack(pady=10)

        # Time range selection frame
        time_range_frame = tk.LabelFrame(self, text="Select Time Ranges", padx=10, pady=5)
        time_range_frame.pack(pady=5, padx=10, fill="x")

        self.no_limit_var = tk.BooleanVar()
        self.year_10_var = tk.BooleanVar()
        self.year_5_var = tk.BooleanVar()
        self.year_3_var = tk.BooleanVar()

        tk.Checkbutton(time_range_frame, text="No Time Limit", variable=self.no_limit_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(time_range_frame, text="Past 10 Years", variable=self.year_10_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(time_range_frame, text="Past 5 Years", variable=self.year_5_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(time_range_frame, text="Past 3 Years", variable=self.year_3_var).pack(side=tk.LEFT, padx=5)

        # Disease condition input box
        diseases_label = tk.Label(self, text="Enter Disease Conditions (Format: Disease: Alias1, Alias2, ...)")
        diseases_label.pack(pady=5)
        self.diseases_text = tk.Text(self, height=10, width=50)
        self.diseases_text.pack(pady=5)

        # Test mode options
        test_mode_frame = tk.Frame(self)
        test_mode_frame.pack(pady=10)
        self.test_mode_var = tk.BooleanVar()
        test_mode_checkbox = tk.Checkbutton(test_mode_frame, text="Enable Test Mode", variable=self.test_mode_var)
        test_mode_checkbox.pack(side=tk.LEFT)
        test_limit_label = tk.Label(test_mode_frame, text="Number of Rows to Test:")
        test_limit_label.pack(side=tk.LEFT, padx=5)
        self.test_limit_entry = tk.Entry(test_mode_frame, width=5)
        self.test_limit_entry.insert(0, "10")
        self.test_limit_entry.pack(side=tk.LEFT)

        # Progress bar
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=10)
        self.progress_label = tk.Label(self, text="Progress: 0%")
        self.progress_label.pack()

        # Progress information box
        self.progress_info_text = tk.Text(self, height=15, width=60, state="normal")
        self.progress_info_text.pack(pady=10)

        # Button Frame
        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)
        
        # Run button
        self.run_button = tk.Button(button_frame, text="Start Query", 
                                  command=self.run_program, 
                                  bg="green", fg="white")
        self.run_button.pack(side=tk.LEFT, padx=5)
        
        # Save button (initially disabled)
        self.save_button = tk.Button(button_frame, text="Save Results", 
                                   command=self.save_results,
                                   bg="blue", fg="white",
                                   state="disabled")
        self.save_button.pack(side=tk.LEFT, padx=5)

    def handle_drop(self, event):
        file_path = event.data
        # 處理 Windows 路徑格式 {filepath}
        if file_path.startswith("{") and file_path.endswith("}"):
            file_path = file_path[1:-1]
        
        if file_path.lower().endswith('.xlsx'):
            self.current_file = file_path
            self.file_label.config(text=f"Selected file: {file_path}")
        else:
            messagebox.showerror("Error", "Please drop an Excel file (.xlsx)")

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.current_file = file_path
            self.file_label.config(text=f"Selected file: {file_path}")

    def update_progress_info(self, message):
        self.progress_info_text.insert(tk.END, message + "\n")
        self.progress_info_text.see(tk.END)
        self.update()

    def run_program(self):
        if not self.current_file:
            messagebox.showerror("Error", "Please select or drop an Excel file first!")
            return

        # Get selected date ranges
        selected_ranges = []
        if self.no_limit_var.get():
            selected_ranges.append("no_limit")
        if self.year_10_var.get():
            selected_ranges.append("10_years")
        if self.year_5_var.get():
            selected_ranges.append("5_years")
        if self.year_3_var.get():
            selected_ranges.append("3_years")

        if not selected_ranges:
            messagebox.showerror("Error", "Please select at least one time range!")
            return

        # Read gene list
        try:
            self.df = pd.read_excel(self.current_file)
        except Exception as e:
            messagebox.showerror("Error", f"Unable to read Gene List File: {e}")
            return

        # Check for 'Gene' or 'gene' column
        if "Gene" in self.df.columns:
            gene_column = "Gene"
        elif "gene" in self.df.columns:
            gene_column = "gene"
        else:
            messagebox.showerror("Error", "The file must contain a 'Gene' or 'gene' column!")
            return

        # Enable test mode
        if self.test_mode_var.get():
            test_limit = int(self.test_limit_entry.get())
            self.df = self.df.head(test_limit)
            messagebox.showinfo("Test Mode", f"Test Mode Enabled: Only processing the first {test_limit} genes.")

        # Get disease conditions
        diseases_input = self.diseases_text.get("1.0", tk.END).strip()
        if not diseases_input:
            messagebox.showerror("Error", "Please enter disease conditions!")
            return

        # Parse disease conditions
        try:
            diseases = {}
            for line in diseases_input.split("\n"):
                if line.strip():
                    disease, aliases = line.split(":")
                    diseases[disease.strip()] = [alias.strip() for alias in aliases.split(",")]
        except Exception as e:
            messagebox.showerror("Error", f"Invalid format for disease conditions: {e}")
            return

        # Prepare columns for each time range
        for disease in diseases.keys():
            for range_name in selected_ranges:
                if range_name == "no_limit":
                    column_suffix = "all"
                else:
                    column_suffix = f"{range_name.split('_')[0]}yr"
                self.df[f"{disease}_PMIDs_{column_suffix}"] = ""
                self.df[f"{disease}_Counts_{column_suffix}"] = ""

        # Configure progress bar
        self.progress_bar["maximum"] = len(self.df)
        self.progress_bar["value"] = 0
        self.progress_label.config(text="Progress: 0%")
        self.progress_info_text.delete("1.0", tk.END)
        self.update()

        # Query PubMed for each gene
        for idx, gene in enumerate(self.df[gene_column]):
            self.update_progress_info(f"Processing Gene: {gene} ({idx + 1}/{len(self.df)})")
            
            for disease, aliases in diseases.items():
                query = " OR ".join([f"({alias})" for alias in aliases])
                self.update_progress_info(f"  - Querying Disease: {disease}, Aliases: {', '.join(aliases)}")
                
                # Get results for all selected date ranges
                results = self.fetch_pmids_with_term(gene, query, selected_ranges)
                
                # Update dataframe with results for each time range
                for range_name in selected_ranges:
                    if range_name == "no_limit":
                        column_suffix = "all"
                        year_text = "all years"
                    else:
                        year_text = range_name.split("_")[0]
                        column_suffix = f"{year_text}yr"
                    
                    pmids = results[range_name]
                    self.df.at[idx, f"{disease}_PMIDs_{column_suffix}"] = ", ".join(pmids[:5]) if pmids else "No PMIDs Found"
                    self.df.at[idx, f"{disease}_Counts_{column_suffix}"] = len(pmids)
                    self.update_progress_info(f"    -> Found {len(pmids)} articles for {year_text}")

            # Update progress
            self.progress_bar["value"] = idx + 1
            self.progress_label.config(text=f"Progress: {int((idx + 1) / len(self.df) * 100)}%")
            self.update()

        # Enable save button after query is complete
        self.save_button.config(state="normal")
        messagebox.showinfo("Completed", "Query completed! You can now save the results.")

    def save_results(self):
        # Output file name
        default_filename = "Gene_PMIDs.xlsx"
        output_file = filedialog.asksaveasfilename(
            title="Save Output File As",
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not output_file:
            return

        # Save results
        try:
            self.df.to_excel(output_file, index=False)
            messagebox.showinfo("Saved", f"Results saved to: {output_file}")
            # Disable save button after successful save
            self.save_button.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Unable to save results: {e}")

# 主程式啟動
if __name__ == "__main__":
    app = DragDropApp()
    app.mainloop()
