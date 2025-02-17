import os
import pandas as pd
import requests
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Define PubMed query function
def fetch_pmids_with_term(gene, term):
    try:
        url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
        pmids = []
        retmax = 100
        retstart = 0

        while True:
            params = {
                "db": "pubmed",
                "term": f"({term}) AND ({gene})",
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
        return pmids
    except Exception as e:
        return []

# Update progress information
def update_progress_info(message):
    progress_info_text.insert(tk.END, message + "\n")
    progress_info_text.see(tk.END)  # Auto scroll to the last line
    root.update()

# Main program logic
def run_program():
    # Load gene list file
    input_file = filedialog.askopenfilename(
        title="Select Gene List File", 
        filetypes=[("Excel Files", "*.xlsx")],
        initialdir=os.getcwd()  # Set the default directory to the current working directory
    )
    if not input_file:
        messagebox.showerror("Error", "No Gene List File Selected!")
        return

    # Output file name
    output_file = filedialog.asksaveasfilename(
        title="Save Output File As", 
        defaultextension=".xlsx", 
        filetypes=[("Excel Files", "*.xlsx")],
        initialdir=os.getcwd()  # Set the default directory to the current working directory
    )
    if not output_file:
        messagebox.showerror("Error", "No Output File Selected!")
        return

    # Read gene list
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        messagebox.showerror("Error", f"Unable to read Gene List File: {e}")
        return

    # Check for 'Gene' or 'gene' column
    if "Gene" in df.columns:
        gene_column = "Gene"
    elif "gene" in df.columns:
        gene_column = "gene"
    else:
        messagebox.showerror("Error", "The file must contain a 'Gene' or 'gene' column!")
        return

    # Enable test mode to process only the first N rows
    if test_mode_var.get():
        test_limit = int(test_limit_entry.get())
        df = df.head(test_limit)
        messagebox.showinfo("Test Mode", f"Test Mode Enabled: Only processing the first {test_limit} genes.")

    # Get user input for disease conditions
    diseases_input = diseases_text.get("1.0", tk.END).strip()
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

    # Prepare columns
    for disease in diseases.keys():
        df[f"{disease}_PMIDs"] = ""
        df[f"{disease}_Counts"] = ""

    # Configure progress bar
    progress_bar["maximum"] = len(df)
    progress_bar["value"] = 0
    progress_label.config(text="Progress: 0%")
    progress_info_text.delete("1.0", tk.END)  # Clear progress info
    root.update()

    # Query PubMed for each gene
    for idx, gene in enumerate(df[gene_column]):
        update_progress_info(f"Processing Gene: {gene} ({idx + 1}/{len(df)})")
        for disease, aliases in diseases.items():
            query = " OR ".join([f"({alias})" for alias in aliases])
            update_progress_info(f"  - Querying Disease: {disease}, Aliases: {', '.join(aliases)}")
            pmids = fetch_pmids_with_term(gene, query)
            df.at[idx, f"{disease}_PMIDs"] = ", ".join(pmids[:5]) if pmids else "No PMIDs Found"
            df.at[idx, f"{disease}_Counts"] = len(pmids)
            update_progress_info(f"    -> Found {len(pmids)} related articles")

        # Update progress bar
        progress_bar["value"] = idx + 1
        progress_label.config(text=f"Progress: {int((idx + 1) / len(df) * 100)}%")
        root.update()

    # Save results
    try:
        df.to_excel(output_file, index=False)
        messagebox.showinfo("Completed", f"Results saved to: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Unable to save results: {e}")

# Create GUI
root = tk.Tk()
root.title("PubMed Disease-Related Article Query Tool")

# Title
title_label = tk.Label(root, text="PubMed Disease-Related Article Query Tool", font=("Helvetica", 16))
title_label.pack(pady=10)

# Disease condition input box
diseases_label = tk.Label(root, text="Enter Disease Conditions (Format: Disease: Alias1, Alias2, ...)")
diseases_label.pack(pady=5)
diseases_text = tk.Text(root, height=10, width=50)
diseases_text.pack(pady=5)

# Test mode options
test_mode_frame = tk.Frame(root)
test_mode_frame.pack(pady=10)
test_mode_var = tk.BooleanVar()
test_mode_checkbox = tk.Checkbutton(test_mode_frame, text="Enable Test Mode", variable=test_mode_var)
test_mode_checkbox.pack(side=tk.LEFT)
test_limit_label = tk.Label(test_mode_frame, text="Number of Rows to Test:")
test_limit_label.pack(side=tk.LEFT, padx=5)
test_limit_entry = tk.Entry(test_mode_frame, width=5)
test_limit_entry.insert(0, "10")  # Default test limit is 10
test_limit_entry.pack(side=tk.LEFT)

# Progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=10)
progress_label = tk.Label(root, text="Progress: 0%")
progress_label.pack()

# Progress information box
progress_info_text = tk.Text(root, height=15, width=60, state="normal")
progress_info_text.pack(pady=10)

# Run button
run_button = tk.Button(root, text="Start Query", command=run_program, bg="green", fg="white")
run_button.pack(pady=20)

# Start GUI
root.mainloop()