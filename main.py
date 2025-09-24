# === ver1.4 === #
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import os
import shutil
from fillpdf import fillpdfs
from pdf2image import convert_from_path
from PyPDF2 import PdfMerger   # <--- added for merging

# === Paths ===
TEMPLATE_PDF = "template.pdf"
TEMPLATE_CSV = os.path.join("assets", "BIR_Invoice.csv")
PREVIEW_IMAGE = os.path.join("assets", "blank_invoice.png")
POPLER_PATH = r"C:\poppler-24.07.0\Library\bin"

# === Helpers ===
def split_tin(tin_value):
    """Split a TIN string into 4 parts or return blanks if empty."""
    if pd.isna(tin_value) or str(tin_value).strip() in ("", "-"):
        return ["", "", "", ""]
    parts = str(tin_value).split("-")
    while len(parts) < 4:
        parts.append("")
    return parts[:4]

def format_date_mmddyyyy(date_value):
    """Convert 8/1/2025 to 08012025 (keep blank if missing)"""
    if pd.isna(date_value) or str(date_value).strip() == "":
        return ""
    try:
        d = pd.to_datetime(date_value)
        return d.strftime("%m%d%Y")
    except:
        return str(date_value)

def flatten_pdf_to_jpg_pdf(input_pdf, output_pdf):
    """Render PDF to images then re-save as flattened PDF."""
    images = convert_from_path(input_pdf, dpi=300, poppler_path=POPLER_PATH)
    images[0].save(output_pdf, "PDF", resolution=300.0,
                   save_all=True, append_images=images[1:])

# === Main Process Function ===
def process_csv(csv_path):
    try:
        df = pd.read_csv(csv_path, dtype=str)
        merged_files = []  # collect flattened pdfs here

        for index, row in df.iterrows():
            From = format_date_mmddyyyy(row.get('From', ""))
            To = format_date_mmddyyyy(row.get('To', ""))
            TIN1, TIN2, TIN3, TIN4 = split_tin(row.get('TIN', ""))
            myTIN1, myTIN2, myTIN3, myTIN4 = split_tin(row.get('myTIN', ""))

            def safe(v): return "" if pd.isna(v) else str(v)

            Payee = safe(row.get('Payee', ""))
            Address = safe(row.get('Address', ""))
            Zip_Code = safe(row.get('Zip_Code', ""))
            myPayee = safe(row.get('myPayee', ""))
            myAddress = safe(row.get('myAddress', ""))
            myZip_Code = safe(row.get('myZip_Code', ""))
            IPS_EWT = safe(row.get('IPS_EWT', ""))
            ATC = safe(row.get('ATC', ""))
            M1 = safe(row.get('M1', ""))
            M2 = safe(row.get('M2', ""))
            M3 = safe(row.get('M3', ""))
            M_Total = safe(row.get('M_Total', ""))
            TWQ = safe(row.get('TWQ', ""))
            max_total = safe(row.get('max_total', ""))

            form_fields = list(fillpdfs.get_form_fields(TEMPLATE_PDF).keys())
            data_dict = {
                form_fields[0]: From,
                form_fields[1]: To,
                form_fields[2]: TIN1,
                form_fields[3]: TIN2,
                form_fields[4]: TIN3,
                form_fields[5]: TIN4,
                form_fields[6]: Payee,
                form_fields[7]: Address,
                form_fields[8]: Zip_Code,
                form_fields[9]: myTIN1,
                form_fields[10]: myTIN2,
                form_fields[11]: myTIN3,
                form_fields[12]: myTIN4,
                form_fields[13]: myPayee,
                form_fields[14]: myAddress,
                form_fields[15]: myZip_Code,
                form_fields[16]: IPS_EWT,
                form_fields[17]: ATC,
                form_fields[18]: M1,
                form_fields[19]: M2,
                form_fields[20]: M3,
                form_fields[21]: M_Total,
                form_fields[22]: TWQ,
                form_fields[23]: max_total
            }

            # Step 1: fill PDF
            filled_pdf_path = f"filled_{index+1}.pdf"
            fillpdfs.write_fillable_pdf(TEMPLATE_PDF, filled_pdf_path, data_dict)

            # Step 2: flatten PDF (image to PDF)
            flattened_pdf_path = f"flattened_{index+1}.pdf"
            flatten_pdf_to_jpg_pdf(filled_pdf_path, flattened_pdf_path)

            merged_files.append(flattened_pdf_path)  # add to list
            print(f"✅ Done row {index+1} → {flattened_pdf_path}")

        # Step 3: merge all flattened pdfs
        if merged_files:
            merger = PdfMerger()
            for pdf_file in merged_files:
                merger.append(pdf_file)
            merged_output = "merged_output.pdf"
            merger.write(merged_output)
            merger.close()
            print(f"✅ Merged PDF created: {merged_output}")

        status_var.set("✅ PDFs generated, flattened & merged successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        status_var.set("❌ Failed to generate PDFs.")

# === File Selectors ===
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        csv_path_var.set(file_path)

def run_process():
    path = csv_path_var.get()
    if not path:
        messagebox.showwarning("Missing File", "Please select a CSV file.")
        return
    process_csv(path)

def download_template():
    dest_path = filedialog.asksaveasfilename(
        title="Save CSV Template As",
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv")]
    )
    if dest_path:
        try:
            shutil.copyfile(TEMPLATE_CSV, dest_path)
            messagebox.showinfo("Saved", "CSV template saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save template:\n{e}")

# === Build UI ===
root = tk.Tk()
root.title("PDF Filler + Flattener App ni Gemson")
root.geometry("850x600")
root.resizable(False, False)

# === Preview Image ===
try:
    img = Image.open(PREVIEW_IMAGE)
    img = img.resize((400, 550))
    preview_img = ImageTk.PhotoImage(img)
    img_label = tk.Label(root, image=preview_img)
    img_label.image = preview_img
    img_label.pack(side=tk.LEFT, padx=10, pady=10)
except Exception as e:
    tk.Label(root, text="Preview not found").pack(side=tk.LEFT, padx=10, pady=10)

# === Right-side Frame ===
frame = tk.Frame(root)
frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=20, pady=20)

csv_path_var = tk.StringVar()
status_var = tk.StringVar()

tk.Label(frame, text="CSV File:").pack(anchor="w")
tk.Entry(frame, textvariable=csv_path_var, width=50).pack(pady=5)
tk.Button(frame, text="📂 Browse CSV", command=browse_file, width=30).pack(pady=5)
tk.Button(frame, text="✅ Generate & Flatten PDFs", command=run_process, width=30, bg="green", fg="white").pack(pady=10)
tk.Button(frame, text="⬇️ Download CSV Template", command=download_template, width=30).pack(pady=5)
tk.Label(frame, textvariable=status_var, fg="blue").pack(pady=20)

# === Run ===
root.mainloop()
