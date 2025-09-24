import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import os
import shutil
from fillpdf import fillpdfs
from PyPDF2 import PdfReader, PdfWriter  # <— for flattening

# === Paths ===
TEMPLATE_PDF = "template.pdf"
TEMPLATE_CSV = os.path.join("assets", "BIR_Invoice.csv")
PREVIEW_IMAGE = os.path.join("assets", "blank_invoice.png")

# === Helpers ===
def split_tin(tin_value):
    """Split a TIN string into 4 parts or return blanks if empty."""
    if pd.isna(tin_value) or str(tin_value).strip() == "" or str(tin_value).strip() == "-":
        return ["", "", "", ""]
    parts = str(tin_value).split("-")
    while len(parts) < 4:
        parts.append("")
    return parts[:4]

def flatten_pdf(input_path, output_path):
    """Flatten AcroForm fields (remove fillable fields)."""
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    # remove AcroForm so fields become static
    if "/AcroForm" in reader.trailer["/Root"]:
        del reader.trailer["/Root"]["/AcroForm"]
    with open(output_path, "wb") as f:
        writer.write(f)

# === Main Process Function ===
def process_csv(csv_path):
    try:
        df = pd.read_csv(csv_path, dtype=str)

        for index, row in df.iterrows():
            From = str(row['From']).replace("-", "") if pd.notna(row['From']) else ""
            To = str(row['To']).replace("-", "") if pd.notna(row['To']) else ""
            TIN1, TIN2, TIN3, TIN4 = split_tin(row['TIN'])
            myTIN1, myTIN2, myTIN3, myTIN4 = split_tin(row['myTIN'])

            Payee = row.get('Payee', "")
            Address = row.get('Address', "")
            Zip_Code = row.get('Zip_Code', "")
            myPayee = row.get('myPayee', "")
            myAddress = row.get('myAddress', "")
            myZip_Code = row.get('myZip_Code', "")
            IPS_EWT = row.get('IPS_EWT', "")
            ATC = row.get('ATC', "")
            M1 = row.get('M1', "")
            M2 = row.get('M2', "")
            M3 = row.get('M3', "")
            M_Total = row.get('M_Total', "")
            TWQ = row.get('TWQ', "")
            max_total = row.get('max_total', "")

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

            temp_output = f"temp_filled_{index + 1}.pdf"
            final_output = f"output_filled_{index + 1}.pdf"

            # fill form first into temp
            fillpdfs.write_fillable_pdf(TEMPLATE_PDF, temp_output, data_dict)

            # flatten into final
            flatten_pdf(temp_output, final_output)

            # delete temp file
            os.remove(temp_output)

        status_var.set("✅ Flattened PDFs generated successfully!")

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
root.title("PDF Filler App ni Gemson")
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
tk.Button(frame, text="✅ Generate PDFs", command=run_process, width=30, bg="green", fg="white").pack(pady=10)
tk.Button(frame, text="⬇️ Download CSV Template", command=download_template, width=30).pack(pady=5)
tk.Label(frame, textvariable=status_var, fg="blue").pack(pady=20)

# === Run ===
root.mainloop()
