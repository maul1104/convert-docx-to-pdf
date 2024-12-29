import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert

def convert_docx_to_pdf(filepath, output_dir):
    """
    Fungsi untuk mengonversi DOCX ke PDF menggunakan docx2pdf
    """
    try:
        # Konversi file DOCX ke PDF
        convert(filepath, output_dir)
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False

def select_file():
    """
    Fungsi untuk memilih file DOCX dan menyimpan sebagai PDF.
    """
    filepath = filedialog.askopenfilename(
        filetypes=[("DOCX files", "*.docx")],
        title="Pilih File DOCX"
    )
    if filepath:
        output_dir = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Simpan File PDF"
        )
        if output_dir:
            success = convert_docx_to_pdf(filepath, output_dir)
            if success:
                messagebox.showinfo("Berhasil", f"File berhasil dikonversi ke: {output_dir}")
                os.startfile(output_dir)  # Buka file PDF
            else:
                messagebox.showerror("Gagal", "Konversi gagal.")

def main():
    """
    Fungsi utama untuk antarmuka pengguna.
    """
    root = tk.Tk()
    root.title("Konversi DOCX ke PDF")
    root.geometry("400x200")
    
    instruction_label = tk.Label(
        root, text="Klik tombol di bawah untuk memilih file DOCX dan konversi ke PDF.",
        wraplength=350, justify="center", font=("Arial", 12)
    )
    instruction_label.pack(pady=10)

    select_button = tk.Button(
        root, text="Pilih File DOCX", command=select_file,
        font=("Arial", 12), bg="#4CAF50", fg="white", padx=10, pady=5
    )
    select_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
