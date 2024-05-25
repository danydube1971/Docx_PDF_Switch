import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2docx import Converter

def convert_docx_to_pdf(docx_path, libreoffice_path="/usr/local/bin/libreoffice24.2"):
    try:
        cmd = [libreoffice_path, '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(docx_path), docx_path]
        subprocess.run(cmd, check=True)
        pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
        print(f"Conversion réussie : {docx_path} -> {pdf_path}")
        return True
    except Exception as e:
        print(f"Erreur lors de la conversion : {e}")
        messagebox.showerror("Erreur", f"Une erreur est survenue lors de la conversion : {e}")
        return False

def convert_pdf_to_docx(pdf_path):
    try:
        output = os.path.splitext(pdf_path)[0] + '.docx'
        cv = Converter(pdf_path)
        cv.convert(output, start=0, end=None)
        cv.close()
        print(f"Conversion réussie : {pdf_path} -> {output}")
        return True
    except Exception as e:
        print(f"Erreur lors de la conversion : {e}")
        messagebox.showerror("Erreur", f"Une erreur est survenue lors de la conversion : {e}")
        return False

def convert_all_docx_to_pdf(folder_path, progress_bar, progress_label, root):
    libreoffice_path = "/usr/local/bin/libreoffice24.2"
    if not os.path.exists(libreoffice_path):
        messagebox.showerror("Erreur", "Le binaire LibreOffice n'a pas été trouvé.")
        return
    
    try:
        docx_files = [f for f in os.listdir(folder_path) if f.endswith('.docx')]
        total_files = len(docx_files)
        if total_files == 0:
            messagebox.showinfo("Information", "Aucun fichier DOCX trouvé dans le dossier sélectionné.")
            return
        
        progress_bar["maximum"] = total_files
        for i, filename in enumerate(docx_files):
            docx_path = os.path.join(folder_path, filename)
            if convert_docx_to_pdf(docx_path, libreoffice_path):
                progress_bar["value"] = i + 1
                progress_label.config(text=f"Conversion des fichiers : {i + 1}/{total_files}")
                root.update_idletasks()
        
        messagebox.showinfo("Succès", "Toutes les conversions DOCX en PDF sont terminées.")
        root.destroy()  # Fermer la fenêtre de progression après la fin de la conversion
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue lors de la conversion : {e}")

def convert_all_pdf_to_docx(folder_path, progress_bar, progress_label, root):
    try:
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        total_files = len(pdf_files)
        if total_files == 0:
            messagebox.showinfo("Information", "Aucun fichier PDF trouvé dans le dossier sélectionné.")
            return
        
        progress_bar["maximum"] = total_files
        for i, filename in enumerate(pdf_files):
            pdf_path = os.path.join(folder_path, filename)
            if convert_pdf_to_docx(pdf_path):
                progress_bar["value"] = i + 1
                progress_label.config(text=f"Conversion des fichiers : {i + 1}/{total_files}")
                root.update_idletasks()
        
        messagebox.showinfo("Succès", "Toutes les conversions PDF en DOCX sont terminées.")
        root.destroy()  # Fermer la fenêtre de progression après la fin de la conversion
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue lors de la conversion : {e}")

def select_folder(conversion_type):
    folder_path = filedialog.askdirectory()
    if folder_path:
        create_progress_window(folder_path, conversion_type)

def create_progress_window(folder_path, conversion_type):
    progress_window = tk.Toplevel()
    progress_window.title("Progression de la conversion")
    progress_window.geometry("400x100")
    
    progress_label = tk.Label(progress_window, text="Conversion des fichiers : 0/0")
    progress_label.pack(pady=10)
    
    progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=10)
    
    def start_conversion():
        if conversion_type == 'docx_to_pdf':
            convert_all_docx_to_pdf(folder_path, progress_bar, progress_label, progress_window)
        elif conversion_type == 'pdf_to_docx':
            convert_all_pdf_to_docx(folder_path, progress_bar, progress_label, progress_window)
    
    progress_window.after(100, start_conversion)
    
    progress_window.protocol("WM_DELETE_WINDOW", lambda: None)  # Désactiver la fermeture de la fenêtre pendant la conversion

def create_gui():
    root = tk.Tk()
    root.title("Docx_PDF_Switch")
    root.geometry("400x200")
    
    label = tk.Label(root, text="Sélectionnez le dossier contenant les fichiers à convertir")
    label.pack(pady=20)
    
    select_docx_to_pdf_button = tk.Button(root, text="DOCX en PDF", command=lambda: select_folder('docx_to_pdf'))
    select_docx_to_pdf_button.pack(pady=10)
    
    select_pdf_to_docx_button = tk.Button(root, text="PDF en DOCX", command=lambda: select_folder('pdf_to_docx'))
    select_pdf_to_docx_button.pack(pady=10)
    
    quit_button = tk.Button(root, text="Quitter", command=root.quit)
    quit_button.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()

