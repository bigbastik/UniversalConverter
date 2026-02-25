import os
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from PIL import Image

# Optional imports (non fanno crash se mancano)
try:
    from pdf2image import convert_from_path
except:
    convert_from_path = None

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class UniversalConverter(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- ICONA DEFINITIVA COMPATIBILE PYINSTALLER ---
        if getattr(sys, 'frozen', False):
            # Se eseguito come exe PyInstaller
            base_path = sys._MEIPASS
        else:
            # Se eseguito come script Python normale
            base_path = os.path.dirname(os.path.abspath(__file__))

        ico_path = os.path.join(base_path, "favicon.ico")
        try:
            self.wm_iconbitmap(ico_path)
        except Exception as e:
            print(f"Impossibile caricare icona: {e}")

        self.title("Universal File Converter PRO")
        self.geometry("720x520")
        self.file_paths = []

        # Fix chiusura EXE (evita processo in background)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        title = ctk.CTkLabel(
            self,
            text="Universal Converter\n(PDF, DOCX, XLSX, PPTX, JPG, PNG)",
            font=("Arial", 24, "bold")
        )
        title.pack(pady=20)

        self.select_btn = ctk.CTkButton(
            self,
            text="📂 Seleziona File (anche multipli)",
            command=self.select_files,
            height=45
        )
        self.select_btn.pack(pady=10, padx=40, fill="x")

        self.status = ctk.CTkLabel(self, text="Stato: In attesa...")
        self.status.pack(pady=10)

        # Conversioni principali
        self.create_button("PDF → DOCX", self.pdf_to_docx)
        self.create_button("DOCX/XLSX/PPTX → PDF", self.office_to_pdf_func)
        self.create_button("JPG/PNG → PDF (multipagina)", self.images_to_pdf)
        self.create_button("PDF → Immagini (una per pagina)", self.pdf_to_images)
        self.create_button("PNG → JPG", self.png_to_jpg)
        self.create_button("JPG → PNG", self.jpg_to_png)

    def create_button(self, text, command):
        btn = ctk.CTkButton(self, text=text, command=command, height=40)
        btn.pack(pady=6, padx=60, fill="x")

    def on_closing(self):
        try:
            self.status.configure(text="Chiusura in corso...")
            self.update_idletasks()
        except:
            pass

        try:
            self.quit()
            self.destroy()
        except:
            pass

        # Uccide definitivamente il processo (fix EXE in background)
        os._exit(0)

    def select_files(self):
        paths = filedialog.askopenfilenames()
        if paths:
            self.file_paths = list(paths)
            self.status.configure(text=f"{len(self.file_paths)} file selezionati")
        else:
            self.file_paths = []
            self.status.configure(text="Stato: Nessun file selezionato")

    def pdf_to_docx(self):
        if not self.file_paths:
            messagebox.showerror("Errore", "Nessun file selezionato")
            return

        converted = 0
        errors = []

        for file in self.file_paths:
            if file.lower().endswith(".pdf"):
                output = os.path.splitext(file)[0] + ".docx"
                self.status.configure(text=f"Convertendo: {os.path.basename(file)}")
                self.update_idletasks()

                try:
                    cv = Converter(file)
                    cv.convert(output)
                    cv.close()
                    converted += 1
                except Exception as e:
                    errors.append(f"{os.path.basename(file)}: {str(e)}")

        if converted == 0:
            messagebox.showerror("Errore", "Nessun PDF valido convertito")
        else:
            msg = f"Convertiti {converted} file PDF in DOCX"
            if errors:
                msg += "\n\nErrori:\n" + "\n".join(errors)
            messagebox.showinfo("Completato", msg)

    def office_to_pdf_func(self):
        try:
            import win32com.client
        except ImportError:
            messagebox.showerror(
                "Errore",
                "Installa: pip install pywin32\nE assicurati che Microsoft Office sia installato"
            )
            return

        if not self.file_paths:
            messagebox.showerror("Errore", "Nessun file selezionato")
            return

        converted = 0
        errors = []

        # --- DOCX ---
        docx_files = [os.path.abspath(f) for f in self.file_paths if f.lower().endswith(".docx")]
        if docx_files:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            for file in docx_files:
                output = os.path.splitext(file)[0] + ".pdf"
                self.status.configure(text=f"Convertendo: {os.path.basename(file)}")
                self.update_idletasks()
                try:
                    doc = word.Documents.Open(file, ReadOnly=True)
                    doc.SaveAs(output, FileFormat=17)
                    doc.Close(False)
                    converted += 1
                except Exception as e:
                    errors.append(f"DOCX {os.path.basename(file)}: {str(e)}")
            word.Quit()

        # --- XLSX ---
        xlsx_files = [os.path.abspath(f) for f in self.file_paths if f.lower().endswith(".xlsx")]
        if xlsx_files:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            for file in xlsx_files:
                output = os.path.splitext(file)[0] + ".pdf"
                self.status.configure(text=f"Convertendo: {os.path.basename(file)}")
                self.update_idletasks()
                try:
                    wb = excel.Workbooks.Open(file, ReadOnly=True)
                    wb.ExportAsFixedFormat(0, output)
                    wb.Close(False)
                    converted += 1
                except Exception as e:
                    errors.append(f"XLSX {os.path.basename(file)}: {str(e)}")
            excel.Quit()

        # --- PPTX ---
        pptx_files = [os.path.abspath(f) for f in self.file_paths if f.lower().endswith(".pptx")]
        if pptx_files:
            powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
            powerpoint.Visible = False
            for file in pptx_files:
                output = os.path.splitext(file)[0] + ".pdf"
                self.status.configure(text=f"Convertendo: {os.path.basename(file)}")
                self.update_idletasks()
                try:
                    pres = powerpoint.Presentations.Open(file, WithWindow=False)
                    pres.SaveAs(output, 32)
                    pres.Close()
                    converted += 1
                except Exception as e:
                    errors.append(f"PPTX {os.path.basename(file)}: {str(e)}")
            powerpoint.Quit()

        if converted == 0:
            messagebox.showerror("Errore", "Nessun file Office convertito.\n" +
                                 "\n".join(errors) if errors else "")
        else:
            msg = f"Convertiti {converted} file in PDF"
            if errors:
                msg += "\n\nErrori:\n" + "\n".join(errors)
            messagebox.showinfo("Completato", msg)

    def images_to_pdf(self):
        if not self.file_paths:
            messagebox.showerror("Errore", "Nessun file selezionato")
            return

        images = []
        for file in self.file_paths:
            if file.lower().endswith((".jpg", ".jpeg", ".png")):
                try:
                    img = Image.open(file).convert("RGB")
                    images.append(img)
                except:
                    pass

        if not images:
            messagebox.showerror("Errore", "Nessuna immagine valida selezionata")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if save_path:
            images[0].save(save_path, save_all=True, append_images=images[1:])
            messagebox.showinfo("Successo", "PDF multipagina creato")

    def pdf_to_images(self):
        if convert_from_path is None:
            messagebox.showerror(
                "Errore",
                "Installa poppler + pip install pdf2image"
            )
            return

        if not self.file_paths:
            messagebox.showerror("Errore", "Nessun file selezionato")
            return

        converted = 0

        for file in self.file_paths:
            if file.lower().endswith(".pdf"):
                try:
                    pages = convert_from_path(file)
                    base = os.path.splitext(file)[0]

                    for i, page in enumerate(pages):
                        page.save(f"{base}_pagina_{i+1}.png", "PNG")

                    converted += 1
                except:
                    pass

        if converted == 0:
            messagebox.showerror("Errore", "Nessun PDF convertito")
        else:
            messagebox.showinfo("Successo", f"Convertiti {converted} PDF in immagini")

    def png_to_jpg(self):
        if not self.file_paths:
            messagebox.showerror("Errore", "Nessun file selezionato")
            return

        converted = 0
        for file in self.file_paths:
            if file.lower().endswith(".png"):
                try:
                    img = Image.open(file).convert("RGB")
                    img.save(os.path.splitext(file)[0] + ".jpg", "JPEG")
                    converted += 1
                except:
                    pass

        messagebox.showinfo("Successo", f"PNG → JPG completato ({converted} file)")

    def jpg_to_png(self):
        if not self.file_paths:
            messagebox.showerror("Errore", "Nessun file selezionato")
            return

        converted = 0
        for file in self.file_paths:
            if file.lower().endswith((".jpg", ".jpeg")):
                try:
                    img = Image.open(file)
                    img.save(os.path.splitext(file)[0] + ".png", "PNG")
                    converted += 1
                except:
                    pass

        messagebox.showinfo("Successo", f"JPG → PNG completato ({converted} file)")


if __name__ == "__main__":
    app = UniversalConverter()
    try:
        app.mainloop()
    finally:
        os._exit(0)