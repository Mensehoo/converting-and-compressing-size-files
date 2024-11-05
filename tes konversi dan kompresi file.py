import tkinter as tk
from docx2pdf import convert as docx_to_pdf
import comtypes.client
import zipfile
import fitz
from pptx.util import Inches
from pptx import Presentation
from pptx.util import Pt
from PIL import Image
import tempfile
import os
from tkinter import filedialog, messagebox

# Merge JPG and Convert to PDF Function
def jpg_to_pdf():
    files = filedialog.askopenfilenames(filetypes=[("JPG Files", "*.jpg")])
    if not files:
        return
    images = [Image.open(f).convert("RGB") for f in files]
    pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    if pdf_path:
        images[0].save(pdf_path, save_all=True, append_images=images[1:])
        messagebox.showinfo("Success”, “JPG successfully converted to PDF!”)

# Word to PDF Conversion Function
def word_to_pdf():
    file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    if file:
        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if pdf_path:
            docx_to_pdf(file, pdf_path)
            messagebox.showinfo("Success", "Word successfully converted to PDF!")

# PPT to PDF Conversion Function
def ppt_to_pdf():
    file = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx")])
    if file:
        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if pdf_path:
            # Use absolute paths to avoid path problems
            file = os.path.abspath(file)
            pdf_path = os.path.abspath(pdf_path)

            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = 1
            try:
                ppt = powerpoint.Presentations.Open(file)
                ppt.SaveAs(pdf_path, 32)  # 32 = pdf format
                ppt.Close()
                messagebox.showinfo("Success", "PPT successfully converted to PDF!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to convert: {e}")
            finally:
                powerpoint.Quit()

# PDF Compress Function
def compress_pdf():
    file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file:
        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if pdf_path:
            pdf_document = fitz.open(file)
            
            # Save to new PDF with image compression and junk content removal
            pdf_document.save(pdf_path, garbage=4, deflate=True, deflate_images=True)
            pdf_document.close()
            
            messagebox.showinfo("Success", "PDF successfully compressed to a smaller size!")

# PPT Compress Function
def compress_ppt():
    file = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx")])
    if file:
        ppt = Presentation(file)
        output_path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint Files", "*.pptx")])

        if output_path:
            for slide in ppt.slides:
                for shape in slide.shapes:
                    # Compress images
                    if shape.shape_type == 13:  # Image shape
                        image = shape.image
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_img:
                            # Save temporary image in low quality
                            image_bytes = image.blob
                            with open(tmp_img.name, "wb") as f:
                                f.write(image_bytes)

                            # Open image and compress
                            img = Image.open(tmp_img.name)
                            if img.mode == 'RGBA':  # Convert from RGBA to RGB if necessary
                                img = img.convert('RGB')
                            img = img.resize((int(img.width * 0.4), int(img.height * 0.4)), Image.LANCZOS)  # Resize to 40%
                            img.save(tmp_img.name, quality=50)  # Save with lower image quality

                            # Replace the old image with the new one
                            shape.element.getparent().remove(shape.element)
                            slide.shapes.add_picture(tmp_img.name, shape.left, shape.top, shape.width, shape.height)

                        os.remove(tmp_img.name)  # Delete temporary files

                    # Reduce text size if applicable
                    elif shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if run.font.size and run.font.size.pt > 10:  # Set a minimum limit for font size
                                    run.font.size = Pt(max(10, run.font.size.pt * 0.8))  # Reduce font size by 80%
                    
                    # Remove other unnecessary form elements (optional)
                    if shape.shape_type not in [1, 13]:  # For example, remove elements other than text and images.
                        shape._element.getparent().remove(shape._element)

            # Save the results
            ppt.save(output_path)
            messagebox.showinfo("Success", "PPT successfully compressed to a smaller size!")
            
#GUI with Tkinter
def create_gui():
    window = tk.Tk()
    window.title("File Converter & Compressor")
    window.geometry("400x400")

    tk.Button(window, text="Merge JPG and Convert to PDF", command=jpg_to_pdf).pack(pady=10)
    tk.Button(window, text="Convert Word to PDF", command=word_to_pdf).pack(pady=10)
    tk.Button(window, text="COnvert PPT to PDF", command=ppt_to_pdf).pack(pady=10)
    tk.Button(window, text="Compress PDF", command=compress_pdf).pack(pady=10)
    tk.Button(window, text="Compress PPT", command=compress_ppt).pack(pady=10)

    window.mainloop()

if __name__ == "__main__":
    create_gui()
