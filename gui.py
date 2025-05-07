import os
import fitz
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Enable high-DPI support on Windows
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

def embed_slides_on_template(input_files, output_folder, output_name, template_path, positions, zoom_factor):
    slides_per_page = 4
    output_doc = fitz.open()
    template_doc = fitz.open(template_path)

    slide_count = 0
    for file in input_files:
        if file.lower().endswith('.pdf'):
            doc = fitz.open(file)
            for page_num in range(len(doc)):
                if slide_count % slides_per_page == 0:
                    output_page = output_doc.new_page(width=template_doc[0].rect.width,
                                                      height=template_doc[0].rect.height)
                    output_page.show_pdf_page(output_page.rect, template_doc, 0)
                pos = positions[slide_count % slides_per_page]
                slide = doc.load_page(page_num)

                center_x = (pos[0] + pos[2]) / 2
                center_y = (pos[1] + pos[3]) / 2
                width = (pos[2] - pos[0]) * zoom_factor
                height = (pos[3] - pos[1]) * zoom_factor

                zoomed_rect = fitz.Rect(center_x - width / 2, center_y - height / 2,
                                        center_x + width / 2, center_y + height / 2)

                output_page.show_pdf_page(zoomed_rect, doc, page_num)
                slide_count += 1

    output_path = os.path.join(output_folder, output_name)
    output_doc.save(output_path)
    messagebox.showinfo("Success", f"PDF saved to {output_path}")


def generate_pdf():
    if not input_files:
        messagebox.showerror("Missing Info", "Please select input PDF files.")
        return

    output_folder = output_path.get()
    output_name = output_name_entry.get()
    zoom = float(zoom_entry.get())
    template_path = os.path.join(os.path.dirname(__file__), "template.pdf")

    if not all([output_folder, output_name]):
        messagebox.showerror("Missing Info", "Please fill in all fields.")
        return

    aspect_ratio = get_pdf_aspect_ratio(input_files[0])
    if not aspect_ratio:
        messagebox.showerror("Error", "Could not determine aspect ratio from the first file.")
        return

    template_doc = fitz.open(template_path)
    x0_left = 17.15
    y0 = 18.075
    square_length = 13.85
    line_width = 0.55
    n_squares = 13
    add_y_pixels = n_squares * square_length + (n_squares - 1) * line_width

    y_coords = [y0 + i * (add_y_pixels + square_length + 2 * line_width) for i in range(4)]

    positions = [(x0_left, y, x0_left + aspect_ratio * add_y_pixels, y + add_y_pixels) for y in y_coords]

    embed_slides_on_template(input_files, output_folder, output_name, template_path, positions, zoom)


def get_pdf_aspect_ratio(file_path):
    with fitz.open(file_path) as doc:
        rect = doc[0].rect
        return rect.width / rect.height


def browse_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_path.set(folder)


def select_input_files():
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    if files:
        input_files.clear()
        input_files.extend(files)
        update_input_listbox()


def update_input_listbox():
    input_listbox.delete(0, tk.END)
    for f in input_files:
        input_listbox.insert(tk.END, os.path.basename(f))


# GUI SETUP
root = tk.Tk()
root.title("PDF Slide Combiner")
root.geometry("800x600")
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

frame = ttk.Frame(root, padding=10)
frame.grid(sticky="nsew")
frame.columnconfigure(1, weight=1)

input_files = []
# Set default output to Desktop
output_path = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))

# Input file list
ttk.Label(frame, text="Input PDF Files").grid(row=0, column=0, sticky=tk.W)
ttk.Button(frame, text="Select Files", command=select_input_files).grid(row=0, column=1, sticky=tk.W)
input_listbox = tk.Listbox(frame, height=10)
input_listbox.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=5)
frame.rowconfigure(1, weight=1)

# Output folder
ttk.Label(frame, text="Output Folder").grid(row=2, column=0, sticky=tk.W, pady=5)
ttk.Entry(frame, textvariable=output_path).grid(row=2, column=1, sticky="ew")
ttk.Button(frame, text="Browse", command=browse_output_folder).grid(row=2, column=2, sticky=tk.W)

# Output file name
ttk.Label(frame, text="Output File Name").grid(row=3, column=0, sticky=tk.W, pady=5)
output_name_entry = ttk.Entry(frame)
output_name_entry.insert(0, "output_notes.pdf")
output_name_entry.grid(row=3, column=1, columnspan=2, sticky="ew")

# Zoom factor
ttk.Label(frame, text="Zoom Factor").grid(row=4, column=0, sticky=tk.W, pady=5)
zoom_entry = ttk.Entry(frame)
zoom_entry.insert(0, "1.0")
zoom_entry.grid(row=4, column=1, columnspan=2, sticky="ew")

# Generate button
ttk.Button(frame, text="Generate PDF", command=generate_pdf).grid(row=5, column=0, columnspan=3, pady=20)

root.mainloop()
