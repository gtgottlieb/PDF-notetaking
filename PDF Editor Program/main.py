import os
from PyPDF2 import PdfMerger, PageObject, PdfReader, PdfWriter, Transformation
from pdf2image import convert_from_path
from PIL import Image
import shutil


def purge(folder_path):
    # Get a list of all files in the folder
    files = os.listdir(folder_path)

    # Iterate over each file in the folder
    for file_name in files:
        # Construct the full path to the file
        file_path = os.path.join(folder_path, file_name)

        # Check if the file is a regular file (not a directory)
        if os.path.isfile(file_path):
            # Delete the file
            os.remove(file_path)


# Convert PDF pages to images
def pdf_to_images(input_folder, output_folder, size):
    images = []
    for filename in os.listdir(input_folder):
        if filename.endswith('.pdf'):
            images = images + convert_from_path(os.path.join(input_folder, filename), size=size)
    for i, image in enumerate(images):
        image.save(f'{output_folder}/page_{i + 1}.png', 'PNG')





def paste_images(background_image_path, image_folder_path, output_folder_path):
    # Open the background image
    background = Image.open(background_image_path)
    position_index = 0
    page_number = 1

    # Loop through image files in the folder
    for i in range(len(os.listdir(image_folder_path))):
        filename = f"page_{i + 1}.png"

        # Check if the file is an image file
        if filename.endswith('.png'):
            # Open the image to be pasted
            image_path = os.path.join(image_folder_path, filename)
            image_to_paste = Image.open(image_path)

            # Calculate position based on position index
            position = positions[position_index]

            # Paste the image onto the background image at the specified position
            background.paste(image_to_paste, position)

            # Increment position index
            position_index += 1
            if position_index == 4:
                # Save the modified background image and scale page
                output_filename = f"Page_{page_number}.pdf"
                background.save(os.path.join(output_folder_path, output_filename), 'PDF')
                scale_pdf(os.path.join(output_folder_path, output_filename))

                # Reset position index and move to the next page
                position_index = 0
                page_number += 1

        # Save the final modified background image
        if position_index != 0:
            output_filename = f"Page_{page_number}.pdf"
            background.save(os.path.join(output_folder_path, output_filename), 'PDF')


def scale_pdf(pdf_path):
    output_path = pdf_path
    pdf = PdfReader(pdf_path)
    writer = PdfWriter()
    page = pdf.pages[0]
    page.scale_by(0.1)
    writer.add_page(page)
    with open(output_path, "wb+") as f:
        writer.write(f)

    writer.close()


def merge_pdfs(input_folder, final_folder):
    output_pdf = os.path.join(final_folder, "Final.pdf")
    # Create a PdfFileMerger object
    merger = PdfMerger()

    # Loop through PDF files in the folder
    for k in range(len(os.listdir(input_folder))):
        filename = f"Page_{k + 1}.pdf"

        # Open each PDF file and append it to the merger object
        pdf_path = os.path.join(input_folder, filename)

        merger.append(pdf_path)

    # Write the merged PDF to the output file
    with open(output_pdf, 'wb') as f:
        merger.write(f)

    merger.close()


input_folder = "Input Folder"
images_folder = "Temp/Images Folder"
output_folder = "Temp/Output Folder"
final_output = "Output Final"
template = "Template Folder/Template.png"

input_aspect_ratio = 4/3 #x/y

positions = [(184, 194), (184, 194 + 154*14 - 5), (184, 194 + 154*28 - 10), (184, 194 + 154*42 - 15)]

size_template = (int(4 * 1588), int(4 * 2246))
size_slides = (int(4 * 497 * input_aspect_ratio), int(4 * 497))


# Restart
print(f"Purging temp files...")
purge(images_folder)
purge(output_folder)

print(f"Converting PDFs to PNGs...")
pdf_to_images(input_folder, images_folder, size_slides)

# pdf_to_images(r"C:\Users\gtgot\OneDrive\Misc\PDF Editor Program", r"C:\Users\gtgot\OneDrive\Misc\PDF Editor Program", size_template)


print(f"Pasting images & resizing pages...")
paste_images(template, images_folder, output_folder)

print(f"Merging PDFs...")
merge_pdfs(output_folder, final_output)


print(f"Purging temp files...")
purge(images_folder)
purge(output_folder)

