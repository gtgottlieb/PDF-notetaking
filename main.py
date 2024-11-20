import fitz
import os


# Template is 39 x 56
# 1 (inner) Square is approximately 13.85 x 13.85
# The (inner) starting point is (17.15, 18.075) (x,y)
# Each line is 0.55 wide

def ensure_directories_exist(input_folder, output_folder):
    """Ensure the existence of Input Folder and Output Folder."""
    if not os.path.exists(input_folder):
        os.makedirs(input_folder)
        print(f"Created directory: {input_folder}")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created directory: {output_folder}")

def get_template_page(template_path):
    template_doc = fitz.open(template_path)
    template_page = template_doc.load_page(0)  # Assuming the template is a single page
    return template_doc, template_page


def draw_grid_on_template(template_path, output_path):
    doc = fitz.open(template_path)
    page = doc[0]  # Assuming you want to draw on the first page
    w, h = page.rect.width, page.rect.height

    # Draw horizontal and vertical lines
    step_size = 10  # Change step size as needed for more or less granularity
    # for i in range(0, int(w), step_size):
    #    page.draw_line((i, 0), (i, h), color=(1, 0, 0), width=0.01)
    # for j in range(0, int(h), step_size):
    #    page.draw_line((0, j), (w, j), color=(0, 1, 0), width=0.01)

    x0 = 17.15
    y0 = 18.075
    square_length = 13.85
    line_width = 0.55
    n_squares = 13

    add_y_pixels_slide = n_squares * square_length + (n_squares - 0.5) * line_width

    rect_slide = fitz.Rect(x0, y0, x0 + 4 / 3 * add_y_pixels_slide, y0 + add_y_pixels_slide)

    rect3 = fitz.Rect(x0, y0, x0 + 13.85, y0 + 13.85)
    rect4 = fitz.Rect(x0, y0, x0 + 14.4, y0 + 14.4)

    page.draw_rect(rect_slide, width=0.01)
    # page.draw_rect(rect3, width=0.01)
    # page.draw_rect(rect4, width=0.01)

    doc.save(output_path)
    os.startfile(output_path)


def embed_slides_on_template(input_folder, output_folder, output_name, template_path, positions, slides_per_page=4):
    output_doc = fitz.open()  # Create a new output document
    template_doc = fitz.open(template_path)  # Load the template document

    output_page = None  # Placeholder for the current output page

    for root, dirs, files in os.walk(input_folder):
        for file in files:
            if file.endswith('.pdf'):
                slide_count = 0
                doc = fitz.open(os.path.join(root, file))
                for page_num in range(len(doc)):
                    if slide_count % slides_per_page == 0:
                        # Duplicate the template for a new output page
                        output_page = output_doc.new_page(-1, width=template_doc[0].rect.width,
                                                          height=template_doc[0].rect.height)
                        output_page.show_pdf_page(output_page.rect, template_doc, 0)

                    # Calculate the position and size for the current slide
                    pos = positions[slide_count % slides_per_page]
                    slide = doc.load_page(page_num)
                    slide_rect = slide.rect
                    target_rect = fitz.Rect(pos[0], pos[1], pos[2], pos[3])
                    t_width = template_doc[0].rect.width
                    writing_rect = fitz.Rect(pos[0], pos[1], t_width - pos[0], pos[3])

                    # Embed the slide onto the current output page
                    output_page.show_pdf_page(target_rect, doc, page_num)
                    output_page.draw_rect(target_rect, width=0.55)
                    output_page.draw_rect(writing_rect, width=0.55)
                    slide_count += 1

                doc.close()

    output_pdf_path = os.path.join(output_folder, output_name)
    output_doc.save(output_pdf_path)
    print(f"Compiled notes saved to {output_pdf_path}")
    os.startfile(output_pdf_path)


def main():
    """
    Edit the following if needs be:
    """
    aspect_ratio = 4/3
    output_name = "Math Methods Week 4 (up to GT2).pdf"

    input_folder = "Input Folder"   
    output_folder = "Output Folder"
    template_path = "Template Folder/Template.pdf"

    ensure_directories_exist(input_folder, output_folder)

    x0 = 17.15
    y0 = 18.075
    square_length = 13.85
    line_width = 0.55
    n_squares = 13
    add_y_pixels_slide = n_squares * square_length + (n_squares - 1) * line_width
    y1 = y0 + add_y_pixels_slide + square_length + 2 * line_width
    y2 = y1 + add_y_pixels_slide + square_length + 2 * line_width
    y3 = y2 + add_y_pixels_slide + square_length + 2 * line_width

    positions = [(x0, y0, x0 + aspect_ratio * add_y_pixels_slide, y0 + add_y_pixels_slide),
                 (x0, y1, x0 + aspect_ratio * add_y_pixels_slide, y1 + add_y_pixels_slide),
                 (x0, y2, x0 + aspect_ratio * add_y_pixels_slide, y2 + add_y_pixels_slide),
                 (x0, y3, x0 + aspect_ratio * add_y_pixels_slide, y3 + add_y_pixels_slide)]

    # scaled_positions = [(x * scaling_factor for x in tup) for tup in positions]
    # scaled_positions = [tuple(gen) for gen in scaled_positions]

    embed_slides_on_template(input_folder, output_folder, output_name, template_path, positions)

    # For checking locations of template:
    # draw_grid_on_template(template_path, os.path.join(output_folder, "template with gridlines.pdf"))


if __name__ == "__main__":
    main()
