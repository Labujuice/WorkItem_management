#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import glob
import re
import tempfile # For temporary files
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
import mermaid # For generating SVG from Mermaid code
import cairosvg # For converting SVG to PNG

def create_single_slide_presentation(md_content, output_path):
    """
    Parses markdown content to create a single-slide presentation.
    The slide will contain the Gantt chart as an image at the top and all other
    relevant information below, with a font size of 12.
    """
    lines = md_content.split('\n')
    
    title = "Work Status Report" # Default title
    mermaid_code = ""
    other_content_lines = []

    # --- 1. Parse Markdown Content ---

    # Find the main title (H1)
    for line in lines:
        if line.startswith('# '):
            title = line[2:].strip()
            break
            
    # Extract the content of the Mermaid block
    in_mermaid = False
    mermaid_block_start_line = -1
    mermaid_block_end_line = -1
    for i, line in enumerate(lines):
        if line.strip() == '```mermaid':
            in_mermaid = True
            mermaid_block_start_line = i
            continue
        if in_mermaid and line.strip() == '```':
            in_mermaid = False
            mermaid_block_end_line = i
            break
        if in_mermaid:
            mermaid_code += line + '\n'
    
    # Gather all other relevant content lines
    for i, line in enumerate(lines):
        # Skip title, comments, horizontal rules, and the mermaid block itself
        if line.startswith('# '): continue
        if line.strip().startswith('<!--'): continue
        if line.strip() == '---': continue
        if mermaid_block_start_line <= i <= mermaid_block_end_line: continue
        
        # Keep the line if it has content
        if line.strip():
            other_content_lines.append(line)

    # --- 2. Create the Presentation ---
    
    prs = Presentation()
    # Use "Title and Content" layout
    slide_layout = prs.slide_layouts[1] 
    slide = prs.slides.add_slide(slide_layout)
    
    # Set the slide title
    slide.shapes.title.text = title
    
    # Get the body text frame
    body_shape = slide.placeholders[1]
    tf = body_shape.text_frame
    tf.clear()  # Clear default text
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    # --- Handle Mermaid Gantt Chart ---
    gantt_image_path = None
    if mermaid_code.strip():
        print("Attempting to render Mermaid Gantt chart to image...")
        temp_svg_path = None
        temp_png_path = None
        try:
            # Create temporary file paths
            with tempfile.NamedTemporaryFile(delete=False, suffix=".svg") as temp_svg_file:
                temp_svg_path = temp_svg_file.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_png_file:
                temp_png_path = temp_png_file.name

            # 1. Generate SVG from Mermaid code directly to file
            mermaid.Mermaid(mermaid_code).to_svg(path=temp_svg_path)
            
            # 2. Convert SVG to PNG
            cairosvg.svg2png(url=temp_svg_path, write_to=temp_png_path)
            gantt_image_path = temp_png_path
            print(f"Mermaid Gantt chart rendered to: {gantt_image_path}")

        except Exception as e:
            print(f"Error rendering Mermaid chart to image: {e}")
            print("Falling back to text version of Gantt chart.")
            # Fallback to text if image generation fails
            p_mermaid = tf.add_paragraph()
            p_mermaid.text = mermaid_code
            for run in p_mermaid.runs:
                run.font.size = Pt(12)
                run.font.name = 'Courier New' # Use a monospaced font for the chart
        finally:
            if temp_svg_path and os.path.exists(temp_svg_path):
                os.remove(temp_svg_path)
            # temp_png_path will be cleaned up after prs.save()

    if gantt_image_path:
        # Position the image at the top of the slide
        left = Inches(0.5)
        top = Inches(1.5) # Below the title
        width = Inches(9)
        height = Inches(3) # Adjust as needed

        pic = slide.shapes.add_picture(gantt_image_path, left, top, width=width, height=height)
        
        # Adjust the text frame position to be below the image
        tf.top = top + height + Inches(0.2) # 0.2 inch padding
        tf.height = prs.slide_height - tf.top - Inches(0.5) # Remaining height

    # Add a separator line if an image was inserted, or if falling back to text
    if gantt_image_path or mermaid_code.strip():
        p_sep = tf.add_paragraph()
        p_sep.text = '\n' + ('-' * 40) + '\n'
        for run in p_sep.runs:
            run.font.size = Pt(12)

    # Add all other content
    for line_text in other_content_lines:
        p = tf.add_paragraph()
        
        # Handle basic indentation for bullet points
        if line_text.strip().startswith('* '):
            p.text = line_text.strip().lstrip('* ')
            p.level = 2
        elif line_text.strip().startswith('- '):
            p.text = line_text.strip().lstrip('- ')
            p.level = 1
        else:
            p.text = line_text
            p.level = 0
        
        # Set font size for every run in the paragraph
        for run in p.runs:
            run.font.size = Pt(12)

    print(f"Saving single-slide presentation to: {output_path}")
    prs.save(output_path)

    # Clean up the generated PNG file after saving the presentation
    if gantt_image_path and os.path.exists(gantt_image_path):
        os.remove(gantt_image_path)


def main():
    """
    Main function to find the markdown file, parse it, and generate a presentation.
    """
    base_path = os.path.dirname(os.path.abspath(__file__))
    people_path = os.path.join(base_path, 'people')

    try:
        md_file = glob.glob(os.path.join(people_path, '*.md'))[0]
    except IndexError:
        print("Error: No markdown file found in the /people directory.")
        return

    print(f"Processing file: {md_file}")

    with open(md_file, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    output_filename = os.path.splitext(os.path.basename(md_file))[0] + '.pptx'
    output_path = os.path.join(people_path, output_filename)
    
    create_single_slide_presentation(md_content, output_path)
    print("Presentation created successfully.")

if __name__ == '__main__':
    main()