#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import glob
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE

def create_single_slide_presentation(md_content, output_path):
    """
    Parses markdown content to create a single-slide presentation.
    The slide will contain the Gantt chart at the top and all other
    relevant information below, with a font size of 12.
    """
    lines = md_content.split('\n')
    
    title = "Work Status Report" # Default title
    mermaid_content = ""
    other_content_lines = []

    # --- 1. Parse Markdown Content ---

    # Find the main title (H1)
    for line in lines:
        if line.startswith('# '):
            title = line[2:].strip()
            break
            
    # Extract the content of the Mermaid block
    in_mermaid = False
    mermaid_lines = []
    for line in lines:
        if line.strip() == '```mermaid':
            in_mermaid = True
            continue
        if in_mermaid and line.strip() == '```':
            in_mermaid = False
            continue # Also skip the closing backticks line
        if in_mermaid:
            mermaid_lines.append(line)
    mermaid_content = "\n".join(mermaid_lines)

    # Gather all other relevant content lines
    in_mermaid = False
    for line in lines:
        # Skip title, comments, horizontal rules, and the mermaid block itself
        if line.startswith('# '): continue
        if line.strip().startswith('<!--'): continue
        if line.strip() == '---': continue
        if line.strip() == '```mermaid':
            in_mermaid = True
            continue
        if in_mermaid and line.strip() == '```':
            in_mermaid = False
            continue
        if in_mermaid: continue
        
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

    # Add Mermaid Gantt chart content first
    p_mermaid = tf.add_paragraph()
    p_mermaid.text = mermaid_content
    # Set font for all runs in the paragraph
    for run in p_mermaid.runs:
        run.font.size = Pt(12)
        run.font.name = 'Courier New' # Use a monospaced font for the chart

    # Add a separator line
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