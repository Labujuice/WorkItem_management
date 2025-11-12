#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import glob
import re
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from markdown_it import MarkdownIt

def parse_markdown_to_slides(md_content):
    """
    Parses markdown content and organizes it into a slide structure.
    Returns a list of dictionaries, where each dict represents a slide.
    """
    slides = []
    lines = md_content.split('\n')
    
    if not lines:
        return []

    # --- Title Slide ---
    title = ""
    subtitle_lines = []
    content_started = False
    
    # Find first H1 for title
    for i, line in enumerate(lines):
        if line.startswith('# '):
            title = line[2:].strip()
            # Capture subsequent lines as subtitle until '---'
            for sub_line in lines[i+1:]:
                if sub_line.strip() == '---':
                    break
                if sub_line.strip():
                    subtitle_lines.append(sub_line.strip())
            break # Found title slide, move on
    
    slides.append({
        'title': title,
        'content': subtitle_lines
    })

    # --- Content Slides ---
    current_slide_content = []
    current_slide_title = ""

    for line in lines:
        # Find H3 headings for new slide titles
        if line.startswith('### '):
            # If there's a pending slide, save it first
            if current_slide_title:
                slides.append({
                    'title': current_slide_title,
                    'content': current_slide_content
                })
            
            # Start a new slide
            current_slide_title = line[4:].strip()
            current_slide_content = []
        # If we are in a content slide, collect lines
        elif current_slide_title:
            # Ignore comments and empty lines
            clean_line = line.strip()
            if clean_line and not clean_line.startswith('<!--') and not clean_line.startswith('```'):
                # Add indentation for bullets
                if clean_line.startswith('* '):
                    current_slide_content.append({'text': clean_line.lstrip('* '), 'level': 1})
                elif clean_line.startswith('- '):
                    current_slide_content.append({'text': clean_line.lstrip('- '), 'level': 0})
                else: # a line that is not a bullet point
                    current_slide_content.append({'text': clean_line, 'level': 0})


    # Add the last slide
    if current_slide_title:
        slides.append({
            'title': current_slide_title,
            'content': current_slide_content
        })
        
    return slides

def create_presentation(slides_data, output_path):
    """
    Creates a PowerPoint presentation from the parsed slide data.
    """
    prs = Presentation()
    
    # Title Slide Layout
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    if slides_data:
        title.text = slides_data[0]['title']
        subtitle.text = "\n".join(slides_data[0]['content'])

    # Content Slides Layout
    content_layout = prs.slide_layouts[1]
    for slide_info in slides_data[1:]:
        slide = prs.slides.add_slide(content_layout)
        title = slide.shapes.title
        body = slide.placeholders[1]
        
        title.text = slide_info['title']
        
        tf = body.text_frame
        tf.clear() # Clear existing text
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        for item in slide_info['content']:
            p = tf.add_paragraph()
            p.text = item['text']
            p.level = item['level']

    print(f"Saving presentation to: {output_path}")
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

    slides_data = parse_markdown_to_slides(md_content)
    
    if not slides_data:
        print("No content found to generate slides.")
        return

    output_filename = os.path.splitext(os.path.basename(md_file))[0] + '.pptx'
    output_path = os.path.join(people_path, output_filename)
    
    create_presentation(slides_data, output_path)
    print("Presentation created successfully.")

if __name__ == '__main__':
    main()
