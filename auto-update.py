#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import glob
import re
from datetime import date, timedelta
import yaml

def calculate_workdays(start_date_str, due_date_str):
    """Calculates the number of workdays (Mon-Fri) between two dates."""
    try:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
        due_date = datetime.strptime(due_date_str, '%Y-%m-%d').date()
        
        if start_date > due_date:
            return 0
            
        workdays = 0
        current_date = start_date
        while current_date <= due_date:
            if current_date.weekday() < 5:  # Monday is 0 and Sunday is 6
                workdays += 1
            current_date += timedelta(days=1)
        return workdays
    except (ValueError, TypeError):
        return 0

def update_project_files(projects_path):
    """
    Scans project files, updates estimated_workdays, and returns project data.
    """
    project_files = glob.glob(os.path.join(projects_path, '*.md'))
    all_projects = []

    for project_file in project_files:
        with open(project_file, 'r+') as f:
            content = f.read()
            
            # Extract YAML frontmatter
            match = re.search(r'^---\s*\n(.*?)\n---\s*\n', content, re.DOTALL)
            if not match:
                continue

            frontmatter_str = match.group(1)
            project_data = yaml.safe_load(frontmatter_str)

            # Get required fields
            start_date = project_data.get('start_date')
            due_date = project_data.get('due_date')
            
            if start_date and due_date:
                # Calculate and update workdays
                workdays = calculate_workdays(str(start_date), str(due_date))
                
                new_content = re.sub(
                    r'estimated_workdays:.*',
                    'estimated_workdays: {}'.format(workdays),
                    content
                )
                
                if new_content != content:
                    f.seek(0)
                    f.write(new_content)
                    f.truncate()

                project_data['estimated_workdays'] = workdays
            
            all_projects.append(project_data)

    return all_projects

def update_people_file(people_file, projects):
    """
    Updates the people file with lists of projects based on their status.
    """
    status_map = {
        'In-Progress': '進行中項目',
        'Pending': '等待進行項目',
        'Completed': '以完成項目'
    }
    
    categorized_projects = {
        '進行中項目': [],
        '等待進行項目': [],
        '以完成項目': []
    }

    for proj in projects:
        proj_status = proj.get('status')
        proj_title = proj.get('title', 'N/A')
        if proj_status in status_map:
            category = status_map[proj_status]
            categorized_projects[category].append(proj_title)

    # Build the markdown block
    update_block = "<!-- AUTO_UPDATE_START -->\n"
    for category, titles in categorized_projects.items():
        update_block += "### {}\n".format(category)
        if titles:
            for title in titles:
                update_block += "- {}\n".format(title)
        else:
            update_block += "- 無\n"
        update_block += "\n"
    update_block += "<!-- AUTO_UPDATE_END -->"

    # Read the people file and replace the block
    with open(people_file, 'r+') as f:
        content = f.read()
        # Use re.DOTALL to match across newlines
        pattern = r'<!-- AUTO_UPDATE_START -->.*?<!-- AUTO_UPDATE_END -->'
        
        if re.search(pattern, content, re.DOTALL):
            new_content = re.sub(pattern, update_block, content, flags=re.DOTALL)
        else:
            new_content = content + "\n\n" + update_block

        if new_content != content:
            f.seek(0)
            f.write(new_content)
            f.truncate()

def main():
    """Main function to run the update process."""
    base_path = os.path.dirname(os.path.abspath(__file__))
    projects_path = os.path.join(base_path, 'projects')
    people_path = os.path.join(base_path, 'people')

    # Find the single .md file in the people directory
    try:
        people_file = glob.glob(os.path.join(people_path, '*.md'))[0]
    except IndexError:
        print("Error: No markdown file found in /people directory.")
        return

    print("Starting update process...")
    
    # 1. Update project files and get data
    projects_data = update_project_files(projects_path)
    print("Updated {} project files.".format(len(projects_data)))

    # 2. Update the people file
    update_people_file(people_file, projects_data)
    print("Updated people file: {}".format(os.path.basename(people_file)))
    
    print("Update process finished.")

if __name__ == '__main__':
    # PyYAML requires datetime strings to be handled carefully.
    # Let's monkey-patch the constructor.
    from datetime import datetime
    yaml.SafeLoader.add_constructor('tag:yaml.org,2002:timestamp', 
        lambda loader, node: loader.construct_scalar(node))

    main()
