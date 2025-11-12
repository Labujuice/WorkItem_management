#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import glob
import re
from datetime import date, timedelta, datetime
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
    Scans project files, updates estimated_workdays, and returns project data with paths.
    """
    project_files = glob.glob(os.path.join(projects_path, '*.md'))
    all_projects_with_paths = []

    for project_file in project_files:
        with open(project_file, 'r+') as f:
            content = f.read()
            
            match = re.search(r'^---\s*\n(.*?)\n---\s*\n', content, re.DOTALL)
            if not match:
                continue

            frontmatter_str = match.group(1)
            project_data = yaml.safe_load(frontmatter_str)

            start_date = project_data.get('start_date')
            due_date = project_data.get('due_date')
            
            if start_date and due_date:
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
            
            all_projects_with_paths.append({'data': project_data, 'path': project_file})

    return all_projects_with_paths

def update_people_file(people_file, projects_with_paths):
    """
    Updates the people file with lists of projects based on their status and new format.
    """
    status_map = {
        'In-Progress': 'In-Progress',
        'Pending': 'Pending',
        'Completed': 'Completed'
    }
    
    categorized_projects = {
        'In-Progress': [],
        'Pending': [],
        'Completed': []
    }

    today = date.today()

    for item in projects_with_paths:
        proj_data = item['data']
        proj_path = item['path']
        
        proj_status = proj_data.get('status')
        if proj_status not in status_map:
            continue

        # --- Progress Calculation ---
        progress = 0
        if proj_status == 'Completed':
            progress = 100
        elif proj_status == 'In-Progress':
            start_date_str = str(proj_data.get('start_date', ''))
            total_workdays = proj_data.get('estimated_workdays', 0)
            
            if total_workdays and total_workdays > 0:
                days_passed = calculate_workdays(start_date_str, today.strftime('%Y-%m-%d'))
                progress = min(int((days_passed / total_workdays) * 100), 100)

        # --- Format Output String ---
        proj_id = proj_data.get('id', 'N/A')
        proj_title = proj_data.get('title', 'N/A')
        proj_project = proj_data.get('project', '')
        
        # Create relative path for the link
        relative_path = os.path.relpath(proj_path, os.path.dirname(people_file))

        output_line = "[{id} {title} {project} {progress}%]({path})".format(
            id=proj_id,
            title=proj_title,
            project=proj_project,
            progress=progress,
            path=relative_path.replace(os.path.sep, '/') # Ensure forward slashes for markdown links
        )
        
        category = status_map[proj_status]
        categorized_projects[category].append(output_line)

    # Build the markdown block
    update_block = "<!-- AUTO_UPDATE_START -->\n"
    for category, lines in categorized_projects.items():
        update_block += f"### {category}\n"
        if lines:
            for line in lines:
                update_block += f"- {line}\n"
        else:
            update_block += "- 無\n"
        update_block += "\n"
    update_block += "<!-- AUTO_UPDATE_END -->"

    # Read the people file and replace the block
    with open(people_file, 'r+') as f:
        content = f.read()
        pattern = r'<!-- AUTO_UPDATE_START -->.*?<!-- AUTO_UPDATE_END -->'
        
        if re.search(pattern, content, re.DOTALL):
            new_content = re.sub(pattern, update_block, content, flags=re.DOTALL)
        else:
            new_content = content.rstrip() + "\n\n" + update_block

        if new_content != content:
            f.seek(0)
            f.write(new_content)
            f.truncate()

def main():
    """Main function to run the update process."""
    base_path = os.path.dirname(os.path.abspath(__file__))
    projects_path = os.path.join(base_path, 'projects')
    people_path = os.path.join(base_path, 'people')

    try:
        people_file = glob.glob(os.path.join(people_path, '*.md'))[0]
    except IndexError:
        print("Error: No markdown file found in /people directory.")
        return

    print("Starting update process...")
    
    projects_data = update_project_files(projects_path)
    print("Updated {} project files.".format(len(projects_data)))

    update_people_file(people_file, projects_data)
    print("Updated people file: {}".format(os.path.basename(people_file)))
    
    print("Update process finished.")

if __name__ == '__main__':
    from datetime import datetime
    yaml.SafeLoader.add_constructor('tag:yaml.org,2002:timestamp', 
        lambda loader, node: loader.construct_scalar(node))

    main()