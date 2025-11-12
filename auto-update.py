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

            report_match = re.search(r'## 進度報告\n(.*?)(?=\n## |\Z)', content, re.DOTALL)
            project_data['progress_report'] = report_match.group(1).strip() if report_match else ""

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
        'In-Progress': 'In-Progress Projects',
        'Pending': 'Pending Projects',
        'Completed': 'Completed Projects'
    }
    
    categorized_projects = {
        'In-Progress Projects': [],
        'Pending Projects': [],
        'Completed Projects': []
    }

    gantt_sections = {
        "In-Progress": [],
        "Pending": [],
        "Completed": []
    }
    completed_projects_for_gantt = []
    today = date.today()

    for item in projects_with_paths:
        proj_data = item['data']
        proj_path = item['path']
        
        proj_status = proj_data.get('status')
        if proj_status not in status_map:
            continue

        # --- Gantt Chart Data Collection ---
        start_date_str = str(proj_data.get('start_date', ''))
        due_date_str = str(proj_data.get('due_date', ''))
        proj_title = proj_data.get('title', 'N/A')

        if start_date_str and due_date_str and 'TBD' not in start_date_str:
            task_line = f"    {proj_title} :{start_date_str}, {due_date_str}"
            if proj_status == 'In-Progress':
                gantt_sections["In-Progress"].append(task_line)
            elif proj_status == 'Pending':
                gantt_sections["Pending"].append(task_line)
            elif proj_status == 'Completed':
                end_date_str = str(proj_data.get('actual_end_date') or proj_data.get('due_date'))
                try:
                    end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
                    completed_projects_for_gantt.append({'date': end_date, 'task': task_line})
                except (ValueError, TypeError):
                    continue

        # --- Progress Calculation ---
        progress = 0
        if proj_status == 'Completed':
            progress = 100
        elif proj_status == 'In-Progress':
            total_workdays = proj_data.get('estimated_workdays', 0)
            if total_workdays and total_workdays > 0:
                days_passed = calculate_workdays(start_date_str, today.strftime('%Y-%m-%d'))
                progress = min(int((days_passed / total_workdays) * 100), 100)

        # --- Format Output String ---
        proj_id = proj_data.get('id', 'N/A')
        relative_path = os.path.relpath(proj_path, os.path.dirname(people_file))
        output_line = "[{id} {title} {project} {progress}% {workdays}d]({path})".format(
            id=proj_id, title=proj_data.get('title', 'N/A'), project=proj_data.get('project', ''),
            progress=progress, workdays=proj_data.get('estimated_workdays', 0),
            path=relative_path.replace(os.path.sep, '/')
        )
        
        output_item = f"- {output_line}"
        if proj_status == 'In-Progress':
            report_content = proj_data.get('progress_report', '')
            if report_content:
                indented_report = "\n".join([f"    {line}" for line in report_content.split('\n') if line.strip()])
                output_item += f"\n{indented_report}"
        
        category = status_map[proj_status]
        categorized_projects[category].append(output_item)

    # Sort and slice completed projects for Gantt
    completed_projects_for_gantt.sort(key=lambda x: x['date'], reverse=True)
    gantt_sections["Completed"] = [p['task'] for p in completed_projects_for_gantt[:3]]

    # Sort projects within each category by filename in descending order
    for category in categorized_projects:
        # Extract the filename from the markdown link path for sorting
        categorized_projects[category].sort(key=lambda item: os.path.basename(re.search(r'\((.*?)\)', item).group(1)), reverse=True)

    # --- Build the markdown block ---
    update_block = "<!-- AUTO_UPDATE_START -->\n"

    if any(gantt_sections.values()):
        update_block += "```mermaid\n"
        update_block += "gantt\n"
        update_block += "    dateFormat  YYYY-MM-DD\n"
        update_block += "    axisFormat %Y-%m\n"
        update_block += "    title Projects Overview\n"
        
        for section_name, tasks in gantt_sections.items():
            if tasks:
                update_block += f"\n    section {section_name}\n"
                update_block += "\n".join(tasks)
        update_block += "\n```\n\n"

    for category, items in categorized_projects.items():
        update_block += f"### {category}\n"
        if items:
            for item in items:
                update_block += f"{item}\n"
        else:
            update_block += "- 無\n"
        update_block += "\n"
    update_block += "<!-- AUTO_UPDATE_END -->"

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
