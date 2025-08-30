#!/usr/bin/env python3

import requests
import pdfplumber
from pathlib import Path
from io import BytesIO
from ics import Calendar, Event
from datetime import datetime, timedelta
import re
from typing import List, Tuple, Dict, Optional


COURSE_SCHEDULES = {
    "1A": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/1A2025-1.pdf",
    "1B": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/1B2025-1.pdf", 
    "2A": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/2A2025-1.pdf",
    "2B": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/2B2025-1.pdf",
    "3A": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/3A2025-1.pdf",
    "3B": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/3B2025-1.pdf",
    "4A": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/4A2025-1.pdf",
    "4B": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/4B2025-1.pdf",
    "5AB": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/5AB2025-1.pdf",
    "6AB": "https://ami.lnu.edu.ua/wp-content/uploads/2025/08/6AB2025-1.pdf"
}


def get_user_input():
    print("\n=== Ukrainian University Schedule Parser ===")
    print("Available courses:")
    for course_code in COURSE_SCHEDULES.keys():
        print(f"  {course_code}")
    
    course = input("\nEnter course code (e.g., 4B): ").strip().upper()
    if course not in COURSE_SCHEDULES:
        print(f"Unknown course: {course}. Using 4B as default.")
        course = "4B"
    
    print("\nSemester options:")
    print("  1 - First semester (starts with numerator week)")
    print("  2 - Second semester (starts with denominator week)")
    
    semester = input("Enter semester (1 or 2): ").strip()
    if semester not in ["1", "2"]:
        print("Invalid semester. Using 1 as default.")
        semester = "1"
    
    start_date_str = input("Enter semester start date (YYYY-MM-DD, e.g., 2025-09-01): ").strip()
    try:
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    except ValueError:
        print("Invalid date format. Using 2025-09-01 as default.")
        start_date = datetime(2025, 9, 1)
    
    group = input("Enter specific group (e.g., ПМА-41) or press Enter for all groups: ").strip()
    
    return {
        'course': course,
        'semester': int(semester),
        'start_date': start_date,
        'group': group if group else None
    }


def is_numerator_week(date: datetime, semester_start: datetime, semester: int) -> bool:
    week_number = (date - semester_start).days // 7
    
    if semester == 1:
        return week_number % 2 == 0
    else:
        return week_number % 2 == 1


def extract_week_info(subject_text: str) -> Tuple[str, str]:
    original_text = subject_text
    subject_text = subject_text.lower()
    
    if 'числ' in subject_text or 'чис' in subject_text:
        cleaned = re.sub(r'[\(\[]?числ[^\)\]]*[\)\]]?', '', original_text, flags=re.IGNORECASE)
        return 'numerator', cleaned.strip()
    elif 'знам' in subject_text or 'зн' in subject_text:
        cleaned = re.sub(r'[\(\[]?знам[^\)\]]*[\)\]]?', '', original_text, flags=re.IGNORECASE)
        return 'denominator', cleaned.strip()
    else:
        return 'all', original_text


def download_pdf(url, local_path):
    try:
        print(f"Downloading from: {url}")
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        with open(local_path, 'wb') as f:
            f.write(response.content)
        print(f"Downloaded and saved to: {local_path}")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Failed to download from {url}: {e}")
        return False


def extract_text_from_pdf(pdf_path):
    text_content = []
    schedule_data = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"Processing page {page_num}")
                
                tables = page.extract_tables()
                if tables:
                    print(f"Found {len(tables)} tables on page {page_num}")
                    for table_num, table in enumerate(tables):
                        print(f"Processing table {table_num + 1}")
                        schedule_data.extend(parse_table_directly(table))
                
                text = page.extract_text()
                if text:
                    text_content.append(text)
        
        return "\n".join(text_content), schedule_data
        
    except Exception as e:
        print(f"Error extracting from PDF: {e}")
        return "", []


def parse_table_directly(table):
    schedule = []
    
    if not table or len(table) < 2:
        return schedule
    
    print("Table structure:")
    for i, row in enumerate(table[:5]):
        print(f"Row {i}: {row}")
    
    header_row = None
    for i, row in enumerate(table):
        if row and any(cell and ('ПМА' in str(cell) or 'ПМП' in str(cell) or 'ІН' in str(cell)) for cell in row if cell):
            header_row = i
            print(f"Found header row at index {i}: {row}")
            break
    
    if header_row is None:
        print("No header row found")
        return schedule
    
    groups = []
    group_columns = []
    for col_idx, cell in enumerate(table[header_row]):
        if cell and ('ПМА' in str(cell) or 'ПМП' in str(cell) or 'ІН' in str(cell)):
            groups.append(cell.strip())
            group_columns.append(col_idx)
    
    print(f"Found groups: {groups}")
    print(f"Group columns: {group_columns}")
    
    current_day = "Monday"
    days_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    current_day_index = 0
    rows_without_time = 0
    
    for row_idx in range(header_row + 1, len(table)):
        row = table[row_idx]
        if not row:
            continue
            
        if row[0] and len(str(row[0])) > 5:
            day_text = str(row[0]).lower().replace('\n', '').replace(' ', '')
            
            if 'колідено' in day_text or 'понеділок' in day_text or 'понеділ' in day_text:
                current_day = "Monday"
                current_day_index = 0
                print(f"Found day: {current_day}")
            elif 'вівторок' in day_text or 'кротвіво' in day_text:
                current_day = "Tuesday"
                current_day_index = 1
                print(f"Found day: {current_day}")
            elif 'середа' in day_text or 'адерес' in day_text:
                current_day = "Wednesday"
                current_day_index = 2
                print(f"Found day: {current_day}")
            elif 'четвер' in day_text or 'ревтеч' in day_text:
                current_day = "Thursday"
                current_day_index = 3
                print(f"Found day: {current_day}")
            elif 'п\'ятниця' in day_text or 'пятниця' in day_text or 'яцинтяп' in day_text:
                current_day = "Friday"
                current_day_index = 4
                print(f"Found day: {current_day}")
        
        time_info = ""
        if len(row) > 1 and row[1]:
            time_info = str(row[1])
            rows_without_time = 0
        else:
            rows_without_time += 1
            if rows_without_time > 5 and current_day_index < len(days_order) - 1:
                current_day_index += 1
                current_day = days_order[current_day_index]
                rows_without_time = 0
                print(f"Auto-progressed to day: {current_day}")
        
        time_matches = re.findall(r'(\d{3,4})\s*-\s*(\d{3,4})', time_info)
        
        if time_matches:
            for start_raw, end_raw in time_matches:
                start_time = format_time(start_raw)
                end_time = format_time(end_raw)
                
                for i, col_idx in enumerate(group_columns):
                    if col_idx < len(row) and row[col_idx]:
                        group = groups[i]
                        subject_text = extract_subject_from_cell(row[col_idx])
                        if subject_text:
                            week_type, cleaned_subject = extract_week_info(subject_text)
                            
                            schedule.append({
                                'day': current_day,
                                'start_time': start_time,
                                'end_time': end_time,
                                'group': group,
                                'subject': cleaned_subject,
                                'week_type': week_type,
                                'original_text': subject_text
                            })
                            print(f"Found: {current_day} {start_time}-{end_time} {group}: {cleaned_subject} ({week_type})")
        else:
            for i, col_idx in enumerate(group_columns):
                if col_idx < len(row) and row[col_idx]:
                    group = groups[i]
                    subject_text = extract_subject_from_cell(row[col_idx])
                    if subject_text:
                        week_type, cleaned_subject = extract_week_info(subject_text)
                        schedule.append({
                            'day': current_day,
                            'start_time': 'Unknown',
                            'end_time': 'Unknown',
                            'group': group,
                            'subject': cleaned_subject,
                            'week_type': week_type,
                            'original_text': subject_text
                        })
                        print(f"Found (no time): {current_day} {group}: {cleaned_subject} ({week_type})")
    
    return schedule


def format_time(time_str):
    time_mappings = {
        "830": "08:30", "950": "09:50",
        "1010": "10:10", "1130": "11:30", 
        "1150": "11:50", "1310": "13:10",
        "1330": "13:30", "1450": "14:50",
        "1505": "15:05", "1625": "16:25",
        "1640": "16:40", "1800": "18:00",
        "1810": "18:10", "1930": "19:30"
    }
    
    if time_str in time_mappings:
        return time_mappings[time_str]
    
    if len(time_str) == 4:
        return f"{time_str[:2]}:{time_str[2:]}"
    elif len(time_str) == 3:
        return f"0{time_str[0]}:{time_str[1:]}"
    
    return time_str


def extract_subject_from_cell(cell):
    if not cell:
        return None
    
    cell_text = str(cell).strip()
    if len(cell_text) < 3:
        return None
    
    cell_text = re.sub(r'\s+', ' ', cell_text)
    
    subject_patterns = [
        r'([А-ЯІЇЄҐ][А-ЯІЇЄҐа-яіїєґ\s]{5,40})',
        r'(ПРОГРАМУВАННЯ[^,]*)',
        r'(МАТЕМАТИЧ[^,]*)',
        r'(МЕТОДИ[^,]*)',
        r'(ОБЧИСЛЮВАЛЬНА[^,]*)'
    ]
    
    for pattern in subject_patterns:
        match = re.search(pattern, cell_text)
        if match:
            return match.group(1).strip()
    
    return cell_text[:100]


def parse_schedule_from_text(text):
    schedule = []
    
    lines = text.split('\n')
    
    day_pattern = r'(Понеділок|Вівторок|Середа|Четвер|П\'ятниця|Субота|Неділя)'
    time_pattern = r'(\d{3,4})\s*-?\s*(\d{3,4})'
    
    current_day = None
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
            
        day_match = re.search(day_pattern, line)
        if day_match:
            current_day = day_match.group(1)
            print(f"Found day: {current_day}")
            continue
        
        time_match = re.search(time_pattern, line)
        if time_match and current_day:
            start_raw = time_match.group(1)
            end_raw = time_match.group(2)
            
            start_time = format_time(start_raw)
            end_time = format_time(end_raw)
            
            subject_part = re.sub(time_pattern, '', line).strip()
            
            if subject_part:
                subject_part = re.sub(r'\s+', ' ', subject_part)
                
                week_type, cleaned_subject = extract_week_info(subject_part)
                
                if len(cleaned_subject) > 3:
                    schedule.append({
                        'day': current_day,
                        'start_time': start_time,
                        'end_time': end_time,
                        'group': 'Unknown',
                        'subject': cleaned_subject,
                        'week_type': week_type,
                        'original_text': subject_part
                    })
                    print(f"Added: {current_day} {start_time}-{end_time} {cleaned_subject} ({week_type})")
    
    return schedule


def create_ics_from_schedule(schedule, user_config, output_file="schedule_enhanced.ics"):
    cal = Calendar()
    
    semester_start = user_config['start_date']
    semester = user_config['semester']
    selected_group = user_config['group']
    
    for week in range(18):
        week_start_date = semester_start + timedelta(weeks=week)
        is_numerator = is_numerator_week(week_start_date, semester_start, semester)
        
        week_type_name = 'numerator' if is_numerator else 'denominator'
        print(f"Week {week + 1}: {week_start_date.strftime('%Y-%m-%d')} - {week_type_name}")
        
        weekday_mapping = {
            "Monday": 0, "Tuesday": 1, "Wednesday": 2, 
            "Thursday": 3, "Friday": 4, "Saturday": 5, "Sunday": 6
        }
        
        for entry in schedule:
            if selected_group:
                entry_group = entry['group'].replace(' ', '').replace('-', '').upper()
                selected_group_clean = selected_group.replace(' ', '').replace('-', '').upper()
                if selected_group_clean not in entry_group:
                    continue
            
            entry_week_type = entry['week_type']
            
            if entry_week_type == 'all':
                should_occur = True
            elif entry_week_type == 'numerator':
                should_occur = is_numerator
            elif entry_week_type == 'denominator':
                should_occur = not is_numerator
            else:
                should_occur = True
            
            if not should_occur:
                continue
            
            if entry['start_time'] == 'Unknown' or entry['end_time'] == 'Unknown':
                continue
                
            try:
                start_hour, start_min = map(int, entry['start_time'].split(':'))
                end_hour, end_min = map(int, entry['end_time'].split(':'))
            except ValueError:
                continue
            
            day_name = entry['day']
            if day_name == 'Unknown' or day_name not in weekday_mapping:
                continue
                
            weekday = weekday_mapping[day_name]
            event_date = week_start_date + timedelta(days=weekday)
            
            event = Event()
            event.name = f"{entry['group']}: {entry['subject']}"
            event.begin = event_date.replace(hour=start_hour, minute=start_min)
            event.end = event_date.replace(hour=end_hour, minute=end_min)
            
            event.description = (
                f"Group: {entry['group']}\n"
                f"Subject: {entry['subject']}\n"
                f"Week type: {week_type_name}\n"
                f"Original: {entry['original_text']}"
            )
            
            cal.events.add(event)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.writelines(cal)
    
    print(f"Enhanced ICS file created: {output_file}")
    print(f"Generated events for {len(cal.events)} classes across 18 weeks")


def main():
    config = get_user_input()
    
    pdf_url = COURSE_SCHEDULES[config['course']]
    pdf_name = f"{config['course']}2025-1.pdf"
    
    print(f"\nProcessing course: {config['course']}")
    print(f"Semester: {config['semester']}")
    print(f"Start date: {config['start_date'].strftime('%Y-%m-%d')}")
    print(f"Selected group: {config['group'] or 'All groups'}")
    
    schedules_dir = Path(__file__).resolve().parent.parent / "schedules"
    schedules_dir.mkdir(parents=True, exist_ok=True)
    local_pdf_path = schedules_dir / pdf_name
    
    if not local_pdf_path.exists():
        if not download_pdf(pdf_url, local_pdf_path):
            print("Failed to download PDF file")
            return
    else:
        print(f"Using local file: {local_pdf_path}")
    
    print("Extracting data from PDF...")
    text, table_schedule = extract_text_from_pdf(local_pdf_path)
    
    text_file = Path(f"extracted_text_{config['course']}.txt")
    with open(text_file, 'w', encoding='utf-8') as f:
        f.write(text)
    print(f"Extracted text saved to: {text_file}")
    
    schedule = table_schedule if table_schedule else parse_schedule_from_text(text)
    
    if not schedule:
        print("No schedule data found")
        print("Please check the extracted text and customize the parsing logic")
        return
    
    print("Creating enhanced ICS file with week alternation...")
    output_file = f"schedule_{config['course']}_{config['group'] or 'all'}.ics"
    create_ics_from_schedule(schedule, config, output_file)
    
    print(f"\nFound {len(schedule)} unique schedule entries:")
    for entry in schedule[:10]:
        print(f"  {entry['day']} {entry['start_time']}-{entry['end_time']} "
              f"{entry['group']}: {entry['subject']} ({entry['week_type']})")
    
    if len(schedule) > 10:
        print(f"  ... and {len(schedule) - 10} more entries")
    
    print(f"\n✅ SUCCESS: ICS calendar file created: {output_file}")
    print("📱 You can now import this file to your phone calendar!")
    print("💡 The file contains all classes with proper week alternation (numerator/denominator)")
    
    if config['group']:
        print(f"📋 Filtered for group: {config['group']}")
    else:
        print("📋 Includes all groups - import and filter in your calendar app")


if __name__ == "__main__":
    main()
