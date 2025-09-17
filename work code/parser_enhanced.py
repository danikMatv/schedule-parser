#!/usr/bin/env python3

import requests
import pdfplumber
from pathlib import Path
from ics import Calendar, Event
from datetime import datetime, timedelta
import re
from typing import List, Tuple, Dict, Optional
import json


COURSE_SCHEDULES = {
    "1A": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/1B2025-1s.pdf",
    "1B": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/1A2025-1s.pdf", 
    "2A": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/2A2025-1s.pdf",
    "2B": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/2B2025-1s.pdf",
    "3A": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/3A2025-1s.pdf",
    "3B": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/3B2025-1s.pdf",
    "4A": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/4A2025-1s.pdf",
    "4B": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/4B2025-1s.pdf",
    "5AB": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/5AB2025-1s.pdf",
    "6AB": "https://ami.lnu.edu.ua/wp-content/uploads/2025/09/6AB2025-1s.pdf"
}

TIME_SLOTS = {
    "I": ("08:30", "09:50"),
    "II": ("10:10", "11:30"),
    "III": ("11:50", "13:10"),
    "IV": ("13:30", "14:50"),
    "V": ("15:05", "16:25"),
    "VI": ("16:40", "18:00"),
    "VII": ("18:10", "19:30"),
    "VIII": ("19:40", "21:00"),
    "І": ("08:30", "09:50"),
    "ІІ": ("10:10", "11:30"),
    "ІІІ": ("11:50", "13:10"),
    "ІV": ("13:30", "14:50"),
}

DAYS_UA = {
    "понеділок": "Monday",
    "вівторок": "Tuesday", 
    "середа": "Wednesday",
    "четвер": "Thursday",
    "п'ятниця": "Friday",
    "пʼятниця": "Friday",
    "субота": "Saturday"
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


def extract_week_info(text: str) -> Tuple[str, str]:
    if not text:
        return 'all', text
        
    original_text = text
    text_lower = text.lower()
    
    # Add more Ukrainian patterns
    numerator_patterns = ['числ', 'чис', '(ч)', 'numerator', 'непарн', 'нп']
    denominator_patterns = ['знам', 'зн', '(з)', 'denominator', 'парн', 'пр']
    
    # Look for patterns at word boundaries
    if re.search(r'\b(числ|чис|непарн|нп)\b', text_lower):
        cleaned = re.sub(r'[\(\[]?(числ|чис|непарн|нп)[^\)\]]*[\)\]]?', '', original_text, flags=re.IGNORECASE)
        return 'numerator', cleaned.strip()
    
    if re.search(r'\b(знам|зн|парн|пр)\b', text_lower):
        cleaned = re.sub(r'[\(\[]?(знам|зн|парн|пр)[^\)\]]*[\)\]]?', '', original_text, flags=re.IGNORECASE)
        return 'denominator', cleaned.strip()
    
    return 'all', original_text


def parse_improved_table(table, debug=False):
    schedule = []
    
    if not table or len(table) < 2:
        if debug:
            print("Table is too small or empty")
        return schedule
    
    if debug:
        print(f"Table has {len(table)} rows")
        for i, row in enumerate(table[:3]):  # Print first 3 rows
            print(f"Row {i}: {row}")

    header_row_idx = None
    groups = []
    group_columns = {}
    
    for i, row in enumerate(table[:10]):
        if row and any(cell and ('ПМ' in str(cell) or 'ІН' in str(cell)) for cell in row if cell):
            header_row_idx = i
            for j, cell in enumerate(row):
                if cell and ('ПМ' in str(cell) or 'ІН' in str(cell)):
                    group_name = str(cell).strip()
                    groups.append(group_name)
                    group_columns[group_name] = j
            break
    
    if not groups:
        print("No groups found in table")
        return schedule
    
    if debug:
        print(f"Found groups: {groups}")
        print(f"Group columns: {group_columns}")
    
    current_day = None
    current_time_slot = None
    
    for row_idx in range(header_row_idx + 1 if header_row_idx else 1, len(table)):
        row = table[row_idx]
        if not row:
            continue
        
        # Improved day detection
        if row[0]:
            day_text = str(row[0]).lower().replace('\n', '').replace(' ', '')
            if debug:
                print(f"Checking for day in: '{day_text}'")
            
            # Handle vertical text - reverse the string to get correct order
            reversed_day = day_text[::-1]
            if debug:
                print(f"Reversed day text: '{reversed_day}'")
                
            # Check for standard day names in both normal and reversed form
            day_patterns = {
                'понеділок': 'Monday',
                'поне': 'Monday',  # partial match
                'вівторок': 'Tuesday', 
                'вівт': 'Tuesday',  # partial match
                'середа': 'Wednesday',
                'сере': 'Wednesday',  # partial match
                'четвер': 'Thursday',
                'четв': 'Thursday',  # partial match
                'п\'ятниця': 'Friday',
                'пятн': 'Friday',  # partial match
                'субота': 'Saturday',
                'субо': 'Saturday'  # partial match
            }
            
            # Check both original and reversed text
            for pattern, en_day in day_patterns.items():
                if pattern in day_text or pattern in reversed_day:
                    current_day = en_day
                    if debug:
                        print(f"Found day: {current_day} (matched '{pattern}')")
                    break
            
            # Special hardcoded mappings for your specific vertical format
            day_mappings = {
                'коліденоп': 'Monday',    # понеділок reversed
                'коротвів': 'Tuesday',     # вівторок reversed  
                'адерес': 'Wednesday',     # середа reversed
                'ревтеч': 'Thursday',      # четвер reversed
                'яцинтя\'п': 'Friday',    # п'ятниця reversed
                'атобус': 'Saturday'       # субота reversed
            }
            
            if day_text in day_mappings:
                current_day = day_mappings[day_text]
                if debug:
                    print(f"Found day using mapping: {current_day}")
        
        time_found = False
        for col_idx in range(min(2, len(row))):
            if row[col_idx]:
                cell_text = str(row[col_idx]).strip()
                if debug:
                    print(f"Checking time slot: '{cell_text}'")
                if cell_text in TIME_SLOTS:
                    current_time_slot = cell_text
                    time_found = True
                    if debug:
                        print(f"Found time slot: {current_time_slot}")
                    break
                time_match = re.search(r'(\d{1,2})[:\s\-]?(\d{2})[\s\-]*[-–][\s\-]*(\d{1,2})[:\s\-]?(\d{2})', cell_text)
                if time_match:
                    start_time = f"{time_match.group(1).zfill(2)}:{time_match.group(2)}"
                    end_time = f"{time_match.group(3).zfill(2)}:{time_match.group(4)}"
                    time_found = True
                    break
        
        if current_day and (current_time_slot or time_found):
            for group_name, col_idx in group_columns.items():
                if col_idx < len(row) and row[col_idx]:
                    subject_text = str(row[col_idx]).strip()
                    
                    if len(subject_text) < 3:
                        continue
                    
                    week_type, cleaned_subject = extract_week_info(subject_text)
                    
                    if current_time_slot and current_time_slot in TIME_SLOTS:
                        start_time, end_time = TIME_SLOTS[current_time_slot]
                    elif not time_found:
                        continue
                    
                    room_match = re.search(r'(\d{2,3}[а-я]?)', cleaned_subject)
                    room = room_match.group(1) if room_match else ""
                    
                    schedule.append({
                        'day': current_day,
                        'start_time': start_time,
                        'end_time': end_time,
                        'group': group_name,
                        'subject': cleaned_subject,
                        'week_type': week_type,
                        'room': room,
                        'original_text': subject_text
                    })
                    
                    if debug:
                        print(f"Added: {current_day} {start_time}-{end_time} {group_name}: {cleaned_subject[:30]}... ({week_type})")
    
    return schedule


def extract_schedule_from_pdf(pdf_path, debug=False):
    if debug:
        print(f"Opening PDF: {pdf_path}")
        if not Path(pdf_path).exists():
            print(f"ERROR: PDF file does not exist at {pdf_path}")
            return []
    
    all_schedules = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                if debug:
                    print(f"\nProcessing page {page_num}")
                
                tables = page.extract_tables()
                if tables:
                    for table_num, table in enumerate(tables):
                        if debug:
                            print(f"Processing table {table_num + 1} with {len(table)} rows")
                        schedule_data = parse_improved_table(table, debug=debug)
                        all_schedules.extend(schedule_data)
                
                text = page.extract_text()
                if text and not tables:
                    if debug:
                        print("No tables found, trying text parsing")
        
        return all_schedules
        
    except Exception as e:
        print(f"Error extracting from PDF: {e}")
        return []


def create_ics_from_schedule(schedule, user_config, output_file="schedule.ics"):
    cal = Calendar()
    
    semester_start = user_config['start_date']
    semester = user_config['semester']
    selected_group = user_config['group']
    
    for week in range(18):
        week_start_date = semester_start + timedelta(weeks=week)
        is_numerator = is_numerator_week(week_start_date, semester_start, semester)
        
        week_type_name = 'numerator' if is_numerator else 'denominator'
        
        weekday_mapping = {
            "Monday": 0, "Tuesday": 1, "Wednesday": 2, 
            "Thursday": 3, "Friday": 4, "Saturday": 5
        }
        
        for entry in schedule:
            if selected_group:
                if selected_group.upper() not in entry['group'].upper():
                    continue
            
            entry_week_type = entry['week_type']
            if entry_week_type == 'numerator' and not is_numerator:
                continue
            elif entry_week_type == 'denominator' and is_numerator:
                continue
            
            if entry['day'] in weekday_mapping:
                weekday = weekday_mapping[entry['day']]
                event_date = week_start_date + timedelta(days=weekday)
                
                try:
                    start_hour, start_min = map(int, entry['start_time'].split(':'))
                    end_hour, end_min = map(int, entry['end_time'].split(':'))
                except ValueError:
                    continue
                
                event = Event()
                event.name = f"{entry['group']}: {entry['subject'][:50]}"
                event.begin = event_date.replace(hour=start_hour, minute=start_min)
                event.end = event_date.replace(hour=end_hour, minute=end_min)
                
                if entry.get('room'):
                    event.location = f"Room {entry['room']}"
                
                event.description = (
                    f"Group: {entry['group']}\n"
                    f"Subject: {entry['subject']}\n"
                    f"Week type: {week_type_name}\n"
                    f"Original: {entry['original_text']}"
                )
                
                cal.events.add(event)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.writelines(cal)
    
    print(f"\n✅ ICS file created: {output_file}")
    print(f"📅 Generated {len(cal.events)} events across 18 weeks")
    
    return cal


def download_pdf(url, local_path):
    try:
        print(f"Downloading from: {url}")
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        with open(local_path, 'wb') as f:
            f.write(response.content)
        print(f"Downloaded to: {local_path}")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Failed to download: {e}")
        return False


def main():
    config = get_user_input()
    
    pdf_url = COURSE_SCHEDULES[config['course']]
    pdf_name = f"{config['course']}2025-1.pdf"
    
    print(f"\n📚 Processing course: {config['course']}")
    print(f"📆 Semester: {config['semester']}")
    print(f"📅 Start date: {config['start_date'].strftime('%Y-%m-%d')}")
    print(f"👥 Selected group: {config['group'] or 'All groups'}")
    
    schedules_dir = Path("schedules")
    schedules_dir.mkdir(exist_ok=True)
    local_pdf_path = schedules_dir / pdf_name
    
    if not local_pdf_path.exists():
        if not download_pdf(pdf_url, local_pdf_path):
            print("Failed to download PDF")
            return
    else:
        print(f"Using existing file: {local_pdf_path}")
    
    print("\nExtracting schedule from PDF...")
    schedule = extract_schedule_from_pdf(local_pdf_path, debug=True)
    
    if not schedule:
        print(" No schedule data found")
        print("This might be due to PDF structure. Try manual input or contact support.")
        return
    
    print(f"\n📊 Found {len(schedule)} schedule entries")
    
    json_file = schedules_dir / f"{config['course']}_schedule.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(schedule, f, ensure_ascii=False, indent=2)
    print(f"💾 Saved schedule data to: {json_file}")
    
    print("\n📋 Sample schedule entries:")
    for entry in schedule[:5]:
        print(f"  {entry['day']} {entry['start_time']}-{entry['end_time']} "
              f"{entry['group']}: {entry['subject'][:30]}... ({entry['week_type']})")
    
    output_file = f"schedule_{config['course']}_{config['group'] or 'all'}.ics"
    create_ics_from_schedule(schedule, config, output_file)
    
    # Print summary info
    if not config['group']:
        print("3. Includes all groups - you can filter in your calendar app")
    else:
        print(f"3. Filtered for group: {config['group']}")
    
    print(f"\nCalendar file is ready: {output_file}")

if __name__ == "__main__":
    main()