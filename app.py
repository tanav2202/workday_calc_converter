import gradio as gr
import pandas as pd
import re
from datetime import datetime, timedelta
from icalendar import Calendar, Event
import uuid
import pytz
import openpyxl
from io import BytesIO
import tempfile
import os

def excel_date_to_datetime(excel_date):
    """Convert Excel serial date to Python datetime."""
    if isinstance(excel_date, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=excel_date)
    return excel_date

def parse_meeting_pattern(meeting_pattern):
    """Parse the meeting pattern string to extract schedule information."""
    if not meeting_pattern or pd.isna(meeting_pattern):
        return []
    
    patterns = meeting_pattern.split('\n\n')
    schedules = []
    
    for pattern in patterns:
        pattern = pattern.strip()
        if not pattern:
            continue
            
        parts = pattern.split(' | ')
        
        if len(parts) >= 6:
            date_range = parts[0].strip()
            days = parts[1].strip()
            time_range = parts[2].strip()
            campus = parts[3].strip()
            building = parts[4].strip()
            location_parts = parts[5:]
            
            try:
                start_date_str, end_date_str = date_range.split(' - ')
                start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
                end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
            except Exception:
                continue
            
            try:
                time_parts = time_range.split(' - ')
                start_time_str = time_parts[0].strip()
                end_time_str = time_parts[1].strip()
                
                def normalize_time_string(time_str):
                    time_str = time_str.replace('a.m.', 'AM').replace('p.m.', 'PM')
                    time_str = time_str.replace('a.m', 'AM').replace('p.m', 'PM')
                    return time_str
                
                start_time_str = normalize_time_string(start_time_str)
                end_time_str = normalize_time_string(end_time_str)
                
                start_time = datetime.strptime(start_time_str, '%I:%M %p').time()
                end_time = datetime.strptime(end_time_str, '%I:%M %p').time()
            except Exception:
                continue
            
            day_mapping = {
                'Mon': 0, 'Tue': 1, 'Wed': 2, 'Thu': 3, 'Fri': 4, 'Sat': 5, 'Sun': 6
            }
            
            weekdays = []
            for day_name in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']:
                if day_name in days:
                    weekdays.append(day_mapping[day_name])
            
            location = f"{building}, {' | '.join(location_parts)}, {campus}".strip(', ')
            
            if weekdays:
                schedules.append({
                    'start_date': start_date,
                    'end_date': end_date,
                    'start_time': start_time,
                    'end_time': end_time,
                    'weekdays': weekdays,
                    'location': location
                })
    
    return schedules

def read_course_data_from_excel(excel_file_path):
    """Read course data from Excel file."""
    workbook = openpyxl.load_workbook(excel_file_path)
    worksheet = workbook['View Courses for Student']
    
    headers = []
    for col in range(1, 15):
        cell_value = worksheet.cell(row=3, column=col).value
        headers.append(cell_value if cell_value else '')
    
    courses = []
    for row_num in range(4, 100):
        row_data = []
        has_data = False
        
        for col in range(1, len(headers) + 1):
            cell_value = worksheet.cell(row=row_num, column=col).value
            if cell_value is not None:
                row_data.append(cell_value)
                has_data = True
            else:
                row_data.append('')
        
        if has_data and len(row_data) > 1 and row_data[1]:
            courses.append(row_data)
    
    df = pd.DataFrame(courses, columns=headers)
    df.columns = [col.strip() if col else f'Column_{i}' for i, col in enumerate(df.columns)]
    
    return df

def create_ics_from_excel_df(df):
    """Convert DataFrame to ICS format."""
    if df.empty:
        return None, "No course data found!"
    
    cal = Calendar()
    cal.add('prodid', '-//Course Schedule//Course Schedule//EN')
    cal.add('version', '2.0')
    cal.add('calscale', 'GREGORIAN')
    cal.add('method', 'PUBLISH')
    cal.add('x-wr-calname', 'Course Schedule')
    cal.add('x-wr-timezone', 'America/Vancouver')
    
    vancouver_tz = pytz.timezone('America/Vancouver')
    events_created = 0
    
    for index, row in df.iterrows():
        course_listing = str(row.get('Course Listing', '')).strip()
        meeting_patterns = str(row.get('Meeting Patterns', '')).strip()
        
        if not course_listing or not meeting_patterns or meeting_patterns == 'nan':
            continue
            
        section = str(row.get('Section', '')).strip()
        instructor = str(row.get('Instructor', '')).strip()
        instructional_format = str(row.get('Instructional Format', '')).strip()
        
        schedules = parse_meeting_pattern(meeting_patterns)
        
        for schedule in schedules:
            current_date = schedule['start_date']
            while current_date <= schedule['end_date']:
                if current_date.weekday() in schedule['weekdays']:
                    event = Event()
                    
                    start_datetime = datetime.combine(current_date, schedule['start_time'])
                    end_datetime = datetime.combine(current_date, schedule['end_time'])
                    
                    start_datetime = vancouver_tz.localize(start_datetime)
                    end_datetime = vancouver_tz.localize(end_datetime)
                    
                    event.add('uid', str(uuid.uuid4()))
                    event.add('dtstart', start_datetime)
                    event.add('dtend', end_datetime)
                    event.add('dtstamp', datetime.now(vancouver_tz))
                    
                    summary = course_listing
                    if instructional_format:
                        summary += f" - {instructional_format}"
                    event.add('summary', summary)
                    
                    description_parts = []
                    if section:
                        description_parts.append(f"Section: {section}")
                    if instructor:
                        description_parts.append(f"Instructor: {instructor}")
                    if instructional_format:
                        description_parts.append(f"Format: {instructional_format}")
                    
                    if description_parts:
                        event.add('description', '\\n'.join(description_parts))
                    
                    event.add('location', schedule['location'])
                    event.add('categories', ['EDUCATION', 'COURSE'])
                    
                    cal.add_component(event)
                    events_created += 1
                
                current_date += timedelta(days=1)
    
    return cal.to_ical(), events_created

def process_excel_file(file):
    """Process uploaded Excel file."""
    if file is None:
        return None, "Please upload an Excel file."
    
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(file)
            tmp_file_path = tmp_file.name
        
        df = read_course_data_from_excel(tmp_file_path)
        
        if df.empty:
            return None, "No course data found."
        
        ical_content, events_created = create_ics_from_excel_df(df)
        
        if ical_content is None:
            return None, "Failed to create calendar events."
        
        output_path = tempfile.NamedTemporaryFile(delete=False, suffix='.ics', mode='wb')
        output_path.write(ical_content)
        output_path.close()
        
        os.unlink(tmp_file_path)
        
        success_message = f"✅ Success! Found {len(df)} courses, created {events_created} events."
        
        return output_path.name, success_message
        
    except Exception as e:
        return None, f"Error: {str(e)}"

def create_app():
    with gr.Blocks(title="Excel to ICS Converter") as demo:
        gr.HTML("<h1>📅 Excel to ICS Calendar Converter</h1>")
        
        with gr.Row():
            with gr.Column():
                gr.HTML("""
                <h3>How to Use:</h3>
                <ol>
                    <li>Upload your Excel course file</li>
                    <li>Click Convert</li>
                    <li>Download the ICS file</li>
                </ol>
                """)
            
            with gr.Column():
                file_input = gr.File(
                    label="Upload Excel File",
                    file_types=[".xlsx", ".xls"],
                    type="binary"
                )
                
                convert_btn = gr.Button("Convert to Calendar", variant="primary")
                
                status_output = gr.Markdown("Upload a file to start.")
                
                download_output = gr.File(label="Download ICS File", visible=False)
        
        def process_and_update(file):
            ics_file, message = process_excel_file(file)
            if ics_file:
                return message, gr.File(value=ics_file, visible=True)
            else:
                return message, gr.File(visible=False)
        
        convert_btn.click(
            fn=process_and_update,
            inputs=[file_input],
            outputs=[status_output, download_output]
        )
    
    return demo

app = create_app()

# For Vercel
def handler(request):
    return app.launch(share=False, server_name="0.0.0.0", server_port=8000)

