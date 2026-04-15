"""
Custom HTTP Server for Gantt Chart Generator
=============================================
Serves static files and handles PPTX generation via POST /api/generate-pptx
Output files are saved to ./output/ folder.

Usage:
    python server.py
    Open http://localhost:8090
"""

import http.server
import json
import os
import sys
import io
import re
import csv
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import unquote

# Fix Windows console encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "output"
PORT = 8090

# ========================================
# PPTX Color Config (Light Theme for PPT)
# ========================================
# Light theme colors - white/light background, dark text
BG_COLOR = RGBColor(0xff, 0xff, 0xff)         # White background
BG_SECONDARY = RGBColor(0xf0, 0xf2, 0xf5)     # Light gray header
BG_ROW_ALT = RGBColor(0xf8, 0xf9, 0xfb)       # Very light row alt
TEXT_PRIMARY = RGBColor(0x1a, 0x1a, 0x2e)       # Dark text
TEXT_SECONDARY = RGBColor(0x64, 0x74, 0x8b)     # Medium gray
TEXT_MUTED = RGBColor(0x94, 0xa3, 0xb8)         # Light gray
GRID_COLOR = RGBColor(0xe2, 0xe8, 0xf0)         # Light grid
TODAY_COLOR = RGBColor(0xeb, 0x57, 0x57)         # Red today line
TITLE_COLOR = RGBColor(0x1e, 0x3a, 0x5f)        # Dark blue title

TYPE_COLORS = {
    'Milestone':   {'rgb': RGBColor(0xf2, 0x99, 0x4a), 'label': 'MS'},
    'Goal':        {'rgb': RGBColor(0x4c, 0xaf, 0x50), 'label': 'Goal'},
    'Review':      {'rgb': RGBColor(0x29, 0x96, 0xd6), 'label': 'Rev'},
    'Event':       {'rgb': RGBColor(0x9c, 0x27, 0xb0), 'label': 'Evt'},
    'タスク':       {'rgb': RGBColor(0x42, 0x7a, 0xd6), 'label': 'Task'},
    'Task':        {'rgb': RGBColor(0x42, 0x7a, 0xd6), 'label': 'Task'},
    'Sub task':    {'rgb': RGBColor(0x64, 0x95, 0xed), 'label': 'Sub'},
    'Validation':  {'rgb': RGBColor(0xe5, 0x3e, 0x3e), 'label': 'Val'},
    'SOP':         {'rgb': RGBColor(0xe6, 0xa8, 0x17), 'label': 'SOP'},
    'Certificate': {'rgb': RGBColor(0x2f, 0x80, 0xed), 'label': 'Cert'},
    'PPAP':        {'rgb': RGBColor(0xe5, 0x4b, 0x4b), 'label': 'PPAP'},
    'Top-level initiative': {'rgb': RGBColor(0x3b, 0x5b, 0xdb), 'label': 'Init'},
}
DEFAULT_TYPE = {'rgb': RGBColor(0x82, 0x82, 0x82), 'label': '?'}
COMPLETED_COLOR = RGBColor(0x27, 0xae, 0x60)

# Layout constants
SLIDE_WIDTH = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)
LEFT_MARGIN = Inches(0.3)
TOP_MARGIN = Inches(0.6)
HEADER_HEIGHT = Inches(0.3)
ROW_HEIGHT = Inches(0.3)
LABEL_WIDTH = Inches(3.2)
CHART_LEFT = LEFT_MARGIN + LABEL_WIDTH + Inches(0.1)
CHART_WIDTH = SLIDE_WIDTH - CHART_LEFT - Inches(0.3)
TASKS_PER_SLIDE = 18


# ========================================
# Date Parsing
# ========================================
def parse_date(date_str):
    if not date_str or not date_str.strip():
        return None
    s = date_str.strip()
    # Remove time portions: "12:00 午前", "12:00 AM", etc.
    s = re.sub(r'\s+\d{1,2}:\d{2}(:\d{2})?\s*(午前|午後|AM|PM)?', '', s, flags=re.IGNORECASE).strip()
    
    # Format: DD/M/YY  (Jira: day first, 2-digit year last, e.g. "20/4/26" = Apr 20 2026)
    m = re.match(r'^(\d{1,4})/(\d{1,2})/(\d{1,4})$', s)
    if m:
        a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a > 31:
            # YYYY/M/D
            year, month, day = a, b, c
        elif c <= 99:
            # DD/M/YY — Jira format
            day, month, year = a, b, c + 2000
        else:
            # D/M/YYYY
            day, month, year = a, b, c
        try:
            return datetime(year, month, day)
        except ValueError:
            return None
    
    # Format: YYYY-MM-DD
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None



def parse_iso_date(s):
    """Parse ISO date string from JSON (YYYY-MM-DD or ISO 8601)."""
    if not s:
        return None
    try:
        # Handle "YYYY/MM/DD" and "YYYY-MM-DD"
        s = s.strip().replace('/', '-')
        # Remove time portion if present
        if 'T' in s:
            s = s.split('T')[0]
        parts = s.split('-')
        if len(parts) == 3:
            return datetime(int(parts[0]), int(parts[1]), int(parts[2]))
    except (ValueError, IndexError):
        pass
    return None


# ========================================
# Timeline Helpers
# ========================================
def days_in_month(dt):
    if dt.month == 12:
        return (datetime(dt.year + 1, 1, 1) - datetime(dt.year, dt.month, 1)).days
    return (datetime(dt.year, dt.month + 1, 1) - datetime(dt.year, dt.month, 1)).days


def get_months_between(start, end):
    months = []
    current = datetime(start.year, start.month, 1)
    while current < end:
        months.append(current)
        if current.month == 12:
            current = datetime(current.year + 1, 1, 1)
        else:
            current = datetime(current.year, current.month + 1, 1)
    return months


# ========================================
# Helpers
# ========================================
def add_text_box(slide, text, x, y, w, h,
                 font_size=7, color=TEXT_PRIMARY, bold=False,
                 align=PP_ALIGN.LEFT, font_name='Meiryo UI'):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    p.text = str(text)
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    tf._txBody.bodyPr.set('anchor', 'ctr')
    return txBox


# ========================================
# PPTX Generation from JSON data
# ========================================
def generate_pptx_from_data(tasks_data, output_filename, title="Gantt Chart", timeline_start=None, timeline_end=None):
    """Generate PPTX from task list sent by browser.
    timeline_start/timeline_end: optional custom bounds as date strings.
    """
    
    # Parse dates in task data
    tasks = []
    for t in tasks_data:
        start = parse_iso_date(t.get('startDate', ''))
        end = parse_iso_date(t.get('endDate', ''))
        if start or end:
            tasks.append({
                'key': t.get('key', ''),
                'type': t.get('type', 'Task'),
                'summary': t.get('summary', ''),
                'status': t.get('status', 'To Do'),
                'assignee': t.get('assignee', ''),
                'start_date': start,
                'end_date': end,
                'parent': t.get('parent', ''),
            })
    
    if not tasks:
        return None, "No tasks with dates found"
    
    # Sort & group by type
    type_order = ['Milestone', 'Top-level initiative', 'Goal', 'Review', 'Event',
                  'SOP', 'Certificate', 'Validation', 'PPAP', 'タスク', 'Task', 'Sub task']
    
    groups = {}
    for task in tasks:
        g = task['type']
        if g not in groups:
            groups[g] = []
        groups[g].append(task)
    
    def type_key(name):
        try: return type_order.index(name)
        except ValueError: return 999
    
    grouped = []
    for gname in sorted(groups.keys(), key=type_key):
        items = sorted(groups[gname], key=lambda t: t['start_date'] or datetime(9999, 1, 1))
        grouped.append({'is_group': True, 'label': gname, 'count': len(items)})
        grouped.extend(items)
    
    # Timeline bounds - use custom if provided, otherwise auto-detect
    custom_start = parse_iso_date(timeline_start) if timeline_start else None
    custom_end = parse_iso_date(timeline_end) if timeline_end else None
    
    all_dates = []
    for t in tasks:
        if t['start_date']: all_dates.append(t['start_date'])
        if t['end_date']: all_dates.append(t['end_date'])
    
    earliest = min(all_dates)
    latest = max(all_dates)
    
    if custom_start:
        bounds_start = datetime(custom_start.year, custom_start.month, 1)
    else:
        bounds_start = datetime(earliest.year, earliest.month, 1) - timedelta(days=30)
        bounds_start = datetime(bounds_start.year, bounds_start.month, 1)
    
    if custom_end:
        # End of the month of custom_end
        end_month = custom_end.month + 1
        end_year = custom_end.year
        if end_month > 12:
            end_month = 1
            end_year += 1
        bounds_end = datetime(end_year, end_month, 1)
    else:
        end_month = latest.month + 2
        end_year = latest.year
        if end_month > 12:
            end_month -= 12
            end_year += 1
        bounds_end = datetime(end_year, end_month, 1)
    
    total_days = max((bounds_end - bounds_start).days, 1)
    months = get_months_between(bounds_start, bounds_end)
    today = datetime.now()
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    
    # Paginate
    pages = []
    current_page = []
    task_count = 0
    for item in grouped:
        current_page.append(item)
        if not item.get('is_group'):
            task_count += 1
            if task_count >= TASKS_PER_SLIDE:
                pages.append(current_page)
                current_page = []
                task_count = 0
    if current_page:
        pages.append(current_page)
    
    for page_idx, page_items in enumerate(pages):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # White background
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_COLOR
        
        # Title
        page_label = f" ({page_idx+1}/{len(pages)})" if len(pages) > 1 else ""
        add_text_box(slide, f"{title}{page_label}",
                     LEFT_MARGIN, Inches(0.12), Inches(10), Inches(0.4),
                     font_size=16, color=TITLE_COLOR, bold=True)
        
        # Date subtitle
        add_text_box(slide, f"{bounds_start.strftime('%Y/%m')} - {bounds_end.strftime('%Y/%m')}",
                     LEFT_MARGIN, Inches(0.42), Inches(5), Inches(0.18),
                     font_size=7, color=TEXT_MUTED)
        
        # Month headers
        chart_width_emu = CHART_WIDTH
        month_x = CHART_LEFT
        
        for month_dt in months:
            days = days_in_month(month_dt)
            month_frac = days / total_days
            month_width = int(chart_width_emu * month_frac)
            
            shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, month_x, TOP_MARGIN, month_width, HEADER_HEIGHT
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = BG_SECONDARY
            shape.line.color.rgb = GRID_COLOR
            shape.line.width = Pt(0.5)
            
            tf = shape.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            month_names_jp = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
            p.text = f"{month_dt.year}/{month_names_jp[month_dt.month - 1]}"
            p.font.size = Pt(7)
            p.font.color.rgb = TEXT_SECONDARY
            p.font.name = 'Meiryo UI'
            p.alignment = PP_ALIGN.CENTER
            tf._txBody.bodyPr.set('anchor', 'ctr')
            
            # Grid line
            gl = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, month_x, TOP_MARGIN + HEADER_HEIGHT,
                Pt(0.5), Inches(len(page_items) * 0.3 + 0.1)
            )
            gl.fill.solid()
            gl.fill.fore_color.rgb = GRID_COLOR
            gl.line.fill.background()
            
            month_x += month_width
        
        # Today line
        today_days = (today - bounds_start).days
        if 0 <= today_days <= total_days:
            today_frac = today_days / total_days
            today_x = int(CHART_LEFT + chart_width_emu * today_frac)
            
            tl = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, today_x, TOP_MARGIN - Inches(0.05),
                Pt(2), Inches(len(page_items) * 0.3 + 0.5)
            )
            tl.fill.solid()
            tl.fill.fore_color.rgb = TODAY_COLOR
            tl.line.fill.background()
            
            add_text_box(slide, 'Today',
                         today_x - Inches(0.2), TOP_MARGIN - Inches(0.2),
                         Inches(0.4), Inches(0.15),
                         font_size=5, color=TODAY_COLOR, align=PP_ALIGN.CENTER)
        
        # Task rows
        row_y = TOP_MARGIN + HEADER_HEIGHT + Inches(0.05)
        row_idx = 0
        
        for item in page_items:
            y = row_y + int(row_idx * ROW_HEIGHT)
            
            if item.get('is_group'):
                bg_shape = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, LEFT_MARGIN, y,
                    SLIDE_WIDTH - LEFT_MARGIN * 2, ROW_HEIGHT
                )
                bg_shape.fill.solid()
                bg_shape.fill.fore_color.rgb = RGBColor(0xec, 0xef, 0xf4)
                bg_shape.line.fill.background()
                
                tc = TYPE_COLORS.get(item['label'], DEFAULT_TYPE)
                add_text_box(slide, f"  {item['label']}  ({item['count']})",
                             LEFT_MARGIN + Inches(0.05), y,
                             LABEL_WIDTH, ROW_HEIGHT,
                             font_size=7, color=tc['rgb'], bold=True)
                row_idx += 1
                continue
            
            # Alternating rows
            if row_idx % 2 == 0:
                bg_shape = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, LEFT_MARGIN, y,
                    SLIDE_WIDTH - LEFT_MARGIN * 2, ROW_HEIGHT
                )
                bg_shape.fill.solid()
                bg_shape.fill.fore_color.rgb = BG_ROW_ALT
                bg_shape.line.fill.background()
            
            tc = TYPE_COLORS.get(item['type'], DEFAULT_TYPE)
            
            # Key
            add_text_box(slide, item['key'],
                         LEFT_MARGIN + Inches(0.05), y,
                         Inches(0.7), ROW_HEIGHT,
                         font_size=6, color=TEXT_MUTED)
            
            # Type badge
            badge_w = Inches(0.4)
            badge_h = Inches(0.18)
            badge_y = y + (ROW_HEIGHT - badge_h) // 2
            badge = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                LEFT_MARGIN + Inches(0.75), badge_y, badge_w, badge_h
            )
            badge.fill.solid()
            r, g, b = tc['rgb'][0], tc['rgb'][1], tc['rgb'][2]
            badge.fill.fore_color.rgb = RGBColor(
                min(255, int(r * 0.15 + 255 * 0.85)),
                min(255, int(g * 0.15 + 255 * 0.85)),
                min(255, int(b * 0.15 + 255 * 0.85))
            )
            badge.line.fill.background()
            tf = badge.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            p.text = tc['label']
            p.font.size = Pt(5)
            p.font.color.rgb = tc['rgb']
            p.font.bold = True
            p.font.name = 'Meiryo UI'
            p.alignment = PP_ALIGN.CENTER
            tf._txBody.bodyPr.set('anchor', 'ctr')
            
            # Summary
            add_text_box(slide, item['summary'],
                         LEFT_MARGIN + Inches(1.2), y,
                         LABEL_WIDTH - Inches(1.2), ROW_HEIGHT,
                         font_size=7, color=TEXT_PRIMARY)
            
            # Bar
            if item['start_date'] or item['end_date']:
                start = item['start_date'] or item['end_date']
                end = item['end_date'] or item['start_date']
                
                start_days = (start - bounds_start).days
                duration_days = max((end - start).days, 1)
                
                bar_x = int(CHART_LEFT + chart_width_emu * (start_days / total_days))
                bar_w = max(int(chart_width_emu * (duration_days / total_days)), Inches(0.06))
                bar_h = Inches(0.18)
                bar_y = y + (ROW_HEIGHT - bar_h) // 2
                
                is_completed = item['status'] in ('完了', 'Done')
                bar_color = COMPLETED_COLOR if is_completed else tc['rgb']
                
                if item['type'] == 'Milestone':
                    diamond = slide.shapes.add_shape(
                        MSO_SHAPE.DIAMOND,
                        bar_x - Inches(0.075), bar_y,
                        Inches(0.15), bar_h
                    )
                    diamond.fill.solid()
                    diamond.fill.fore_color.rgb = bar_color
                    diamond.line.fill.background()
                else:
                    bar = slide.shapes.add_shape(
                        MSO_SHAPE.ROUNDED_RECTANGLE,
                        bar_x, bar_y, bar_w, bar_h
                    )
                    bar.fill.solid()
                    bar.fill.fore_color.rgb = bar_color
                    bar.line.fill.background()
                    
                    if bar_w > Inches(0.8):
                        tf = bar.text_frame
                        tf.word_wrap = False
                        p = tf.paragraphs[0]
                        p.text = item['summary']
                        p.font.size = Pt(5)
                        p.font.color.rgb = RGBColor(0xff, 0xff, 0xff)
                        p.font.name = 'Meiryo UI'
                        p.alignment = PP_ALIGN.LEFT
                        tf._txBody.bodyPr.set('anchor', 'ctr')
                    
                    # Start date label
                    if start_days >= 0:
                        add_text_box(slide, start.strftime('%m/%d'),
                                     bar_x - Inches(0.35), bar_y,
                                     Inches(0.33), bar_h,
                                     font_size=4, color=TEXT_MUTED, align=PP_ALIGN.RIGHT)
                    
                    if duration_days > 1:
                        add_text_box(slide, end.strftime('%m/%d'),
                                     bar_x + bar_w + Inches(0.02), bar_y,
                                     Inches(0.33), bar_h,
                                     font_size=4, color=TEXT_MUTED, align=PP_ALIGN.LEFT)
            
            row_idx += 1
        
        # Legend
        legend_y = SLIDE_HEIGHT - Inches(0.35)
        legend_x = LEFT_MARGIN
        
        for label, color in [
            ('Milestone', TYPE_COLORS['Milestone']['rgb']),
            ('Goal', TYPE_COLORS['Goal']['rgb']),
            ('Review', TYPE_COLORS['Review']['rgb']),
            ('Event', TYPE_COLORS['Event']['rgb']),
            ('Task', TYPE_COLORS['Task']['rgb']),
            ('Validation', TYPE_COLORS['Validation']['rgb']),
            ('Completed', COMPLETED_COLOR),
        ]:
            sw = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                legend_x, legend_y + Inches(0.03), Inches(0.2), Inches(0.1)
            )
            sw.fill.solid()
            sw.fill.fore_color.rgb = color
            sw.line.fill.background()
            
            add_text_box(slide, label,
                         legend_x + Inches(0.22), legend_y,
                         Inches(0.6), Inches(0.15),
                         font_size=6, color=TEXT_SECONDARY)
            legend_x += Inches(0.9)
        
        # Today legend
        ts = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, legend_x, legend_y, Pt(2), Inches(0.15))
        ts.fill.solid()
        ts.fill.fore_color.rgb = TODAY_COLOR
        ts.line.fill.background()
        add_text_box(slide, 'Today', legend_x + Inches(0.05), legend_y, Inches(0.5), Inches(0.15),
                     font_size=6, color=TODAY_COLOR)
    
    # Save
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / output_filename
    prs.save(str(output_path))
    return str(output_path.resolve()), None


# ========================================
# HTTP Handler
# ========================================
class GanttHandler(http.server.SimpleHTTPRequestHandler):
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(SCRIPT_DIR), **kwargs)
    
    def do_POST(self):
        if self.path == '/api/generate-pptx':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            
            try:
                data = json.loads(body.decode('utf-8'))
                tasks = data.get('tasks', [])
                title = data.get('title', 'Gantt Chart')
                filename = data.get('filename', f'gantt_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pptx')
                
                # Sanitize filename
                filename = re.sub(r'[^\w\-_\.]', '_', filename)
                if not filename.endswith('.pptx'):
                    filename += '.pptx'
                
                # Extract custom timeline bounds if provided
                tl_start = data.get('timelineStart', None)
                tl_end = data.get('timelineEnd', None)
                
                output_path, error = generate_pptx_from_data(
                    tasks, filename, title,
                    timeline_start=tl_start,
                    timeline_end=tl_end
                )
                
                if error:
                    self.send_response(400)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'error': error}).encode('utf-8'))
                else:
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    result = {
                        'success': True,
                        'path': output_path,
                        'filename': filename,
                    }
                    self.wfile.write(json.dumps(result).encode('utf-8'))
                    print(f"  [OK] Generated: {output_path}")
                    
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode('utf-8'))
                print(f"  [ERROR] {e}")
        else:
            self.send_response(404)
            self.end_headers()
    
    def log_message(self, format, *args):
        # Quieter logging
        if '/api/' in str(args[0]) if args else False:
            print(f"  [API] {args[0]}")
        else:
            super().log_message(format, *args)


# ========================================
# Main
# ========================================
if __name__ == '__main__':
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    print("=" * 50)
    print("  Gantt Chart Generator Server")
    print("=" * 50)
    print(f"  URL:    http://localhost:{PORT}")
    print(f"  Output: {OUTPUT_DIR.resolve()}")
    print(f"  Press Ctrl+C to stop")
    print("=" * 50)
    
    server = http.server.HTTPServer(('', PORT), GanttHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Server stopped.")
        server.server_close()
