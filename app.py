# Copyright (C) 2025 CHAIB YAHIA
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

# ================== الجزء الأول: استيراد المكتبات الأساسية ==================
from flask import Flask, jsonify, request, render_template, send_file
import os
import json
import sys
import random
import re
from collections import defaultdict, Counter, deque
import io
from datetime import datetime
import math
import copy
import webbrowser
from waitress import serve
from threading import Timer
from ortools.sat.python import cp_model
import queue
from flask import stream_with_context, Response
from concurrent.futures import ThreadPoolExecutor
import signal
import threading
import time
import uuid
import sqlite3

# --- إضافة جديدة: لاستيراد مكتبة التعامل مع اكسل ---
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

from docx import Document
from docx.shared import Cm
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH



# ================== الجزء الثاني: الإعدادات الأولية والدوال المساعدة ==================

log_queue = queue.Queue()
executor = ThreadPoolExecutor(max_workers=1)

def get_correct_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ================== بداية الكود الجديد لتحديد مسار البيانات بذكاء ==================
import shutil # تأكد من وجود هذه المكتبة مع بقية المكتبات المستوردة

def get_data_directory():
    """
    يحدد المسار الصحيح لمجلد 'data' بناءً على بيئة التشغيل.
    """
    # التحقق مما إذا كان البرنامج يعمل كملف تنفيذي (مُجمّد)
    if getattr(sys, 'frozen', False):
        executable_path = os.path.dirname(sys.executable)

        # التحقق مما إذا كان يعمل من مجلد نظام محمي مثل Program Files
        if 'program files' in executable_path.lower():
            # الحالة 1: البرنامج مثبت -> استخدم مجلد AppData
            app_name = "ExamGuardScheduler" # اسم فريد لبرنامجك
            app_data_path = os.path.join(os.getenv('APPDATA'), app_name)

            # عند أول تشغيل، قم بنسخ مجلد 'data' من مكان التثبيت إلى AppData
            install_source_data_dir = os.path.join(executable_path, 'data')
            if not os.path.exists(app_data_path) and os.path.exists(install_source_data_dir):
                shutil.copytree(install_source_data_dir, app_data_path)

            return app_data_path
        else:
            # الحالة 2: البرنامج يعمل كـ exe ولكن من مجلد عادي (مثل dist)
            return os.path.join(executable_path, 'data')
    else:
        # الحالة 3: البرنامج يعمل كسكربت .py في بيئة التطوير
        return os.path.join(os.path.dirname(__file__), 'data')

# استدعاء الدالة لتحديد المسارات النهائية
DATA_DIR = get_data_directory()
DB_PATH = os.path.join(DATA_DIR, 'database.db')
# ================== نهاية الكود الجديد ==================

# --- دوال التعامل مع قاعدة البيانات ---

def get_db_connection():
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS professors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE
    )''')
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS halls (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        type TEXT NOT NULL
    )''')
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS levels (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE
    )''')
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS subjects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        level_id INTEGER NOT NULL,
        FOREIGN KEY (level_id) REFERENCES levels (id) ON DELETE CASCADE
    )''')
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS assignments (
        professor_id INTEGER NOT NULL,
        subject_id INTEGER NOT NULL,
        PRIMARY KEY (professor_id, subject_id),
        FOREIGN KEY (professor_id) REFERENCES professors (id) ON DELETE CASCADE,
        FOREIGN KEY (subject_id) REFERENCES subjects (id) ON DELETE CASCADE
    )''')
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    )''')
    conn.commit()
    conn.close()
    print("Database initialized successfully.")

# --- دوال مساعدة أخرى ---

def clean_string_for_matching(text):
    if not isinstance(text, str): return text
    text = text.strip()
    return re.sub(r'\s+', ' ', text)

def parse_unique_id(unique_id, all_levels):
    sorted_levels = sorted(all_levels, key=len, reverse=True)
    for level in sorted_levels:
        if unique_id.endswith(f"({level})"):
            return unique_id[:-(len(level) + 2)].strip(), level
    last_paren_index = unique_id.rfind('(')
    if last_paren_index != -1 and unique_id.endswith(')'):
        return unique_id[:last_paren_index].strip(), unique_id[last_paren_index + 1:-1].strip()
    return unique_id, None

def apply_styles_to_cell(cell, font, alignment, border, fill=None):
    cell.font = font
    cell.alignment = alignment
    cell.border = border
    if fill: cell.fill = fill

# ================== دوال الخوارزمية (المنطق الأساسي) ==================
# الصق محتوى الدوال الأصلية هنا

def calculate_balanced_distribution(total_large, total_other, num_profs, w_large, w_other):
    if num_profs == 0: return []
    total_workload = (total_large * w_large) + (total_other * w_other)
    target_workload = total_workload / num_profs
    
    distribution = []
    base_large = total_large // num_profs
    remainder_large = total_large % num_profs
    
    for i in range(num_profs):
        large_count = base_large + 1 if i < remainder_large else base_large
        rem_workload = target_workload - (large_count * w_large)
        other_count = max(0, round(rem_workload / w_other))
        distribution.append({'large': large_count, 'other': other_count, 'total_workload': (large_count * w_large) + (other_count * w_other)})
    return distribution

def generate_balance_report(prof_stats, prof_targets):
    patterns = defaultdict(int)
    for stats in prof_stats.values():
        patterns[(stats['large'], stats['other'])] += 1
        
    target_patterns = defaultdict(int)
    if prof_targets:
        for target in prof_targets.values():
            target_patterns[(target['large'], target['other'])] += 1

    report_details = []
    all_keys = sorted(list(set(patterns.keys()) | set(target_patterns.keys())))
    total_deviation = 0
    
    for key in all_keys:
        actual = patterns.get(key, 0)
        target = target_patterns.get(key, 0)
        deviation = actual - target
        total_deviation += abs(deviation)
        report_details.append({'pattern': f"{key[0]} كبيرة + {key[1]} أخرى", 'target_count': target, 'actual_count': actual, 'deviation': deviation})

    balance_score = max(0, 100 - (total_deviation * 2))
    return {'details': report_details, 'balance_score': round(balance_score)}

def calculate_schedule_balance_score(schedule, all_professors, settings, num_professors):
    if not schedule or not all_professors: return 0.0
    large_hall_weight = float(settings.get('largeHallWeight', 3.0))
    other_hall_weight = float(settings.get('otherHallWeight', 1.0))
    guards_large_hall = int(settings.get('guardsLargeHall', 4))
    enable_custom_targets = settings.get('enableCustomTargets', False)
    custom_target_patterns = settings.get('customTargetPatterns', [])
    
    prof_stats = {prof: {'large': 0, 'other': 0} for prof in all_professors}
    all_exams = [exam for date_slots in schedule.values() for time_slots in date_slots.values() for exam in time_slots]
    for exam in all_exams:
        guards_copy = list(exam.get('guards', []))
        large_guards_needed = sum(guards_large_hall for h in exam.get('halls', []) if h.get('type') == 'كبيرة')
        large_hall_guards, other_hall_guards = guards_copy[:large_guards_needed], guards_copy[large_guards_needed:]
        for guard in large_hall_guards:
            if guard in prof_stats: prof_stats[guard]['large'] += 1
        for guard in other_hall_guards:
            if guard in prof_stats: prof_stats[guard]['other'] += 1
    prof_targets_map = {}
    if enable_custom_targets and custom_target_patterns:
        prof_targets_list = []
        for pattern in custom_target_patterns:
            for _ in range(pattern.get('count', 0)):
                prof_targets_list.append({'large': pattern.get('large', 0), 'other': pattern.get('other', 0)})
        shuffled_profs = list(all_professors); random.shuffle(shuffled_profs)
        prof_targets_map = {prof: prof_targets_list[i] for i, prof in enumerate(shuffled_profs) if i < len(prof_targets_list)}
    else:
        total_large_slots_final = sum(stats['large'] for stats in prof_stats.values())
        total_other_slots_final = sum(stats['other'] for stats in prof_stats.values())
        prof_targets_list = calculate_balanced_distribution(total_large_slots_final, total_other_slots_final, num_professors, large_hall_weight, other_hall_weight)
        if prof_targets_list:
            prof_targets_map = {prof: prof_targets_list[i % len(prof_targets_list)] for i, prof in enumerate(sorted(prof_stats.keys()))}
    if not prof_targets_map: return 0.0
    balance_report = generate_balance_report(prof_stats, prof_targets_map)
    return balance_report.get('balance_score', 0.0)

def is_assignment_valid(prof, exam, prof_assignments, prof_large_counts, settings, date_map):
    """
    دالة مركزية للتحقق مما إذا كان تعيين حارس لامتحان معين صالحاً أم لا.
    """
    # استخلاص الإعدادات من القاموس
    duty_patterns = settings.get('dutyPatterns', {})
    unavailable_days = settings.get('unavailableDays', {})
    max_shifts = int(settings.get('maxShifts', '0')) if settings.get('maxShifts', '0') != '0' else float('inf')
    max_large_hall_shifts = int(settings.get('maxLargeHallShifts', '2')) if settings.get('maxLargeHallShifts', '2') != '0' else float('inf')
    
    # 1. التحقق من التزامن (مشغول في نفس الوقت)
    if any(e['date'] == exam['date'] and e['time'] == exam['time'] for e in prof_assignments.get(prof, [])):
        return False

    # 2. التحقق من أيام الغياب
    if exam['date'] in unavailable_days.get(prof, []):
        return False

    # 3. التحقق من سقف الحصص الإجمالي
    if len(prof_assignments.get(prof, [])) >= max_shifts:
        return False

    # 4. التحقق من سقف حصص القاعة الكبيرة
    is_large_hall_exam = any(h['type'] == 'كبيرة' for h in exam['halls'])
    if is_large_hall_exam and prof_large_counts.get(prof, 0) >= max_large_hall_shifts:
        return False

    # 5. التحقق من نمط الحراسة
    # --- بداية: أسطر التعريفات التي كانت مفقودة ---
    prof_pattern = duty_patterns.get(prof, 'flexible_2_days')
    duties_dates = {d['date'] for d in prof_assignments.get(prof, [])}
    is_new_day = exam['date'] not in duties_dates
    num_duty_days = len(duties_dates)
    # --- نهاية: أسطر التعريفات التي كانت مفقودة ---

    if is_new_day:
        if (prof_pattern == 'one_day_only' and num_duty_days >= 1) or \
           (prof_pattern == 'flexible_2_days' and num_duty_days >= 2) or \
           (prof_pattern == 'flexible_3_days' and num_duty_days >= 3) or \
           (prof_pattern == 'consecutive_strict' and num_duty_days >= 2):
            return False
        elif prof_pattern == 'consecutive_strict' and num_duty_days == 1:
            idx1 = date_map.get(list(duties_dates)[0])
            idx2 = date_map.get(exam['date'])
            if idx1 is None or idx2 is None or abs(idx1 - idx2) != 1:
                return False

    # إذا نجح في كل الاختبارات، يكون التعيين صالحاً
    return True

def is_schedule_valid(schedule, settings, all_professors, duty_patterns, date_map):
    """
    النسخة المحدثة: مع إضافة التحقق من قيد "أزواج الأساتذة".
    """
    unavailable_days = settings.get('unavailableDays', {})
    max_shifts = int(settings.get('maxShifts', '0')) if settings.get('maxShifts', '0') != '0' else float('inf')
    max_large_hall_shifts = int(settings.get('maxLargeHallShifts', '2')) if settings.get('maxLargeHallShifts', '2') != '0' else float('inf')

    prof_assignments = defaultdict(list)
    prof_large_counts = defaultdict(int)
    
    all_exams = [exam for date_slots in schedule.values() for time_slots in date_slots.values() for exam in time_slots]

    # التحقق من القيود الأساسية (حصص متزامنة، غياب، عدد أقصى للحصص)
    for exam in all_exams:
        is_large_exam = any(h['type'] == 'كبيرة' for h in exam.get('halls', []))
        for guard in exam.get('guards', []):
            if guard == "**نقص**":
                return False 
            
            for other_exam in prof_assignments.get(guard, []):
                if other_exam['date'] == exam['date'] and other_exam['time'] == exam['time']:
                    return False 
                    
            if exam['date'] in unavailable_days.get(guard, []):
                return False 
            
            prof_assignments[guard].append(exam)
            if is_large_exam:
                prof_large_counts[guard] += 1

    for prof in all_professors:
        if len(prof_assignments[prof]) > max_shifts:
            return False
        if prof_large_counts[prof] > max_large_hall_shifts:
            return False
            
    # التحقق من نمط أيام الحراسة
    prof_assigned_slots = defaultdict(list)
    for exam in all_exams:
        for guard in exam.get('guards', []):
            if guard != "**نقص**":
                prof_assigned_slots[guard].append((exam['date'], exam['time']))

    for prof, pattern in duty_patterns.items():
        if not prof_assigned_slots.get(prof):
            continue
        duties_dates_indices = sorted(list({date_map.get(d_date) for d_date, d_time in prof_assigned_slots.get(prof, []) if date_map.get(d_date) is not None}))
        if not duties_dates_indices and prof_assigned_slots.get(prof):
             continue
        num_unique_duty_days = len(duties_dates_indices)
        if pattern == 'consecutive_strict':
            if num_unique_duty_days > 0 and (num_unique_duty_days != 2 or (duties_dates_indices[1] - duties_dates_indices[0] != 1)):
                return False
        elif pattern == 'one_day_only':
            if num_unique_duty_days > 1:
                return False
        elif pattern == 'flexible_2_days':
            if num_unique_duty_days > 0 and num_unique_duty_days != 2:
                return False
        elif pattern == 'flexible_3_days':
            if num_unique_duty_days > 0 and (num_unique_duty_days < 2 or num_unique_duty_days > 3):
                return False
    
    # --- بداية: التحقق من قيد أزواج الأساتذة ---
    professor_pairs = settings.get('professorPartnerships', []) # تم تغيير اسم المفتاح
    if professor_pairs:
        # أولاً، قم ببناء قاموس يحتوي على أيام الحراسة لكل أستاذ
        prof_duty_days = defaultdict(set)
        for guard, duties in prof_assigned_slots.items():
            for duty_date, _ in duties:
                prof_duty_days[guard].add(duty_date)

        # ثانياً، قم بالمرور على كل زوج وتحقق من تطابق أيام الحراسة
        for pair in professor_pairs:
            if len(pair) == 2:
                prof1_name, prof2_name = pair[0], pair[1]
                prof1_days = prof_duty_days.get(prof1_name, set())
                prof2_days = prof_duty_days.get(prof2_name, set())
                # إذا كانت مجموعة أيام الأستاذ الأول لا تساوي مجموعة أيام الثاني، فالقيد مكسور
                if prof1_days != prof2_days:
                    return False
    # --- نهاية: التحقق من قيد أزواج الأساتذة ---

    # إذا نجح الجدول في كل الاختبارات
    return True

# ================== الواجهة الرئيسية لتشغيل الخوارزمية (النسخة النهائية مع التحسينات) ==================

def analyze_failure_reason(exam, all_professors, prof_assignments, unavailable_days, max_shifts, max_large_hall_shifts, prof_large_hall_counts, duty_patterns, date_map, current_schedule):
    date, time = exam['date'], exam['time']
    is_large_hall_exam = any(h['type'] == 'كبيرة' for h in exam['halls'])
    
    reasons = Counter()
    for prof in all_professors:
        if any(prof in e['guards'] for e in current_schedule[date][time]):
            reasons['مشغول في نفس الوقت'] += 1
            continue
        if date in unavailable_days.get(prof, []):
            reasons['غير متاح في هذا اليوم'] += 1
            continue
        if len(prof_assignments.get(prof, [])) >= max_shifts:
            reasons['وصل للحد الأقصى من الحصص'] += 1
            continue
        if is_large_hall_exam and prof_large_hall_counts.get(prof, 0) >= max_large_hall_shifts:
            reasons['وصل للحد الأقصى من القاعات الكبيرة'] += 1
            continue
        # (Could add more detailed checks for duty patterns here if needed)

    if not reasons:
        return "سبب غير معروف، قد يكون متعلقاً بقيود نمط الحراسة المعقدة."

    most_common_reason, count = reasons.most_common(1)[0]
    return f"لم يتم العثور على حارس للامتحان '{exam['subject']}' في {date} {time}. السبب الأكثر شيوعًا: {count} أستاذًا كانوا '{most_common_reason}'."

def run_subject_optimization_phase(schedule, assignments, all_levels_list, subject_owners, settings, log_q):
    """
    المرحلة 1.5: تقوم هذه الدالة بتحسين جدول المواد عبر محاولة تجميع مواد كل أستاذ
    في أقل عدد ممكن من الأيام، مع احترام كل القيود.
    """
    log_q.put(">>> بدء المرحلة 1.5: تحسين تجميع مواد الأساتذة...")
    
    # استخراج الإعدادات ذات الصلة
    optimization_attempts = int(settings.get('optimizationAttempts', 5000)) # يمكننا إضافة هذا الإعداد لاحقاً للتحكم
    last_day_restriction = settings.get('lastDayRestriction', 'none')
    exam_schedule_settings = settings.get('examSchedule', {})
    
    optimized_schedule = copy.deepcopy(schedule)

    # --- إعدادات قيد آخر يوم ---
    sorted_dates = sorted(exam_schedule_settings.keys())
    last_exam_day = sorted_dates[-1] if sorted_dates else None
    restricted_slots_on_last_day = []
    if last_exam_day and last_day_restriction != 'none':
        try:
            last_day_all_slots = sorted(exam_schedule_settings.get(last_exam_day, []), key=lambda x: x['time'])
            num_to_restrict = int(last_day_restriction.split('_')[1])
            restricted_slots_on_last_day = [s['time'] for s in last_day_all_slots[-num_to_restrict:]]
        except (ValueError, IndexError):
            pass # تجاهل في حالة وجود خطأ

    # --- بناء هياكل بيانات مساعدة ---
    prof_to_exams = defaultdict(list)
    all_exams_list = []
    for date, time_slots in optimized_schedule.items():
        for time, exams in time_slots.items():
            for exam in exams:
                all_exams_list.append(exam)
                owner = exam.get('professor')
                if owner and owner != "غير محدد":
                    prof_to_exams[owner].append(exam)

    # --- دالة لتقييم "جودة" الجدول الحالي ---
    def calculate_schedule_cost(p_to_e):
        total_cost = 0
        for prof, exams in p_to_e.items():
            if len(exams) > 1:
                unique_days = {e['date'] for e in exams}
                total_cost += len(unique_days)
        return total_cost

    initial_cost = calculate_schedule_cost(prof_to_exams)
    log_q.put(f"... التكلفة الأولية لتشتت الأيام: {initial_cost}")

    # --- حلقة التحسين والتبديل ---
    for attempt in range(optimization_attempts):
        # اختيار أستاذ لديه مواد متفرقة لمحاولة تحسين جدوله
        prof_to_optimize = None
        candidates = [p for p, e in prof_to_exams.items() if len({ex['date'] for ex in e}) > 1]
        if not candidates:
            break # لا يوجد أساتذة لتحسينهم، نخرج من الحلقة
        prof_to_optimize = random.choice(candidates)

        prof_exams = prof_to_exams[prof_to_optimize]
        if len(prof_exams) < 2: continue
        
        # اختيار امتحانين مختلفين في أيام مختلفة لنفس الأستاذ
        exam_a = random.choice(prof_exams)
        other_exams = [e for e in prof_exams if e['date'] != exam_a['date']]
        if not other_exams: continue
        exam_b = random.choice(other_exams)

        # نحاول نقل الامتحان (B) إلى نفس يوم الامتحان (A)
        target_date = exam_a['date']
        source_date = exam_b['date']
        time_slot = exam_b['time']
        
        # البحث عن شريك تبديل مناسب في اليوم المستهدف
        swap_candidates = optimized_schedule.get(target_date, {}).get(time_slot, [])
        if not swap_candidates: continue
        
        # يجب أن يكون الشريك من نفس المستوى لضمان صحة الفترة الزمنية
        exam_c = next((c for c in swap_candidates if c['level'] == exam_b['level'] and c['subject'] != exam_b['subject']), None)

        if not exam_c: continue
        
        # --- التحقق من القيود قبل التبديل ---
        # 1. هل التبديل سيضع أي امتحان في فترة محظورة في آخر يوم؟
        if (target_date == last_exam_day and time_slot in restricted_slots_on_last_day) or \
           (source_date == last_exam_day and exam_c['time'] in restricted_slots_on_last_day):
            continue

        # 2. هل الأستاذ الخاص بالامتحان (C) سيتحسن جدوله أو يبقى كما هو؟
        owner_c = exam_c.get('professor')
        if owner_c and owner_c != "غير محدد":
            current_days_c = {e['date'] for e in prof_to_exams.get(owner_c, [])}
            # إضافة يوم جديد وإزالة يوم قديم
            new_days_c = (current_days_c - {target_date}) | {source_date}
            # إذا كان التبديل سيزيد من عدد أيام الأستاذ الآخر، فهو تبديل سيء
            if len(new_days_c) > len(current_days_c):
                continue

        # --- تنفيذ التبديل ---
        # إزالة B و C من أماكنهما الأصلية
        optimized_schedule[source_date][time_slot].remove(exam_b)
        optimized_schedule[target_date][time_slot].remove(exam_c)
        
        # إضافة B و C إلى أماكنهما الجديدة
        optimized_schedule[target_date][time_slot].append(exam_b)
        optimized_schedule[source_date][time_slot].append(exam_c)
        
        # تحديث بيانات الامتحانين
        exam_b['date'], exam_c['date'] = target_date, source_date
        
        # تحديث هياكل البيانات المساعدة لإعادة الحساب
        prof_to_exams[prof_to_optimize].remove(exam_b)
        prof_to_exams[prof_to_optimize].append(exam_b) # تم تحديث تاريخه
        if owner_c in prof_to_exams:
            prof_to_exams[owner_c].remove(exam_c)
            prof_to_exams[owner_c].append(exam_c) # تم تحديث تاريخه

    final_cost = calculate_schedule_cost(prof_to_exams)
    log_q.put(f"✓ انتهاء المرحلة 1.5. التكلفة النهائية لتشتت الأيام: {final_cost}")
    
    return optimized_schedule

def run_post_processing_swaps(schedule, prof_assignments, prof_workload, prof_large_counts, settings, all_professors, date_map, swap_attempts, locked_guards=set()):
    large_hall_weight = float(settings.get('largeHallWeight', 3.0))
    other_hall_weight = float(settings.get('otherHallWeight', 1.0))
    duty_patterns = settings.get('dutyPatterns', {})
    unavailable_days = settings.get('unavailableDays', {})
    max_shifts = int(settings.get('maxShifts', '0')) if settings.get('maxShifts', '0') != '0' else float('inf')
    max_large_hall_shifts = int(settings.get('maxLargeHallShifts', '2')) if settings.get('maxLargeHallShifts', '2') != '0' else float('inf')
    
    temp_schedule = copy.deepcopy(schedule)
    
    # بناء/تحديث القواميس المساعدة من الجدول الحالي لضمان دقتها
    temp_assignments = defaultdict(list)
    temp_large_counts = defaultdict(int)
    temp_workload = defaultdict(float)
    all_exams_flat = [exam for day in temp_schedule.values() for slot in day.values() for exam in slot]
    for exam in all_exams_flat:
        is_large = any(h['type'] == 'كبيرة' for h in exam.get('halls', []))
        duty_weight = large_hall_weight if is_large else other_hall_weight
        for g in exam.get('guards', []):
            if g != "**نقص**":
                temp_assignments[g].append(exam)
                temp_workload[g] += duty_weight
                if is_large:
                    temp_large_counts[g] += 1

    for _ in range(swap_attempts):
        if not temp_workload or len(temp_workload) < 2: break
        
        most_burdened_prof = max(temp_workload, key=temp_workload.get)
        least_burdened_prof = min(temp_workload, key=temp_workload.get)

        if most_burdened_prof == least_burdened_prof or temp_workload[most_burdened_prof] <= temp_workload[least_burdened_prof]:
            break 
            
        swap_found = False
        
        possible_swaps = [
            exam for exam in temp_assignments.get(most_burdened_prof, [])
            if (exam.get('uuid'), most_burdened_prof) not in locked_guards
        ]
        random.shuffle(possible_swaps)

        for exam in possible_swaps:
            date, time = exam['date'], exam['time']
            is_large_hall_exam = any(h['type'] == 'كبيرة' for h in exam.get('halls', []))
            # تجميع الإعدادات في قاموس واحد
            settings_for_validation = {
                'dutyPatterns': duty_patterns,
                'unavailableDays': unavailable_days,
                'maxShifts': max_shifts,
                'maxLargeHallShifts': max_large_hall_shifts
            }

            # استدعاء واحد يحل محل كل الشروط السابقة
            # لاحظ أننا نتحقق من صلاحية التعيين للأستاذ الأقل عبئاً (least_burdened_prof)
            if is_assignment_valid(least_burdened_prof, exam, temp_assignments, temp_large_counts, settings_for_validation, date_map):
                # --- بداية التعديل الشامل والمقترح ---
                exam_in_schedule = next((e for e in temp_schedule[date][time] if e.get('uuid') == exam.get('uuid')), None)
                if not exam_in_schedule: continue

                # الخطوة 1: الحذف الآمن للحارس القديم من قائمة الحراس
                try:
                    # نجد أول ظهور للحارس ونحذفه. هذا أكثر أمانًا من .remove()
                    guard_index_to_remove = exam_in_schedule['guards'].index(most_burdened_prof)
                    del exam_in_schedule['guards'][guard_index_to_remove]
                    exam_in_schedule['guards'].append(least_burdened_prof)
                except ValueError:
                    # إذا لم يتم العثور على الحارس، نتجاهل هذا التبديل وننتقل للتالي
                    # هذا يعالج الخطأ مباشرة "list.remove(x): x not in list"
                    continue

                # الخطوة 2: تحديث إحصائيات عبء العمل
                duty_weight = large_hall_weight if is_large_hall_exam else other_hall_weight
                temp_workload[most_burdened_prof] -= duty_weight
                temp_workload[least_burdened_prof] += duty_weight
                if is_large_hall_exam:
                    temp_large_counts[most_burdened_prof] -= 1
                    temp_large_counts[least_burdened_prof] = temp_large_counts.get(least_burdened_prof, 0) + 1
                
                # الخطوة 3: الحذف الآمن للامتحان من قائمة مهام الحارس القديم
                exam_to_remove_index = -1
                for i, assigned_exam in enumerate(temp_assignments[most_burdened_prof]):
                    if assigned_exam.get('uuid') == exam.get('uuid'):
                        exam_to_remove_index = i
                        break
                
                if exam_to_remove_index != -1:
                    del temp_assignments[most_burdened_prof][exam_to_remove_index]
                else:
                    # هذا لا يفترض أن يحدث، ولكنه إجراء احترازي لمنع أخطاء مستقبلية
                    continue 

                # الخطوة 4: إضافة الامتحان لقائمة مهام الحارس الجديد
                temp_assignments[least_burdened_prof].append(exam)
                
                swap_found = True
                break
                # --- نهاية التعديل الشامل والمقترح ---
        
        if not swap_found:
            break 
            
    return temp_schedule, temp_assignments, temp_workload, temp_large_counts

# في ملف app.py، قم باستبدال هذه الدالة بالكامل

def run_simulated_annealing(schedule, prof_assignments, prof_workload, prof_large_counts, settings, all_professors, date_map, duty_patterns, annealing_iterations, annealing_temp, annealing_cooling):
    """
    النسخة النهائية والمحسنة.
    تستهدف الأنماط المخصصة إذا كانت مفعلة، وإلا تقوم بالموازنة العامة.
    """
    # --- استخلاص الإعدادات ---
    large_hall_weight = float(settings.get('largeHallWeight', 3.0))
    other_hall_weight = float(settings.get('otherHallWeight', 1.0))
    enable_custom_targets = settings.get('enableCustomTargets', False)
    custom_target_patterns = settings.get('customTargetPatterns', [])
    guards_large_hall = int(settings.get('guardsLargeHall', 4))

    # --- تحديد دالة الطاقة التي سنستخدمها ---
    energy_function = None

    if enable_custom_targets and custom_target_patterns:
        # الهدف: تقليل الانحراف عن الأنماط المستهدفة
        target_counts = Counter((p['large'], p['other']) for p in custom_target_patterns for _ in range(p.get('count', 0)))
        all_target_patterns = set(target_counts.keys())

        def calculate_pattern_deviation_energy(sch):
            prof_stats = {prof: {'large': 0, 'other': 0} for prof in all_professors}
            all_exams = [exam for date_slots in sch.values() for time_slots in date_slots.values() for exam in time_slots]
            for exam in all_exams:
                guards_copy = [g for g in exam.get('guards', []) if g != "**نقص**"]
                large_guards_needed = sum(guards_large_hall for h in exam.get('halls', []) if h.get('type') == 'كبيرة')
                large_hall_guards = guards_copy[:large_guards_needed]
                other_hall_guards = guards_copy[large_guards_needed:]
                for guard in large_hall_guards:
                    if guard in prof_stats: prof_stats[guard]['large'] += 1
                for guard in other_hall_guards:
                    if guard in prof_stats: prof_stats[guard]['other'] += 1
            
            actual_counts = Counter((s['large'], s['other']) for s in prof_stats.values())
            
            total_deviation = 0
            all_current_patterns = set(actual_counts.keys()) | all_target_patterns
            for pattern in all_current_patterns:
                total_deviation += abs(actual_counts.get(pattern, 0) - target_counts.get(pattern, 0))
            
            return total_deviation
        
        energy_function = calculate_pattern_deviation_energy

    else:
        # الهدف الافتراضي: الموازنة العامة
        def calculate_workload_energy(sch):
            workload_dict = defaultdict(float)
            all_exams = [exam for date_slots in sch.values() for time_slots in date_slots.values() for exam in time_slots]
            for exam in all_exams:
                 is_large_exam = any(h['type'] == 'كبيرة' for h in exam.get('halls', []))
                 duty_weight = large_hall_weight if is_large_exam else other_hall_weight
                 for guard in exam.get('guards', []):
                     if guard != "**نقص**" and guard in all_professors:
                         workload_dict[guard] += duty_weight
            
            if not workload_dict: return 0.0
            workloads = list(workload_dict.values())
            return max(workloads) - min(workloads)

        energy_function = calculate_workload_energy

    # --- إعدادات الحالة الأولية ---
    current_schedule = copy.deepcopy(schedule)
    current_energy = energy_function(current_schedule)
    
    best_schedule = copy.deepcopy(current_schedule)
    best_energy = current_energy

    temp = annealing_temp
    
    # --- حلقة التلدين المحاكي الرئيسية ---
    for i in range(annealing_iterations):
        if temp < 0.01: break

        all_duties = [{'exam': ex, 'guard': g, 'idx': g_idx} 
                      for ex_list in current_schedule.values() 
                      for exs in ex_list.values() for ex in exs 
                      for g_idx, g in enumerate(ex.get('guards', [])) if g != "**نقص**"]

        if not all_duties: break

        duty_to_change = random.choice(all_duties)
        exam = duty_to_change['exam']
        prof1 = duty_to_change['guard']
        guard_idx = duty_to_change['idx']

        possible_new_profs = [p for p in all_professors if p != prof1]
        if not possible_new_profs: continue
        prof2 = random.choice(possible_new_profs)
        
        schedule_copy = copy.deepcopy(current_schedule)
        exam_in_copy = None
        for date_key, time_slots in schedule_copy.items():
            if date_key == exam['date']:
                for time_key, exams in time_slots.items():
                    if time_key == exam['time']:
                        for ex_item in exams:
                            if ex_item['subject'] == exam['subject'] and ex_item['level'] == exam['level']:
                                exam_in_copy = ex_item
                                break
                if exam_in_copy: break
        
        if not exam_in_copy: continue
        
        exam_in_copy['guards'][guard_idx] = prof2
        
        if not is_schedule_valid(schedule_copy, settings, all_professors, duty_patterns, date_map):
            continue

        new_energy = energy_function(schedule_copy)
        
        if new_energy < current_energy or random.random() < math.exp((current_energy - new_energy) / temp):
            current_schedule = schedule_copy
            current_energy = new_energy
            
            if current_energy < best_energy:
                best_energy = current_energy
                best_schedule = copy.deepcopy(current_schedule)
        
        temp *= annealing_cooling

    # إعادة بناء بقية بيانات الحل النهائي
    final_assignments = defaultdict(list)
    final_large_counts = defaultdict(int)
    final_workload = defaultdict(float)
    for ex_list in best_schedule.values():
        for exs in ex_list.values():
            for ex in exs:
                is_large = any(h['type'] == 'كبيرة' for h in ex['halls'])
                duty_weight = large_hall_weight if is_large else other_hall_weight
                for g in ex.get('guards', []):
                    if g != "**نقص**":
                        final_assignments[g].append(ex)
                        final_workload[g] += duty_weight
                        if is_large:
                            final_large_counts[g] += 1
    
    return best_schedule, final_assignments, final_workload, final_large_counts

# أضف هذه الدالة الكاملة في ملف app.py
# في ملف app.py، استبدل هذه الدالة بالكامل
# في ملف app.py، قم باستبدال هذه الدالة بالكامل
def run_tabu_search(initial_schedule, settings, all_professors, duty_patterns, date_map, log_q, locked_guards=set()):
    """
    النسخة النهائية والمصححة: مع دالة تكلفة ذكية تستهدف الأنماط اليدوية بشكل صحيح.
    """
    log_q.put(">>> تشغيل البحث المحظور (Tabu Search)...")

    # --- استخلاص الإعدادات ---
    max_iterations = int(settings.get('tabuIterations', 100))
    tabu_tenure = int(settings.get('tabuTenure', 15))
    neighborhood_size = int(settings.get('tabuNeighborhoodSize', 50))
    large_hall_weight = float(settings.get('largeHallWeight', 3.0))
    other_hall_weight = float(settings.get('otherHallWeight', 1.0))
    guards_large_hall = int(settings.get('guardsLargeHall', 4))
    # --- إضافة جديدة: استخلاص إعدادات الأنماط اليدوية ---
    enable_custom_targets = settings.get('enableCustomTargets', False)
    custom_target_patterns = settings.get('customTargetPatterns', [])

    # =================================================================
    # ## بداية دالة التكلفة المصححة ##
    # =================================================================
    def calculate_cost(sch):
        # 1. حساب الإحصائيات الأساسية (عدد الحصص الكبيرة والأخرى لكل أستاذ)
        prof_stats = defaultdict(lambda: {'large': 0, 'other': 0})
        for day in sch.values():
            for slot in day.values():
                for exam in slot:
                    guards_copy = [g for g in exam.get('guards', []) if g != "**نقص**"]
                    large_guards_needed = sum(guards_large_hall for h in exam.get('halls', []) if h.get('type') == 'كبيرة')
                    
                    # تقسيم الحراس بشكل صحيح
                    large_hall_guards = guards_copy[:large_guards_needed]
                    other_hall_guards = guards_copy[large_guards_needed:]

                    for guard in large_hall_guards:
                        if guard in all_professors: prof_stats[guard]['large'] += 1
                    for guard in other_hall_guards:
                        if guard in all_professors: prof_stats[guard]['other'] += 1

        # 2. تحديد الهدف بناءً على الإعدادات
        if enable_custom_targets and custom_target_patterns:
            # --- الهدف: تقليل الانحراف عن الأنماط اليدوية ---
            target_counts = Counter((p['large'], p['other']) for p in custom_target_patterns for _ in range(p.get('count', 0)))
            actual_counts = Counter((s['large'], s['other']) for s in prof_stats.values())
            
            total_deviation = 0
            all_patterns = set(target_counts.keys()) | set(actual_counts.keys())
            for pattern in all_patterns:
                total_deviation += abs(actual_counts.get(pattern, 0) - target_counts.get(pattern, 0))
            
            return total_deviation
        else:
            # --- الهدف الافتراضي: موازنة عبء العمل ---
            workload = {p: s['large'] * large_hall_weight + s['other'] * other_hall_weight for p, s in prof_stats.items()}
            if not workload: return 0.0
            workloads = list(workload.values())
            return max(workloads) - min(workloads) if workloads else 0.0
    # =================================================================
    # ## نهاية دالة التكلفة المصححة ##
    # =================================================================

    # --- الإعدادات الأولية للبحث ---
    current_solution = copy.deepcopy(initial_schedule)
    best_solution = copy.deepcopy(current_solution)
    best_cost = calculate_cost(best_solution)
    tabu_list = deque(maxlen=tabu_tenure)
    log_q.put(f"... [Tabu Search] التكلفة الأولية = {best_cost:.2f}")

    # --- حلقة البحث الرئيسية ---
    for i in range(max_iterations):
        # =================> بداية الإضافة الجديدة <=================
        # إرسال التقدم بناءً على الدورة الحالية للبحث المحظور
        percent_complete = int(((i + 1) / max_iterations) * 100)
        log_q.put(f"PROGRESS:{percent_complete}")
        # =================> نهاية الإضافة الجديدة <=================
        best_neighbor, best_neighbor_cost, best_neighbor_move = None, float('inf'), None
        all_duties = [(exam, guard, d_idx) 
                      for day in current_solution.values() for slot in day.values() for exam in slot
                      for d_idx, guard in enumerate(exam.get('guards',[])) 
                      if guard != "**نقص**" and (exam.get('uuid'), guard) not in locked_guards]
    
        if not all_duties:
            log_q.put("... لا توجد مهام حراسة قابلة للتغيير.")
            break

        # ... (بقية كود حلقة البحث يبقى كما هو دون تغيير) ...
        for _ in range(neighborhood_size):
            exam_to_change, prof1, guard_idx = random.choice(all_duties)
            possible_profs = [p for p in all_professors if p != prof1]
            if not possible_profs: continue
            prof2 = random.choice(possible_profs)

            neighbor = copy.deepcopy(current_solution)
            exam_in_neighbor = next(e for e in neighbor[exam_to_change['date']][exam_to_change['time']] if e.get('uuid') == exam_to_change.get('uuid'))
            
            old_guards = list(exam_in_neighbor['guards'])
            old_guards[guard_idx] = prof2
            exam_in_neighbor['guards'] = old_guards

            if not is_schedule_valid(neighbor, settings, all_professors, duty_patterns, date_map):
                continue

            neighbor_cost = calculate_cost(neighbor)
            move = (exam_to_change.get('uuid'), prof1, prof2)
            reverse_move = (exam_to_change.get('uuid'), prof2, prof1)

            if reverse_move in tabu_list:
                if neighbor_cost < best_cost: # Aspiration criterion
                    best_neighbor, best_neighbor_cost, best_neighbor_move = neighbor, neighbor_cost, move
                    break
            elif neighbor_cost < best_neighbor_cost:
                best_neighbor, best_neighbor_cost, best_neighbor_move = neighbor, neighbor_cost, move

        if best_neighbor is None:
            log_q.put(f"... [Tabu Search] لم يتم العثور على جار صالح في الدورة {i+1}. إنهاء البحث.")
            break

        current_solution = best_neighbor
        tabu_list.append(best_neighbor_move)

        if best_neighbor_cost < best_cost:
            best_cost = best_neighbor_cost
            best_solution = best_neighbor
            log_q.put(f"... [Tabu Search] دورة {i+1}: تم العثور على حل أفضل بتكلفة = {best_cost:.2f}")

    log_q.put(f"✓ البحث المحظور انتهى بأفضل تكلفة: {best_cost:.2f}")
    return best_solution, None, None, None


# ================== بداية الكود الجديد والمصحح بالكامل ==================

def calculate_cost(
    schedule, all_professors, settings, duty_patterns, date_map
):
    """
    النسخة النهائية من دالة التكلفة.
    - تعطي الأولوية القصوى لإصلاح النقص في الحراسة.
    - إذا كان الجدول كاملاً، تقوم باستهداف أنماط التوزيع اليدوية (إذا كانت مفعلة).
    - إذا لم تكن مفعلة، تعمل على الموازنة العامة لعبء العمل.
    """
    # --- الخطوة 1: حساب النقص في الحراسة والتحقق من صلاحية القيود ---
    shortage_count = 0
    all_exams_flat = [exam for day in schedule.values() for slot in day.values() for exam in slot]
    for exam in all_exams_flat:
        shortage_count += exam.get('guards', []).count("**نقص**")

    # إذا كان هناك نقص، فالتكلفة عالية جداً لتركيز البحث على إصلاحه
    if shortage_count > 0:
        return (shortage_count * 1000), shortage_count, 0

    # إذا لم يكن هناك نقص، تحقق من القيود الصارمة الأخرى
    violations = 0
    if not is_schedule_valid(schedule, settings, all_professors, duty_patterns, date_map):
        violations = 1 # يمكن تفصيل هذا لاحقاً لحساب عدد الخروقات بدقة

    # --- الخطوة 2: حساب تكلفة التوازن بناءً على الإعدادات ---
    balance_cost = 0.0
    enable_custom_targets = settings.get('enableCustomTargets', False)
    custom_target_patterns = settings.get('customTargetPatterns', [])

    # حساب إحصائيات التوزيع الحالية (كم حصة كبيرة وأخرى لكل أستاذ)
    prof_stats = {prof: {'large': 0, 'other': 0} for prof in all_professors}
    guards_large_hall = int(settings.get('guardsLargeHall', 4))
    for exam in all_exams_flat:
        guards_copy = [g for g in exam.get('guards', []) if g != "**نقص**"]
        large_guards_needed = sum(guards_large_hall for h in exam.get('halls', []) if h.get('type') == 'كبيرة')
        large_hall_guards = guards_copy[:large_guards_needed]
        other_hall_guards = guards_copy[large_guards_needed:]
        for guard in large_hall_guards:
            if guard in prof_stats: prof_stats[guard]['large'] += 1
        for guard in other_hall_guards:
            if guard in prof_stats: prof_stats[guard]['other'] += 1

    # --- ✨ المنطق الذكي للتبديل بين أهداف التوازن ---
    if enable_custom_targets and custom_target_patterns:
        # الهدف: تقليل الانحراف عن الأنماط اليدوية
        target_counts = Counter((p['large'], p['other']) for p in custom_target_patterns for _ in range(p.get('count', 0)))
        actual_counts = Counter((s['large'], s['other']) for s in prof_stats.values())
        
        total_deviation = 0
        all_patterns = set(target_counts.keys()) | set(actual_counts.keys())
        for pattern in all_patterns:
            total_deviation += abs(actual_counts.get(pattern, 0) - target_counts.get(pattern, 0))
        balance_cost = total_deviation
    else:
        # الهدف الافتراضي: موازنة عبء العمل
        large_hall_weight = float(settings.get('largeHallWeight', 3.0))
        other_hall_weight = float(settings.get('otherHallWeight', 1.0))
        prof_workload = {p: s['large'] * large_hall_weight + s['other'] * other_hall_weight for p, s in prof_stats.items()}
        
        if prof_workload:
            workload_values = list(prof_workload.values())
            balance_cost = max(workload_values) - min(workload_values) if workload_values else 0.0

    # --- الخطوة 3: دمج كل المكونات في تكلفة نهائية واحدة ---
    total_cost = (violations * 100) + balance_cost
    return total_cost, shortage_count, violations


def run_large_neighborhood_search(
    initial_schedule, settings, all_professors, duty_patterns, date_map, log_q, locked_guards=set()
):
    """
    النسخة النهائية والمحسنة من LNS مع نسبة تدمير ديناميكية.
    """
    log_q.put(">>> تشغيل LNS مع نسبة تدمير ديناميكية...")

    # --- 1. استخلاص الإعدادات ---
    iterations = int(settings.get('lnsIterations', 100))
    # <--- بداية التعديل: إعدادات نسبة التدمير الديناميكية
    initial_destroy_fraction = float(settings.get('lnsDestroyFraction', 0.2))
    min_destroy_fraction = 0.05  # أقل نسبة تدمير يمكن الوصول إليها
    destroy_fraction_decay_rate = 0.995 # معدل التخفيض في كل دورة
    # <--- نهاية التعديل

    initial_temp = 10.0
    cooling_rate = 0.99
    
    settings_for_validation = {
        'dutyPatterns': duty_patterns,
        'unavailableDays': settings.get('unavailableDays', {}),
        'maxShifts': settings.get('maxShifts', '0'),
        'maxLargeHallShifts': settings.get('maxLargeHallShifts', '2')
    }
    large_hall_weight = float(settings.get('largeHallWeight', 3.0))
    other_hall_weight = float(settings.get('otherHallWeight', 1.0))

    # --- 2. الحل المبدئي وحساب التكلفة الأولية ---
    current_solution = copy.deepcopy(initial_schedule)
    best_solution_so_far = copy.deepcopy(current_solution)
    
    current_cost, shortage, violations = calculate_cost(current_solution, all_professors, settings, duty_patterns, date_map)
    best_cost_so_far = current_cost
    log_q.put(f"... [LNS] التكلفة الأولية = {current_cost:.2f} (نقص={shortage}, خروقات={violations})")

    temp = initial_temp
    # <--- بداية التعديل: استخدام متغير جديد للنسبة الديناميكية
    dynamic_destroy_fraction = initial_destroy_fraction
    # <--- نهاية التعديل

    # --- 3. حلقة LNS الرئيسية ---
    for i in range(iterations):
        if settings.get('should_stop_event', threading.Event()).is_set():
            break
            
        percent_complete = int(((i + 1) / iterations) * 100)
        log_q.put(f"PROGRESS:{percent_complete}")
        
        ruined_solution = copy.deepcopy(current_solution)

        # --- 4. مرحلة التدمير (Ruin) ---
        duties_to_destroy = []
        for day in ruined_solution.values():
            for slot in day.values():
                for exam in slot:
                    for g_idx, guard in enumerate(exam.get('guards', [])):
                        if guard != "**نقص**" and (exam.get('uuid'), guard) not in locked_guards:
                            duties_to_destroy.append({'exam': exam, 'guard_index': g_idx})

        random.shuffle(duties_to_destroy)
        
        # <--- بداية التعديل: استخدام النسبة الديناميكية لحساب عدد المهام المدمرة
        num_to_destroy = int(len(duties_to_destroy) * dynamic_destroy_fraction)
        # <--- نهاية التعديل
        
        destroyed_slots = []
        for j in range(min(num_to_destroy, len(duties_to_destroy))):
            duty_info = duties_to_destroy[j]
            exam = duty_info['exam']
            g_idx = duty_info['guard_index']
            exam['guards'][g_idx] = "**نقص**"
            destroyed_slots.append(exam)

        # --- 5. مرحلة الإصلاح الذكي (Intelligent Repair) ---
        # ... (هذا الجزء يبقى كما هو دون تغيير) ...
        prof_assignments = defaultdict(list)
        prof_large_counts = defaultdict(int)
        prof_workload = defaultdict(float)

        all_exams_in_ruined = [exam for day in ruined_solution.values() for slot in day.values() for exam in slot]
        for exam in all_exams_in_ruined:
            is_large = any(h['type'] == 'كبيرة' for h in exam.get('halls', []))
            duty_weight = large_hall_weight if is_large else other_hall_weight
            for guard in exam.get('guards', []):
                if guard != "**نقص**":
                    prof_assignments[guard].append(exam)
                    prof_workload[guard] += duty_weight
                    if is_large:
                        prof_large_counts[guard] += 1
        
        for exam_to_repair in destroyed_slots:
            is_large_repair_exam = any(h['type'] == 'كبيرة' for h in exam_to_repair.get('halls',[]))
            repair_duty_weight = large_hall_weight if is_large_repair_exam else other_hall_weight
            
            valid_candidates = []
            for prof in all_professors:
                if prof in exam_to_repair.get('guards', []): continue
                
                if is_assignment_valid(prof, exam_to_repair, prof_assignments, prof_large_counts, settings_for_validation, date_map):
                    valid_candidates.append((prof, prof_workload.get(prof, 0)))

            if valid_candidates:
                best_prof_found, _ = min(valid_candidates, key=lambda item: item[1])
                
                try:
                    shortage_index = exam_to_repair['guards'].index("**نقص**")
                    exam_to_repair['guards'][shortage_index] = best_prof_found
                    prof_assignments[best_prof_found].append(exam_to_repair)
                    prof_workload[best_prof_found] += repair_duty_weight
                    if is_large_repair_exam:
                        prof_large_counts[best_prof_found] = prof_large_counts.get(best_prof_found, 0) + 1
                except ValueError:
                    pass

        # --- 6. مرحلة القبول المصححة والآمنة ---
        repaired_solution = ruined_solution
        new_cost, new_shortage, new_violations = calculate_cost(repaired_solution, all_professors, settings, duty_patterns, date_map)
        
        if new_cost < current_cost:
            current_solution = repaired_solution
            current_cost = new_cost
        else:
            acceptance_probability = math.exp((current_cost - new_cost) / temp) if temp > 0 else 0
            if random.random() < acceptance_probability:
                current_solution = repaired_solution
                current_cost = new_cost
        
        if current_cost < best_cost_so_far:
            best_cost_so_far = current_cost
            best_solution_so_far = copy.deepcopy(current_solution)
            log_q.put(f"... [LNS] دورة {i+1}: تم إيجاد حل أفضل بتكلفة = {best_cost_so_far:.2f} (نقص={new_shortage}, خروقات={new_violations})")

        # --- تحديث المعاملات الديناميكية ---
        temp *= cooling_rate
        # <--- بداية التعديل: تخفيض نسبة التدمير للدورة التالية
        dynamic_destroy_fraction = max(min_destroy_fraction, dynamic_destroy_fraction * destroy_fraction_decay_rate)
        # <--- نهاية التعديل

    log_q.put(f"✓ انتهى LNS المحسن بأفضل تكلفة: {best_cost_so_far:.2f}")
    
    # ... (بقية الدالة تبقى كما هي)
    final_assignments = defaultdict(list)
    final_workload = defaultdict(float)
    final_large_counts = defaultdict(int)
    for day_data in best_solution_so_far.values():
        for slot_data in day_data.values():
            for exam in slot_data:
                is_large_exam = any(h['type'] == 'كبيرة' for h in exam['halls'])
                duty_weight = large_hall_weight if is_large_exam else other_hall_weight
                for guard in exam.get('guards', []):
                    if guard != "**نقص**":
                        final_assignments[guard].append(exam)
                        final_workload[guard] += duty_weight
                        if is_large_exam: final_large_counts[guard] += 1
                        
    return best_solution_so_far, final_assignments, final_workload, final_large_counts

# ================== نهاية الكود الجديد والمصحح بالكامل ==================


def run_variable_neighborhood_search(
    initial_schedule, settings, all_professors, duty_patterns, date_map, log_q, locked_guards=set()
):
    """
    النسخة المحسنة من البحث الجواري المتغير (VNS).
    - تستخدم دالة التكلفة الشاملة calculate_cost التي تعطي الأولوية لإصلاح النقص.
    - Shake: تقوم بتدمير وإصلاح عدد k من المهام لإحداث تغيير فعال.
    - Local Search: تستخدم run_post_processing_swaps لتحسين الحل المهزوز.
    - VNS Metaheuristic: تنتقل بين الجوارات المختلفة (k) للهروب من الحلول المحلية.
    """
    log_q.put(">>> تشغيل البحث الجواري المتغير (VNS) المحسن...")

    # --- استخلاص الإعدادات ---
    iterations = int(settings.get('vnsIterations', 100))
    k_max = int(settings.get('vnsMaxK', 10)) # أقصى عدد مهام يتم تدميرها في الهزة الواحدة
    local_search_swaps = 50 # عدد محاولات البحث المحلي داخل كل دورة

    # إعداد قاموس الإعدادات للفحص مرة واحدة
    settings_for_validation = {
        'dutyPatterns': duty_patterns,
        'unavailableDays': settings.get('unavailableDays', {}),
        'maxShifts': settings.get('maxShifts', '0'),
        'maxLargeHallShifts': settings.get('maxLargeHallShifts', '2')
    }

    # --- الحل المبدئي والتكلفة الأولية ---
    current_solution = copy.deepcopy(initial_schedule)
    best_solution_so_far = copy.deepcopy(current_solution)

    # استخدام دالة التكلفة الشاملة
    current_cost, shortage, violations = calculate_cost(current_solution, all_professors, settings, duty_patterns, date_map)
    best_cost_so_far = current_cost
    log_q.put(f"... [VNS] التكلفة الأولية = {current_cost:.2f} (نقص={shortage}, خروقات={violations})")

    # --- حلقة VNS الرئيسية ---
    i = 0
    while i < iterations:
        # إيقاف الخوارزمية إذا طلب المستخدم ذلك
        if settings.get('should_stop_event', threading.Event()).is_set():
            log_q.put("... [VNS] تم استلام طلب إيقاف.")
            break

        percent_complete = int(((i + 1) / iterations) * 100)
        log_q.put(f"PROGRESS:{percent_complete}")
        
        k = 1
        while k <= k_max:
            # --- 1. مرحلة الهز (Shaking) باستخدام "التدمير والإصلاح" ---
            shaken_solution = copy.deepcopy(current_solution)
            
            # تحديد كل مهام الحراسة المتاحة للتدمير (غير المقفلة)
            duties_to_destroy = []
            for day in shaken_solution.values():
                for slot in day.values():
                    for exam in slot:
                        for g_idx, guard in enumerate(exam.get('guards', [])):
                            if guard != "**نقص**" and (exam.get('uuid'), guard) not in locked_guards:
                                duties_to_destroy.append({'exam': exam, 'guard_index': g_idx})

            if not duties_to_destroy: break # لا يوجد ما يمكن تدميره
            
            random.shuffle(duties_to_destroy)
            
            # تدمير عدد 'k' من المهام
            destroyed_slots_info = []
            for j in range(min(k, len(duties_to_destroy))):
                duty_info = duties_to_destroy[j]
                exam = duty_info['exam']
                g_idx = duty_info['guard_index']
                # وضع علامة "**نقص**" مكان الحارس المدمر
                exam['guards'][g_idx] = "**نقص**"
                destroyed_slots_info.append({'exam': exam, 'index_to_fill': g_idx})

            # --- إصلاح الحل المدمر ---
            # إعادة بناء سجلات الحراس الحالية بعد التدمير
            prof_assignments = defaultdict(list)
            prof_large_counts = defaultdict(int)
            all_exams_in_shaken = [exam for day in shaken_solution.values() for slot in day.values() for exam in slot]
            for exam in all_exams_in_shaken:
                is_large = any(h['type'] == 'كبيرة' for h in exam.get('halls', []))
                for guard in exam.get('guards', []):
                    if guard != "**نقص**":
                        prof_assignments[guard].append(exam)
                        if is_large:
                            prof_large_counts[guard] += 1
            
            # ملء الخانات المدمرة
            for repair_info in destroyed_slots_info:
                exam_to_repair = repair_info['exam']
                shuffled_profs = list(all_professors)
                random.shuffle(shuffled_profs)
                best_prof_found = None
                
                for prof in shuffled_profs:
                    if prof in exam_to_repair.get('guards', []): continue
                    
                    if is_assignment_valid(prof, exam_to_repair, prof_assignments, prof_large_counts, settings_for_validation, date_map):
                        best_prof_found = prof
                        break

                if best_prof_found:
                    exam_to_repair['guards'][repair_info['index_to_fill']] = best_prof_found
                    prof_assignments[best_prof_found].append(exam_to_repair)
                    if any(h['type'] == 'كبيرة' for h in exam_to_repair.get('halls',[])):
                        prof_large_counts[best_prof_found] += 1

            # --- 2. مرحلة البحث المحلي (Local Search) ---
            local_search_solution, _, _, _ = run_post_processing_swaps(
                shaken_solution, defaultdict(list), defaultdict(float), defaultdict(int), 
                settings, all_professors, date_map, local_search_swaps, locked_guards
            )

            # --- 3. مرحلة التحديث (Move or not) ---
            new_cost, new_shortage, new_violations = calculate_cost(local_search_solution, all_professors, settings, duty_patterns, date_map)

            if new_cost < current_cost:
                current_solution = local_search_solution
                current_cost = new_cost
                log_q.put(f"... [VNS] دورة {i+1}, k={k}: تم العثور على حل أفضل بتكلفة = {current_cost:.2f} (نقص={new_shortage}, خروقات={new_violations})")
                
                if new_cost < best_cost_so_far:
                    best_cost_so_far = new_cost
                    best_solution_so_far = copy.deepcopy(current_solution)

                k = 1 # العودة إلى الجوار الأول بعد إيجاد تحسين
            else:
                k += 1 # الانتقال إلى جوار أكبر (هزة أعنف)
        
        i += 1

    log_q.put(f"✓ انتهى VNS بأفضل تكلفة: {best_cost_so_far:.2f}")
    
    # إعادة بناء البيانات النهائية من أفضل حل تم العثور عليه
    final_assignments = defaultdict(list)
    final_workload = defaultdict(float)
    final_large_counts = defaultdict(int)
    for day in best_solution_so_far.values():
        for slot in day.values():
            for exam in slot:
                 is_large = any(h['type'] == 'كبيرة' for h in exam.get('halls', []))
                 duty_weight = float(settings.get('largeHallWeight', 3.0)) if is_large else float(settings.get('otherHallWeight', 1.0))
                 for guard in exam.get('guards',[]):
                     if guard != "**نقص**":
                         final_assignments[guard].append(exam)
                         final_workload[guard] += duty_weight
                         if is_large:
                             final_large_counts[guard] += 1

    return best_solution_so_far, final_assignments, final_workload, final_large_counts

# ================== نهاية الكود الجديد والمصحح بالكامل ==================




# =====================================================================================
# --- START: FINAL COMPLETE GENETIC ALGORITHM (WITH ALL CONSTRAINTS) ---
# =====================================================================================
import uuid

def run_genetic_algorithm(fixed_subject_schedule, settings, all_professors, assignments, all_levels_list, all_halls, exam_schedule_settings, all_subjects, level_hall_assignments, date_map, log_queue):
    """
    النسخة النهائية والمكتملة من الخوارزمية الجينية مع جميع القيود الصارمة.
    """
    # --- استخلاص الإعدادات ---
    pop_size = int(settings.get('geneticPopulation', 100))
    num_generations = int(settings.get('geneticGenerations', 500))
    crossover_rate = 0.8
    mutation_rate = float(settings.get('geneticMutation', 0.15))
    elitism_count = int(settings.get('geneticElitism', 4))
    
    # --- 1. إعطاء كل امتحان مُعرّف فريد (UUID) وتحضير خانات الحراسة ---
    schedule_with_ids = copy.deepcopy(fixed_subject_schedule)
    all_exams_flat = [exam for slots in schedule_with_ids.values() for exams in slots.values() for exam in exams]
    for exam in all_exams_flat:
        exam['uuid'] = str(uuid.uuid4())

    duty_slots = []
    guards_large_hall = int(settings.get('guardsLargeHall', 4))
    guards_medium_hall = int(settings.get('guardsMediumHall', 2))
    guards_small_hall = int(settings.get('guardsSmallHall', 1))

    for exam in all_exams_flat:
        num_guards_needed = (sum(guards_large_hall for h in exam.get('halls', []) if h.get('type') == 'كبيرة') +
                           sum(guards_medium_hall for h in exam.get('halls', []) if h.get('type') == 'متوسطة') +
                           sum(guards_small_hall for h in exam.get('halls', []) if h.get('type') == 'صغيرة'))
        for _ in range(num_guards_needed):
            duty_slots.append(exam)

    if not duty_slots:
        return fixed_subject_schedule, True

    # --- 2. تعريف دوال مساعدة ---
    
    def build_schedule_from_chromosome(chromosome):
        schedule_copy = copy.deepcopy(schedule_with_ids)
        exam_map = {ex['uuid']: ex for slots in schedule_copy.values() for exams in slots.values() for ex in exams}
        for ex in exam_map.values():
            ex['guards'] = []

        for i, guard in enumerate(chromosome):
            exam_ref_uuid = duty_slots[i]['uuid']
            exam_in_copy = exam_map.get(exam_ref_uuid)
            if exam_in_copy:
                exam_in_copy['guards'].append(guard)
        return schedule_copy

    def create_random_chromosome():
        chromosome = [None] * len(duty_slots)
        prof_assignments = {prof: [] for prof in all_professors}
        prof_large_counts = defaultdict(int)
        
        unavailable_days = settings.get('unavailableDays', {})
        duty_patterns = settings.get('dutyPatterns', {})
        max_shifts = int(settings.get('maxShifts', '0')) if settings.get('maxShifts', '0') != '0' else float('inf')
        max_large_hall_shifts = int(settings.get('maxLargeHallShifts', '2')) if settings.get('maxLargeHallShifts', '2') != '0' else float('inf')
        assign_owner_as_guard = settings.get('assignOwnerAsGuard', False)

        # ## <<< بداية: إضافة منطق تعيين أستاذ المادة الإلزامي >>>
        if assign_owner_as_guard:
            subject_owners = { (clean_string_for_matching(s['name']), clean_string_for_matching(s['level'])): clean_string_for_matching(prof) for prof, uids in assignments.items() for uid in uids for s in all_subjects if f"{s['name']} ({s['level']})" == uid }
            prof_last_exam = {}
            for exam in all_exams_flat:
                owner = subject_owners.get((clean_string_for_matching(exam['subject']), clean_string_for_matching(exam['level'])))
                if owner:
                    exam_date_time = (exam['date'], exam['time'])
                    if owner not in prof_last_exam or exam_date_time > prof_last_exam[owner]['datetime']:
                        prof_last_exam[owner] = {'exam': exam, 'datetime': exam_date_time}

            for owner, data in prof_last_exam.items():
                exam_to_assign = data['exam']
                
                # إيجاد خانة حراسة فارغة لهذا الامتحان
                duty_index_to_fill = -1
                for i, slot_exam in enumerate(duty_slots):
                    if slot_exam['uuid'] == exam_to_assign['uuid'] and chromosome[i] is None:
                        duty_index_to_fill = i
                        break
                
                if duty_index_to_fill != -1:
                    # التحقق من صلاحية هذا التعيين الإلزامي
                    is_large = any(h['type'] == 'كبيرة' for h in exam_to_assign.get('halls', []))
                    if exam_to_assign['date'] not in unavailable_days.get(owner, []) and \
                       len(prof_assignments[owner]) < max_shifts and \
                       (not is_large or prof_large_counts[owner] < max_large_hall_shifts):
                        
                        chromosome[duty_index_to_fill] = owner
                        prof_assignments[owner].append(exam_to_assign)
                        if is_large:
                            prof_large_counts[owner] += 1
        # ## <<< نهاية: منطق التعيين الإلزامي >>>

        shuffled_indices = list(range(len(duty_slots)))
        random.shuffle(shuffled_indices)

        for i in shuffled_indices:
            if chromosome[i] is not None: continue # تخطي الخانات التي تم ملؤها

            exam = duty_slots[i]
            is_large_exam = any(h['type'] == 'كبيرة' for h in exam.get('halls', []))
            
            shuffled_profs = list(all_professors)
            random.shuffle(shuffled_profs)

            assigned_prof = None
            for prof in shuffled_profs:
                # التحقق من أن الأستاذ غير معين بالفعل في هذه الخانة
                if any(chromosome[j] == prof for j, slot in enumerate(duty_slots) if slot['uuid'] == exam['uuid']):
                    continue

                is_busy = any(a_exam['date'] == exam['date'] and a_exam['time'] == exam['time'] for a_exam in prof_assignments[prof])
                if is_busy: continue
                if exam['date'] in unavailable_days.get(prof, []): continue
                if len(prof_assignments[prof]) >= max_shifts: continue
                if is_large_exam and prof_large_counts[prof] >= max_large_hall_shifts: continue

                prof_pattern = duty_patterns.get(prof, 'flexible_2_days')
                current_prof_dates = {e['date'] for e in prof_assignments.get(prof, [])}
                is_new_day = exam['date'] not in current_prof_dates
                
                is_pattern_violated = False
                if is_new_day:
                    num_current_days = len(current_prof_dates)
                    if prof_pattern == 'one_day_only' and num_current_days >= 1: is_pattern_violated = True
                    elif prof_pattern == 'flexible_2_days' and num_current_days >= 2: is_pattern_violated = True
                    elif prof_pattern == 'flexible_3_days' and num_current_days >= 3: is_pattern_violated = True
                    elif prof_pattern == 'consecutive_strict':
                        if num_current_days >= 2: is_pattern_violated = True
                        elif num_current_days == 1:
                            existing_day_idx = date_map.get(list(current_prof_dates)[0])
                            new_day_idx = date_map.get(exam['date'])
                            if existing_day_idx is None or new_day_idx is None or abs(new_day_idx - existing_day_idx) != 1:
                                is_pattern_violated = True
                if is_pattern_violated: continue
                
                assigned_prof = prof
                break
            
            if assigned_prof:
                chromosome[i] = assigned_prof
                prof_assignments[assigned_prof].append(exam)
                if is_large_exam:
                    prof_large_counts[assigned_prof] += 1
            else:
                chromosome[i] = "**نقص**"
                
        return chromosome

    def calculate_fitness(chromosome):
        if "**نقص**" in chromosome:
            return 0.0
        schedule = build_schedule_from_chromosome(chromosome)
        if not is_schedule_valid(schedule, settings, all_professors, settings.get('dutyPatterns', {}), date_map):
            return 0.0
        return calculate_schedule_balance_score(schedule, all_professors, settings, len(all_professors))

    def crossover(parent1, parent2):
        point = random.randint(1, len(parent1) - 1)
        return parent1[:point] + parent2[point:], parent2[:point] + parent1[point:]

    def mutation(chromosome):
        if len(chromosome) < 2: return chromosome
        
        # اختيار خانتين مختلفتين للتبديل
        idx1, idx2 = random.sample(range(len(chromosome)), 2)
        
        # عمل نسخة من الكروموسوم لتجنب التعديل المباشر
        mutated_chromosome = list(chromosome)
        mutated_chromosome[idx1], mutated_chromosome[idx2] = mutated_chromosome[idx2], mutated_chromosome[idx1]
        
        # التحقق من صلاحية الطفرة
        temp_schedule = build_schedule_from_chromosome(mutated_chromosome)
        if is_schedule_valid(temp_schedule, settings, all_professors, settings.get('dutyPatterns', {}), date_map):
            return mutated_chromosome # إرجاع الحل الجديد الصالح
        return chromosome # التراجع عن الطفرة إذا كانت ستنتج حلاً غير صالح

    # --- 3. بدء الخوارزمية ---
    log_queue.put(">>> [Genetic Alg v2] بدء الخوارزمية الجينية المحسنة...")
    log_queue.put("... بناء المجتمع الأولي على جدول المواد الثابت.")
    
    try:
        population = [create_random_chromosome() for _ in range(pop_size)]
    except Exception as e:
        import traceback
        log_queue.put(f"!!! Error during initial population creation: {e}")
        log_queue.put(traceback.format_exc())
        return fixed_subject_schedule, False
    
    best_chromosome, best_fitness = None, -1.0

    for gen in range(num_generations):
        # =================> بداية الإضافة الجديدة <=================
        # إرسال التقدم بناءً على الجيل الحالي
        percent_complete = int(((gen + 1) / num_generations) * 100)
        log_queue.put(f"PROGRESS:{percent_complete}")
        # =================> نهاية الإضافة الجديدة <=================
        population_with_fitness = [(chrom, calculate_fitness(chrom)) for chrom in population]
        
        valid_population_with_fitness = [item for item in population_with_fitness if item[1] > 0]
        if not valid_population_with_fitness:
             log_queue.put(f"... [Generation {gen+1}] Population collapsed. Rebuilding...")
             population = [create_random_chromosome() for _ in range(pop_size)]
             continue

        current_best_chrom, current_best_fitness = max(valid_population_with_fitness, key=lambda item: item[1])
        if current_best_fitness > best_fitness:
            best_fitness, best_chromosome = current_best_fitness, current_best_chrom
            log_queue.put(f"... [Generation {gen+1}/{num_generations}] New best fitness: {best_fitness:.2f}%")

        if best_fitness >= 99.9:
            log_queue.put("✓ تم العثور على حل مثالي. إنهاء البحث.")
            break
        
        valid_population_with_fitness.sort(key=lambda item: item[1], reverse=True)
        next_generation = [item[0] for item in valid_population_with_fitness[:elitism_count]]
        
        while len(next_generation) < pop_size:
            parents = random.choices(valid_population_with_fitness, weights=[f for _, f in valid_population_with_fitness], k=2)
            p1, p2 = parents[0][0], parents[1][0]

            if random.random() < crossover_rate:
                c1, c2 = crossover(p1, p2)
                next_generation.append(mutation(c1) if random.random() < mutation_rate else c1)
                if len(next_generation) < pop_size:
                    next_generation.append(mutation(c2) if random.random() < mutation_rate else c2)
            else:
                next_generation.append(p1)
                if len(next_generation) < pop_size:
                    next_generation.append(p2)
        population = next_generation

    log_queue.put(f">>> [Genetic Alg v2] Finished. Best fitness score: {best_fitness:.2f}%")
    
    if not best_chromosome:
        log_queue.put("!!! لم تتمكن الخوارزمية الجينية من إيجاد حل صالح.")
        any_valid = [chrom for chrom, fit in population_with_fitness if fit > 0]
        if any_valid:
            best_chromosome = any_valid[0]
        else:
            return fixed_subject_schedule, False

    final_schedule = build_schedule_from_chromosome(best_chromosome)
    return final_schedule, True

# --- END: FINAL COMPLETE GENETIC ALGORITHM ---
# =====================================================================================


# الدالة بعد التعديل (استبدل الدالة بالكامل)
# في ملف app.py، قم باستبدال هذه الدالة بالكامل

# في ملف app.py، قم باستبدال هذه الدالة بالكامل
# في ملف app.py، قم باستبدال هذه الدالة بالكامل
def complete_schedule_with_guards(subject_schedule, settings, all_professors, assignments, all_levels_list, duty_patterns, date_map, all_subjects, locked_guards=set()):
    """
    النسخة النهائية والمبسطة: تستخدم دالة التحقق المركزية is_assignment_valid.
    """
    schedule = copy.deepcopy(subject_schedule)
    
    # --- تطبيق التعيينات المقفلة ---
    exam_map_by_uuid = {exam['uuid']: exam for day in schedule.values() for slot in day.values() for exam in slot if 'uuid' in exam}
    for exam_uuid, prof in locked_guards:
        if exam_uuid in exam_map_by_uuid:
            exam_obj = exam_map_by_uuid[exam_uuid]
            if 'guards' not in exam_obj: exam_obj['guards'] = []
            if prof not in exam_obj['guards']: exam_obj['guards'].append(prof)
    
    # --- بناء سجلات الحراس بناءً على الحالة الحالية ---
    prof_assignments = defaultdict(list)
    prof_large_counts = defaultdict(int)
    all_scheduled_exams_flat = [exam for day in schedule.values() for slot in day.values() for exam in slot]

    for exam in all_scheduled_exams_flat:
        is_large = any(h['type'] == 'كبيرة' for h in exam.get('halls', []))
        for guard in exam.get('guards', []):
            if guard in all_professors:
                prof_assignments[guard].append(exam)
                if is_large: prof_large_counts[guard] += 1
    
    # --- ملء الخانات الشاغرة ---
    guards_large_hall = int(settings.get('guardsLargeHall', 4))
    guards_medium_hall = int(settings.get('guardsMediumHall', 2))
    guards_small_hall = int(settings.get('guardsSmallHall', 1))

    duties_to_fill = []
    for exam in all_scheduled_exams_flat:
        num_needed = (sum(guards_large_hall for h in exam.get('halls',[]) if h.get('type')=='كبيرة') +
                      sum(guards_medium_hall for h in exam.get('halls',[]) if h.get('type')=='متوسطة') +
                      sum(guards_small_hall for h in exam.get('halls',[]) if h.get('type')=='صغيرة'))
        num_to_add = num_needed - len(exam.get('guards', []))
        for _ in range(num_to_add):
            duties_to_fill.append(exam)
    random.shuffle(duties_to_fill)

    # إعداد قاموس الإعدادات مرة واحدة خارج الحلقة
    settings_for_validation = {
        'dutyPatterns': settings.get('dutyPatterns', {}),
        'unavailableDays': settings.get('unavailableDays', {}),
        'maxShifts': settings.get('maxShifts', '0'),
        'maxLargeHallShifts': settings.get('maxLargeHallShifts', '2')
    }

    for duty_exam in duties_to_fill:
        shuffled_profs = list(all_professors)
        random.shuffle(shuffled_profs)
        
        assigned = False
        for prof in shuffled_profs:
            if prof in duty_exam.get('guards', []): continue
            
            # ==> هذا هو التعديل الجوهري: استخدام الدالة المركزية <==
            if is_assignment_valid(prof, duty_exam, prof_assignments, prof_large_counts, settings_for_validation, date_map):
                duty_exam['guards'].append(prof)
                prof_assignments[prof].append(duty_exam)
                if any(h['type'] == 'كبيرة' for h in duty_exam['halls']):
                    prof_large_counts[prof] += 1
                assigned = True
                break
        
        if not assigned:
             duty_exam['guards'].append("**نقص**")

    return schedule

# في ملف app.py، استبدل الدالة بالكامل بهذه النسخة النهائية (الإصدار التاسع - المستقر)
def run_constraint_solver(original_schedule, settings, all_professors, assignments, all_levels_list, subject_owners, last_day_restriction, sorted_dates, duty_patterns, date_map, log_q):
    """
    النسخة المحدثة: مع إضافة قيد "أزواج الأساتذة" الرياضي.
    """
    model = cp_model.CpModel()
    all_scheduled_exams = [exam for date_exams in original_schedule.values() for time_exams in date_exams.values() for exam in time_exams]

    # --- استخلاص الإعدادات ---
    large_hall_weight = int(settings.get('largeHallWeight', 3))
    other_hall_weight = int(settings.get('otherHallWeight', 1))
    max_shifts = int(settings.get('maxShifts', '0')) if settings.get('maxShifts', '0') != '0' else float('inf')
    max_large_hall_shifts = int(settings.get('maxLargeHallShifts', '2')) if settings.get('maxLargeHallShifts', '2') != '0' else float('inf')
    enable_custom_targets = settings.get('enableCustomTargets', False)
    custom_target_patterns = settings.get('customTargetPatterns', [])
    solver_timelimit = int(settings.get('solverTimelimit', 30))
    prof_map = {name: i for i, name in enumerate(all_professors)}
    num_professors = len(all_professors)
    guards_large_hall = int(settings.get('guardsLargeHall', 4))
    guards_medium_hall = int(settings.get('guardsMediumHall', 2))
    guards_small_hall = int(settings.get('guardsSmallHall', 1))
    unavailable_days = settings.get('unavailableDays', {})
    
    # --- بناء كل المتغيرات والقيود الصارمة (بما في ذلك قيد الأزواج) ---
    duties = []
    for exam in all_scheduled_exams:
        num_guards_needed = sum(guards_large_hall for h in exam.get('halls', []) if h.get('type') == 'كبيرة') + sum(guards_medium_hall for h in exam.get('halls', []) if h.get('type') == 'متوسطة') + sum(guards_small_hall for h in exam.get('halls', []) if h.get('type') == 'صغيرة')
        for _ in range(num_guards_needed): duties.append({'exam': exam, 'is_large': any(h['type'] == 'كبيرة' for h in exam['halls'])})

    x = {(p_idx, d_idx): model.NewBoolVar(f'x_{p_idx}_{d_idx}') for p_idx in range(num_professors) for d_idx in range(len(duties))}
    for d_idx in range(len(duties)): model.AddExactlyOne(x[p_idx, d_idx] for p_idx in range(num_professors))
    
    slots = defaultdict(list)
    for d_idx, duty in enumerate(duties): slots[(duty['exam']['date'], duty['exam']['time'])].append(d_idx)
    for p_idx in range(num_professors):
        for slot_duties in slots.values(): model.Add(sum(x[p_idx, d] for d in slot_duties) <= 1)
    
    if max_shifts != float('inf'):
        for p_idx in range(num_professors): model.Add(sum(x[p_idx, d_idx] for d_idx in range(len(duties))) <= max_shifts)
    if max_large_hall_shifts != float('inf'):
        large_duty_indices = [d_idx for d_idx, duty in enumerate(duties) if duty['is_large']]
        for p_idx in range(num_professors): model.Add(sum(x[p_idx, d_idx] for d_idx in large_duty_indices) <= max_large_hall_shifts)
    for prof, dates in unavailable_days.items():
        if prof in prof_map:
            p_idx = prof_map[prof]
            for d_idx, duty in enumerate(duties):
                if duty['exam']['date'] in dates: model.Add(x[p_idx, d_idx] == 0)

    is_duty_day, prof_has_any_duty = {}, {}
    for p_idx in range(num_professors):
        prof_has_any_duty[p_idx] = model.NewBoolVar(f'prof_has_duty_{p_idx}')
        duties_for_prof = [x[p_idx, d_idx] for d_idx in range(len(duties))]
        model.Add(sum(duties_for_prof) > 0).OnlyEnforceIf(prof_has_any_duty[p_idx]); model.Add(sum(duties_for_prof) == 0).OnlyEnforceIf(prof_has_any_duty[p_idx].Not())
        for day_idx in range(len(sorted_dates)):
            is_duty_day[p_idx, day_idx] = model.NewBoolVar(f'is_duty_day_{p_idx}_{day_idx}')
            duties_in_this_day = [x[p_idx, d_idx] for d_idx, duty in enumerate(duties) if date_map.get(duty['exam']['date']) == day_idx]
            if duties_in_this_day:
                model.AddBoolOr(duties_in_this_day).OnlyEnforceIf(is_duty_day[p_idx, day_idx]); model.Add(sum(duties_in_this_day) == 0).OnlyEnforceIf(is_duty_day[p_idx, day_idx].Not())
            else: model.Add(is_duty_day[p_idx, day_idx] == False)

    for prof_name, pattern in duty_patterns.items():
        if prof_name in prof_map:
            p_idx = prof_map[prof_name]
            num_unique_duty_days = sum(is_duty_day[p_idx, day_idx] for day_idx in range(len(sorted_dates)))
            if pattern == 'consecutive_strict':
                model.Add(num_unique_duty_days == 2).OnlyEnforceIf(prof_has_any_duty[p_idx])
                is_start_of_consecutive_pair = [model.NewBoolVar(f'is_start_{p_idx}_{d}') for d in range(len(sorted_dates) - 1)]
                for day_idx in range(len(sorted_dates) - 1):
                    model.AddBoolAnd([is_duty_day[p_idx, day_idx], is_duty_day[p_idx, day_idx+1]]).OnlyEnforceIf(is_start_of_consecutive_pair[day_idx])
                    model.AddBoolOr([is_duty_day[p_idx, day_idx].Not(), is_duty_day[p_idx, day_idx+1].Not()]).OnlyEnforceIf(is_start_of_consecutive_pair[day_idx].Not())
                model.Add(sum(is_start_of_consecutive_pair) == 1).OnlyEnforceIf(prof_has_any_duty[p_idx])
            elif pattern == 'one_day_only':
                model.Add(num_unique_duty_days <= 1).OnlyEnforceIf(prof_has_any_duty[p_idx])
            elif pattern == 'flexible_2_days': model.Add(num_unique_duty_days == 2).OnlyEnforceIf(prof_has_any_duty[p_idx])
            elif pattern == 'flexible_3_days':
                model.Add(num_unique_duty_days >= 2).OnlyEnforceIf(prof_has_any_duty[p_idx]); model.Add(num_unique_duty_days <= 3).OnlyEnforceIf(prof_has_any_duty[p_idx])
    
    # --- بداية: إضافة قيد أزواج الأساتذة ---
    professor_pairs = settings.get('professorPartnerships', []) # تم تغيير اسم المفتاح
    if professor_pairs:
        for pair in professor_pairs:
            if len(pair) == 2 and pair[0] in prof_map and pair[1] in prof_map:
                prof1_idx = prof_map[pair[0]]
                prof2_idx = prof_map[pair[1]]
                # لكل يوم في الجدول، يجب أن يكونا إما كلاهما يعمل أو كلاهما لا يعمل
                for day_idx in range(len(sorted_dates)):
                    model.Add(is_duty_day[prof1_idx, day_idx] == is_duty_day[prof2_idx, day_idx])
    # --- نهاية: إضافة قيد أزواج الأساتذة ---

    # --- الهدف (Objective) ---
    prof_large_duties = [model.NewIntVar(0, 100, f'large_{p}') for p in range(num_professors)]
    prof_other_duties = [model.NewIntVar(0, 100, f'other_{p}') for p in range(num_professors)]
    # ... (بقية الدالة تبقى كما هي دون تغيير) ...
    # ...
    #
    for p_idx in range(num_professors):
        model.Add(prof_large_duties[p_idx] == sum(x[p_idx, d_idx] for d_idx, duty in enumerate(duties) if duty['is_large']))
        model.Add(prof_other_duties[p_idx] == sum(x[p_idx, d_idx] for d_idx, duty in enumerate(duties) if not duty['is_large']))

    if enable_custom_targets and custom_target_patterns:
        log_q.put("... استخدام نموذج تقليل الانحراف مع معاقبة الأنماط غير المستهدفة")
        target_counts = Counter((p['large'], p['other']) for p in custom_target_patterns for _ in range(p.get('count', 0)))
        actual_counts = {}
        for l, o in target_counts.keys():
            prof_matches_pattern = [model.NewBoolVar(f'p_{p}_m_{l}_{o}') for p in range(num_professors)]
            for p_idx in range(num_professors):
                model.Add(prof_large_duties[p_idx] == l).OnlyEnforceIf(prof_matches_pattern[p_idx])
                model.Add(prof_other_duties[p_idx] == o).OnlyEnforceIf(prof_matches_pattern[p_idx])
            actual_counts[(l,o)] = sum(prof_matches_pattern)
        deviation_terms = []
        for pattern, target_count in target_counts.items():
            deviation = model.NewIntVar(0, num_professors, f'dev_{pattern}')
            model.AddAbsEquality(deviation, actual_counts[pattern] - target_count)
            deviation_terms.append(deviation)
        total_deviation = sum(deviation_terms)
        num_untracked = model.NewIntVar(0, num_professors, 'num_untracked')
        total_tracked_profs = sum(actual_counts.values())
        model.Add(num_untracked == num_professors - total_tracked_profs)
        model.Minimize(total_deviation + 10 * num_untracked)
    else:
        log_q.put("... استخدام دالة الهدف الافتراضية (الموازنة العامة)")
        prof_workload = [model.NewIntVar(0, 1000, f'workload_{p}') for p in range(num_professors)]
        for p_idx in range(num_professors):
            model.Add(prof_workload[p_idx] == prof_large_duties[p_idx] * large_hall_weight + prof_other_duties[p_idx] * other_hall_weight)
        min_w, max_w = model.NewIntVar(0, 1000, 'min_w'), model.NewIntVar(0, 1000, 'max_w')
        if prof_workload:
            model.AddMinEquality(min_w, prof_workload); model.AddMaxEquality(max_w, prof_workload)
        model.Minimize(max_w - min_w)

    # --- حل النموذج ---
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(solver_timelimit)
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        log_q.put("✓ تم العثور على حل باستخدام البرمجة بالقيود.")
        new_schedule_list = copy.deepcopy(all_scheduled_exams)
        for exam in new_schedule_list: exam['guards'] = []
        exam_map_by_props = {(e['subject'], e['level'], e['date'], e['time']): e for e in new_schedule_list}
        for p_idx in range(num_professors):
            for d_idx, duty in enumerate(duties):
                if solver.Value(x[p_idx, d_idx]):
                    exam_props = duty['exam']
                    exam_key = (exam_props['subject'], exam_props['level'], exam_props['date'], exam_props['time'])
                    if exam_key in exam_map_by_props: exam_map_by_props[exam_key]['guards'].append(all_professors[p_idx])
        final_schedule = defaultdict(lambda: defaultdict(list))
        for exam in new_schedule_list: final_schedule[exam['date']][exam['time']].append(exam)
        return final_schedule, True
    else:
        log_q.put("✗ فشل! لم يتم العثور على أي حل صالح. هذا يؤكد وجود تضارب في القيود نفسها.")
        return original_schedule, False

# ================== الواجهة الرئيسية لتشغيل الخوارزمية ==================
def _run_schedule_logic_in_background(settings, log_q):
    try:
        # --- تحميل البيانات من قاعدة البيانات ---
        conn = get_db_connection()
        all_levels_list = [row['name'] for row in conn.execute("SELECT name FROM levels").fetchall()]
        all_professors = [row['name'] for row in conn.execute("SELECT name FROM professors").fetchall()]
        num_professors = len(all_professors)
        all_subjects_rows = conn.execute("SELECT s.name, l.name as level FROM subjects s JOIN levels l ON s.level_id = l.id").fetchall()
        all_subjects = [dict(row) for row in all_subjects_rows]
        all_halls = [dict(row) for row in conn.execute("SELECT name, type FROM halls").fetchall()]
        assignments_rows = conn.execute('SELECT p.name as prof_name, s.name as subj_name, l.name as level_name FROM assignments a JOIN professors p ON a.professor_id = p.id JOIN subjects s ON a.subject_id = s.id JOIN levels l ON s.level_id = l.id').fetchall()
        assignments = defaultdict(list)
        for row in assignments_rows:
            assignments[row['prof_name']].append(f"{row['subj_name']} ({row['level_name']})")
        conn.close()
        
        intensive_search = settings.get('intensiveSearch', False)
        try:
            num_iterations = int(settings.get('iterations', '200'))
        except (ValueError, TypeError):
            num_iterations = 200
        if not intensive_search:
            num_iterations = 1
        
        # --- إضافة جديدة: تعريف مفتاح جديد في best_result ---
        best_result = {'schedule': None, 'failures': [], 'scheduling_report': [], 'unfilled_slots': float('inf'), 'unscheduled_subjects': [], 'detailed_error': None, 'prof_report': [], 'chart_data': {}, 'balance_report': {}, 'stats_dashboard': {}}
        
        
        exam_schedule_settings = copy.deepcopy(settings.get('examSchedule', {}))
        
        balancing_strategy = settings.get('balancingStrategy', 'advanced')
        swap_attempts = int(settings.get('swapAttempts', 50))
        polishing_swaps = int(settings.get('polishingSwaps', 15))
        annealing_temp = float(settings.get('annealingTemp', 1000.0))
        annealing_cooling = float(settings.get('annealingCooling', 0.995))
        annealing_iterations = int(settings.get('annealingIterations', 1000))
        solver_timelimit = int(settings.get('solverTimelimit', 30))
        
        level_hall_assignments = settings.get('levelHallAssignments', {})
        duty_patterns = settings.get('dutyPatterns', {})
        assign_owner_as_guard = settings.get('assignOwnerAsGuard', False)
        max_shifts_str = settings.get('maxShifts', '0')
        max_shifts = int(max_shifts_str) if max_shifts_str != '0' else float('inf')
        last_day_restriction = settings.get('lastDayRestriction', 'none')
        unavailable_days = settings.get('unavailableDays', {})
        max_large_hall_shifts_str = settings.get('maxLargeHallShifts', '2')
        max_large_hall_shifts = int(max_large_hall_shifts_str) if max_large_hall_shifts_str != '0' else float('inf')
        enable_custom_targets = settings.get('enableCustomTargets', False)
        custom_target_patterns = settings.get('customTargetPatterns', [])
        large_hall_weight = float(settings.get('largeHallWeight', 3.0))
        other_hall_weight = float(settings.get('otherHallWeight', 1.0))
        guards_large_hall_val = int(settings.get('guardsLargeHall', 4))
        guards_medium_hall_val = int(settings.get('guardsMediumHall', 2))
        guards_small_hall_val = int(settings.get('guardsSmallHall', 1))

        sorted_dates = sorted(exam_schedule_settings.keys())
        if not sorted_dates: 
            log_q.put("خطأ: لا توجد أيام امتحانات محددة.")
            log_q.put("DONE" + json.dumps({"success": False, "message": "لا توجد أيام امتحانات محددة."}))
            return

        date_map = {date: i for i, date in enumerate(sorted_dates)}
        
        last_exam_day = sorted_dates[-1] if sorted_dates else None
        restricted_slots_on_last_day = []
        if last_exam_day and last_day_restriction != 'none':
            try:
                last_day_all_slots = sorted(exam_schedule_settings.get(last_exam_day, []), key=lambda x: x['time'])
                num_to_restrict = int(last_day_restriction.split('_')[1])
                restricted_slots_on_last_day = [s['time'] for s in last_day_all_slots[-num_to_restrict:]]
                log_q.put(f"INFO: Last day restriction '{last_day_restriction}' applied. Restricted slots: {restricted_slots_on_last_day}")
            except (ValueError, IndexError):
                log_q.put(f"WARNING: Could not parse last_day_restriction: {last_day_restriction}")


        prof_map = {name: i for i, name in enumerate(all_professors)}
        subject_owners = { (clean_string_for_matching(s['name']), clean_string_for_matching(s['level'])): clean_string_for_matching(prof) for prof, uids in assignments.items() for uid in uids for s in all_subjects if f"{s['name']} ({s['level']})" == uid }
        
        for i in range(num_iterations):
            # =================> بداية التعديل <=================
            # إرسال التقدم فقط إذا كانت الخوارزمية بسيطة
            if balancing_strategy not in ['tabu_search', 'genetic', 'annealing', 'constraint_solver']:
                percent_complete = int(((i + 1) / num_iterations) * 100)
                log_q.put(f"PROGRESS:{percent_complete}")
            # =================> نهاية التعديل <=================
            log_q.put(f">>> [Iteration {i+1}/{num_iterations}] بدء محاولة جديدة...")

            current_schedule = defaultdict(lambda: defaultdict(list))
            current_all_subjects_to_schedule = {(clean_string_for_matching(s['name']), clean_string_for_matching(s['level'])) for s in all_subjects}
            current_prof_assignments = defaultdict(list)
            current_prof_large_counts = defaultdict(int) 
            current_prof_workload_score = defaultdict(float)
            iteration_error = None
            level_time_group = {}

            def schedule_exam_internal(subject_tuple_param, date, time, available_halls_in_slot):
                subject_name, level_key = subject_tuple_param
                level_name_found = next((lvl for lvl in all_levels_list if clean_string_for_matching(lvl) == level_key), level_key)
                halls_for_level_names = set(level_hall_assignments.get(level_name_found, []))
                if not halls_for_level_names: return False
                halls_to_use_names = halls_for_level_names.intersection(available_halls_in_slot)
                if len(halls_to_use_names) != len(halls_for_level_names): return False
                halls_details = [h for h in all_halls if h['name'] in halls_to_use_names]
                exam = {"date": date, "time": time, "subject": subject_name, "level": level_name_found, "professor": subject_owners.get(subject_tuple_param, "غير محدد"), "halls": halls_details, "guards": []}
                current_schedule[date][time].append(exam)
                if subject_tuple_param in current_all_subjects_to_schedule: current_all_subjects_to_schedule.remove(subject_tuple_param)
                for hall_name in halls_for_level_names:
                    if hall_name in available_halls_in_slot: available_halls_in_slot.remove(hall_name)
                return True

            def find_best_candidate(exam, current_pa, current_plc, current_pws):
                nonlocal iteration_error
                date, time = exam['date'], exam['time']
                
                if date == last_exam_day and time in restricted_slots_on_last_day:
                    return None 

                is_large_hall_exam = any(h['type'] == 'كبيرة' for h in exam['halls'])
                candidates = []
                
                # تجميع الإعدادات في قاموس واحد لتمريرها
                settings_for_validation = {
                    'dutyPatterns': duty_patterns,
                    'unavailableDays': unavailable_days,
                    'maxShifts': max_shifts,
                    'maxLargeHallShifts': max_large_hall_shifts
                }

                for prof in all_professors:
                    if prof in exam['guards']: continue

                    # الاستدعاء النظيف للدالة المركزية
                    if not is_assignment_valid(prof, exam, current_pa, current_plc, settings_for_validation, date_map):
                        continue
                    
                    # إذا كان التعيين صالحاً، أكمل حساب النقاط
                    score = 1000.0
                    if balancing_strategy == 'advanced':
                        num_duties = len(current_pa.get(prof, []))
                        penalty = (num_duties ** 2) * 20
                        score -= penalty
                    else:
                        score -= current_pws.get(prof, 0) * 25
                    
                    if is_large_hall_exam:
                        if current_plc.get(prof, 0) == 0: score += 15
                        else: score -= 30
                    
                    if not assign_owner_as_guard and prof == exam.get('professor'): score += 50
                    
                    # ==> تم حذف كتلة التحقق المكررة من هنا <==
                    
                    candidates.append((score + random.uniform(0, 1), prof))
                
                if not candidates:
                    if not iteration_error:
                        iteration_error = analyze_failure_reason(exam, all_professors, current_pa, unavailable_days, max_shifts, max_large_hall_shifts, current_plc, duty_patterns, date_map, current_schedule)
                    return None
                
                return max(candidates, key=lambda x: x[0])[1]

            def fill_slots(required_slots, current_pa, current_plc, current_pws):
                for exam_to_fill in required_slots:
                    best_prof = find_best_candidate(exam_to_fill, current_pa, current_plc, current_pws)
                    if best_prof:
                        exam_to_fill['guards'].append(best_prof)
                        current_pa[best_prof].append(exam_to_fill)
                        is_large_exam_for_fill = any(h['type'] == 'كبيرة' for h in exam_to_fill['halls'])
                        current_pws[best_prof] += large_hall_weight if is_large_exam_for_fill else other_hall_weight
                        if is_large_exam_for_fill: current_plc[best_prof] += 1
                    else:
                        exam_to_fill['guards'].append("**نقص**")
            
            shuffled_dates_for_primary = list(sorted_dates)
            random.shuffle(shuffled_dates_for_primary)
            for date in shuffled_dates_for_primary:
                slots_for_day = exam_schedule_settings.get(date, [])
                random.shuffle(slots_for_day)
                for slot in slots_for_day:
                    if slot.get('type') == 'primary':
                        time = slot['time']
                        
                        if date == last_exam_day and time in restricted_slots_on_last_day:
                            continue

                        available_halls_in_slot = {h['name'] for h in all_halls}
                        shuffled_levels = list(slot.get('levels', []))
                        random.shuffle(shuffled_levels)
                        for level in shuffled_levels:
                            if level not in level_time_group: level_time_group[level] = time
                            matching_subjects = [s for s in current_all_subjects_to_schedule if s[1] == clean_string_for_matching(level)]
                            if matching_subjects:
                                subject_to_schedule = random.choice(matching_subjects)
                                schedule_exam_internal(subject_to_schedule, date, time, available_halls_in_slot)
            
            if current_all_subjects_to_schedule:
                shuffled_dates_for_reserve = list(sorted_dates)
                random.shuffle(shuffled_dates_for_reserve)
                for date in shuffled_dates_for_reserve:
                    for slot in exam_schedule_settings.get(date, []):
                        if slot.get('type') == 'reserve':
                            time = slot['time']

                            if date == last_exam_day and time in restricted_slots_on_last_day:
                                continue
                            
                            scheduled_halls_in_slot = {h['name'] for exams in current_schedule.get(date,{}).get(time,[]) for h in exams.get('halls', [])}
                            available_halls_in_slot = {h['name'] for h in all_halls} - scheduled_halls_in_slot
                            groups_in_slot = {level_time_group.get(e['level']) for e in current_schedule.get(date,{}).get(time,[]) if e.get('level')}
                            groups_in_slot.discard(None)
                            shuffled_levels = list(slot.get('levels', []))
                            random.shuffle(shuffled_levels)
                            for level in shuffled_levels:
                                matching_subjects = [s for s in current_all_subjects_to_schedule if s[1] == clean_string_for_matching(level)]
                                if not matching_subjects: continue
                                
                                subject_to_schedule = random.choice(matching_subjects)
                                level_of_subject = next((lvl for lvl in all_levels_list if clean_string_for_matching(lvl) == subject_to_schedule[1]), None)
                                if not level_of_subject: continue
                                group_of_subject = level_time_group.get(level_of_subject)
                                if not groups_in_slot or group_of_subject in groups_in_slot or not group_of_subject:
                                    if schedule_exam_internal(subject_to_schedule, date, time, available_halls_in_slot):
                                        if group_of_subject: groups_in_slot.add(group_of_subject)
            
            # ===================================================================
            # --- START: المرحلة 1.5 - تحسين تجميع مواد الأساتذة (مشروط) ---
            # ===================================================================
            # يتم تشغيل هذه المرحلة فقط إذا قام المستخدم بتفعيل الخيار من الواجهة
            if settings.get('groupSubjects', False):
                current_schedule = run_subject_optimization_phase(
                    current_schedule, assignments, all_levels_list, subject_owners, settings, log_q
                )
            else:
                log_q.put(">>> تم تخطي المرحلة 1.5 بناءً على اختيار المستخدم.")
            # ===================================================================
            # --- END: المرحلة 1.5 ---
            # ===================================================================

            final_schedule_from_strategy = None
            strategy_success = False

            if balancing_strategy == 'genetic':
                log_q.put(">>> تشغيل خوارزمية الجينات...")
                final_schedule_from_strategy, strategy_success = run_genetic_algorithm(current_schedule, settings, all_professors, assignments, all_levels_list, all_halls, exam_schedule_settings, all_subjects, level_hall_assignments, date_map, log_q)
                
            elif balancing_strategy == 'constraint_solver':
                log_q.put(">>> تشغيل البرمجة بالقيود...")
                final_schedule_from_strategy, strategy_success = run_constraint_solver(current_schedule, settings, all_professors, assignments, all_levels_list, subject_owners, last_day_restriction, sorted_dates, duty_patterns, date_map, log_q)

            elif balancing_strategy == 'lns' or balancing_strategy == 'vns':
                log_q.put(f">>> بدء استراتيجية ({balancing_strategy.upper()}) مع مرحلة إحماء مكثفة...")

                # --- تعريف دالة التكلفة محليًا لإعادة استخدامها ---
                def calculate_cost(sch):
                    prof_stats = defaultdict(lambda: {'large': 0, 'other': 0})
                    for day in sch.values():
                        for slot in day.values():
                            for exam in slot:
                                guards_copy = [g for g in exam.get('guards', []) if g != "**نقص**"]
                                large_guards_needed = sum(guards_large_hall_val for h in exam.get('halls', []) if h.get('type') == 'كبيرة')
                                large_hall_guards, other_hall_guards = guards_copy[:large_guards_needed], guards_copy[large_guards_needed:]
                                for guard in large_hall_guards:
                                    if guard in all_professors: prof_stats[guard]['large'] += 1
                                for guard in other_hall_guards:
                                    if guard in all_professors: prof_stats[guard]['other'] += 1
                    
                    if enable_custom_targets and custom_target_patterns:
                        target_counts = Counter((p['large'], p['other']) for p in custom_target_patterns for _ in range(p.get('count', 0)))
                        actual_counts = Counter((s['large'], s['other']) for s in prof_stats.values())
                        total_deviation = sum(abs(actual_counts.get(p, 0) - target_counts.get(p, 0)) for p in set(target_counts.keys()) | set(actual_counts.keys()))
                        return total_deviation
                    else:
                        workload = {p: s['large'] * large_hall_weight + s['other'] * other_hall_weight for p, s in prof_stats.items()}
                        if not workload: return 0.0
                        workloads = list(workload.values())
                        return max(workloads) - min(workloads) if workloads else 0.0

                # --- مرحلة الإحماء: تشغيل "البحث المرحلي مع صقل" 10 مرات ---
                num_warmup_attempts = 10
                best_initial_solution = None
                best_initial_cost = float('inf')

                log_q.put(f"... بدء مرحلة الإحماء (سيتم تنفيذ {num_warmup_attempts} محاولات سريعة)...")
                for attempt in range(num_warmup_attempts):
                    # 1. نسخ الجدول الأساسي للمواد
                    schedule_for_attempt = copy.deepcopy(current_schedule)
                    all_exams_in_attempt = [exam for day in schedule_for_attempt.values() for exam_list in day.values() for exam in exam_list]
                    for exam in all_exams_in_attempt:
                        if 'uuid' not in exam: exam['uuid'] = str(uuid.uuid4())

                    # 2. ملء الحراس باستخدام الطريقة المرحلية
                    attempt_assignments = defaultdict(list)
                    attempt_large_counts = defaultdict(int)
                    attempt_workload = defaultdict(float)
                    
                    required_slots = []
                    for exam in all_exams_in_attempt:
                        num_needed = sum(guards_large_hall_val for h in exam.get('halls',[]) if h.get('type')=='كبيرة') + sum(guards_medium_hall_val for h in exam.get('halls',[]) if h.get('type')=='متوسطة') + sum(guards_small_hall_val for h in exam.get('halls',[]) if h.get('type')=='صغيرة')
                        for _ in range(num_needed - len(exam.get('guards',[]))): required_slots.append(exam)
                    
                    large_slots = [e for e in required_slots if any(h['type'] == 'كبيرة' for h in e['halls'])]
                    other_slots = [e for e in required_slots if not any(h['type'] == 'كبيرة' for h in e['halls'])]
                    fill_slots(large_slots, attempt_assignments, attempt_large_counts, attempt_workload)
                    fill_slots(other_slots, attempt_assignments, attempt_large_counts, attempt_workload)
                    
                    # 3. صقل الحل الناتج عبر التبديلات
                    polishing_swaps_count = int(settings.get('polishingSwaps', 15))
                    warmed_up_schedule, _, _, _ = run_post_processing_swaps(schedule_for_attempt, attempt_assignments, attempt_workload, attempt_large_counts, settings, all_professors, date_map, polishing_swaps_count)

                    # 4. التحقق من الصلاحية وتقييم التكلفة
                    if is_schedule_valid(warmed_up_schedule, settings, all_professors, duty_patterns, date_map):
                        current_cost = calculate_cost(warmed_up_schedule)
                        if current_cost < best_initial_cost:
                            best_initial_cost = current_cost
                            best_initial_solution = warmed_up_schedule
                            log_q.put(f"... [إحماء {attempt + 1}/{num_warmup_attempts}] تم العثور على حل أولي أفضل بتكلفة = {best_initial_cost:.2f}")

                # --- مرحلة التشغيل المتقدم ---
                if not best_initial_solution:
                    log_q.put(f"!!! فشل: لم يتم العثور على أي حل أولي صالح بعد {num_warmup_attempts} محاولة. قد تكون القيود متضاربة جداً.")
                    final_schedule_from_strategy = current_schedule # إرجاع أي حل لتجنب الخطأ
                else:
                    log_q.put(f"✓ انتهت مرحلة الإحماء. أفضل تكلفة أولية: {best_initial_cost:.2f}. بدء الخوارزمية المتقدمة...")
                    
                    # تحديد الحراس المقفلين (مثل أستاذ المادة)
                    locked_guards = set()
                    if assign_owner_as_guard:
                        prof_last_exam = {}
                        all_exams_flat = [exam for date_exams in best_initial_solution.values() for time_slots_in_day in date_exams.values() for exam in time_slots_in_day]
                        for exam in all_exams_flat:
                            owner = subject_owners.get((clean_string_for_matching(exam['subject']), clean_string_for_matching(exam['level'])))
                            if owner:
                                exam_date_time_str = f"{exam['date']} {exam['time'].split('-')[0]}"
                                if owner not in prof_last_exam or exam_date_time_str > prof_last_exam[owner]['datetime_str']:
                                    prof_last_exam[owner] = {'exam_uuid': exam['uuid'], 'exam_date': exam['date'], 'datetime_str': exam_date_time_str}
                        for owner, data in prof_last_exam.items():
                            if data['exam_date'] not in unavailable_days.get(owner, []):
                                locked_guards.add((data['exam_uuid'], owner))
                    
                    # تشغيل الخوارزمية المختارة على أفضل حل تم إيجاده
                    if balancing_strategy == 'lns':
                        final_schedule_from_strategy, _, _, _ = run_large_neighborhood_search(best_initial_solution, settings, all_professors, duty_patterns, date_map, log_q, locked_guards)
                    elif balancing_strategy == 'vns':
                        final_schedule_from_strategy, _, _, _ = run_variable_neighborhood_search(best_initial_solution, settings, all_professors, duty_patterns, date_map, log_q, locked_guards)

                strategy_success = True
            
            elif balancing_strategy == 'tabu_search':
                # --- بداية: منطق التوليد الأولي التكراري والمحسن ---
                
                # أ. إعدادات ومحاولة بناء جدول أولي صالح
                MAX_INITIAL_ATTEMPTS = 20
                filled_schedule = None
                is_initial_schedule_valid = False

                all_scheduled_exams_flat = [exam for date_exams in current_schedule.values() for time_slots_in_day in date_exams.values() for exam in time_slots_in_day]
                for exam in all_scheduled_exams_flat:
                    if 'uuid' not in exam:
                        exam['uuid'] = str(uuid.uuid4())
                
                # حلقة المحاولات المتعددة
                for attempt in range(MAX_INITIAL_ATTEMPTS):
                    log_q.put(f"... [محاولة {attempt + 1}/{MAX_INITIAL_ATTEMPTS}] لبناء جدول حراس أولي...")
                    
                    schedule_copy = copy.deepcopy(current_schedule)

                    locked_guards = set()
                    if assign_owner_as_guard:
                        prof_last_exam = {}
                        for exam in all_scheduled_exams_flat:
                            owner = subject_owners.get((clean_string_for_matching(exam['subject']), clean_string_for_matching(exam['level'])))
                            if owner:
                                exam_date_time_str = f"{exam['date']} {exam['time'].split('-')[0]}"
                                if owner not in prof_last_exam or exam_date_time_str > prof_last_exam[owner]['datetime_str']:
                                    # ===== السطر الذي تم تصحيحه =====
                                    prof_last_exam[owner] = {'exam_uuid': exam['uuid'], 'exam_date': exam['date'], 'datetime_str': exam_date_time_str}
                        
                        for owner, data in prof_last_exam.items():
                            if data['exam_date'] not in unavailable_days.get(owner, []):
                                locked_guards.add((data['exam_uuid'], owner))

                    # ملء بقية الحراس
                    temp_filled_schedule = complete_schedule_with_guards(schedule_copy, settings, all_professors, assignments, all_levels_list, duty_patterns, date_map, all_subjects, locked_guards)
                    
                    # التحقق من صلاحية الجدول الناتج
                    if is_schedule_valid(temp_filled_schedule, settings, all_professors, duty_patterns, date_map):
                        log_q.put(f"✓ نجاح! تم العثور على جدول أولي صالح في المحاولة رقم {attempt + 1}.")
                        filled_schedule = temp_filled_schedule
                        is_initial_schedule_valid = True
                        break
                
                # ب. استكمال العملية بناءً على نتيجة المحاولات
                if not is_initial_schedule_valid:
                    log_q.put(f"!!! فشل: لم يتم العثور على جدول أولي صالح بعد {MAX_INITIAL_ATTEMPTS} محاولة. قد تكون القيود متضاربة جداً.")
                    final_schedule_from_strategy = temp_filled_schedule
                else:
                    log_q.put("✓ بدء مرحلة الإحماء والبحث المحظور...")
                    warmed_up_schedule, _, _, _ = run_post_processing_swaps(
                        filled_schedule, defaultdict(list), defaultdict(float), defaultdict(int), 
                        settings, all_professors, date_map, 
                        swap_attempts=100,
                        locked_guards=locked_guards
                    )
                    log_q.put("✓ انتهت مرحلة الإحماء.")
                    
                    final_schedule_from_strategy, _, _, _ = \
                        run_tabu_search(warmed_up_schedule, settings, all_professors, duty_patterns, date_map, log_q, locked_guards)
                
                strategy_success = True
                # --- نهاية المنطق النهائي ---

            else: 
                log_q.put(f">>> تشغيل استراتيجية: {balancing_strategy}...")
                
                all_scheduled_exams_flat = [exam for date_exams in current_schedule.values() for time_slots_in_day in date_exams.values() for exam in time_slots_in_day]

                if assign_owner_as_guard:
                    prof_subjects_map = {prof: set(f"{clean_string_for_matching(parse_unique_id(uid, all_levels_list)[0])} ({clean_string_for_matching(parse_unique_id(uid, all_levels_list)[1])})" for uid in uids) for prof, uids in assignments.items()}
                    for prof in all_professors:
                        prof_owned_subjects_set = prof_subjects_map.get(prof, set())
                        owned_exams = [exam for exam in all_scheduled_exams_flat if f"{exam['subject']} ({clean_string_for_matching(exam['level'])})" in prof_owned_subjects_set]
                        if not owned_exams: continue
                        owned_exams.sort(key=lambda x: (x['date'], x['time']))
                        for exam_to_lock in reversed(owned_exams):
                            if exam_to_lock['date'] in unavailable_days.get(prof, []): continue
                            is_busy = any(prof in e['guards'] for e in current_schedule[exam_to_lock['date']][exam_to_lock['time']])
                            if not is_busy:
                                is_large_hall_exam = any(h.get('type') == 'كبيرة' for h in exam_to_lock['halls'])
                                if is_large_hall_exam and current_prof_large_counts[prof] >= max_large_hall_shifts: continue
                                exam_to_lock['guards'].append(prof)
                                current_prof_assignments[prof].append(exam_to_lock)
                                current_prof_workload_score[prof] += large_hall_weight if is_large_hall_exam else other_hall_weight
                                if is_large_hall_exam: current_prof_large_counts[prof] += 1
                                break
                
                all_required_slots = []
                for exam in all_scheduled_exams_flat:
                    num_guards_needed = sum(guards_large_hall_val for h in exam.get('halls', []) if h.get('type') == 'كبيرة') + sum(guards_medium_hall_val for h in exam.get('halls', []) if h.get('type') == 'متوسطة') + sum(guards_small_hall_val for h in exam.get('halls', []) if h.get('type') == 'صغيرة')
                    num_to_add = num_guards_needed - len(exam.get('guards', []))
                    for _ in range(num_to_add): all_required_slots.append(exam)
                
                is_phased_initial_fill = balancing_strategy in ['phased', 'phased_polished']
                if is_phased_initial_fill:
                    large_hall_slots = [exam for exam in all_required_slots if any(h['type'] == 'كبيرة' for h in exam['halls'])]
                    other_hall_slots = [exam for exam in all_required_slots if not any(h['type'] == 'كبيرة' for h in exam['halls'])]
                    random.shuffle(large_hall_slots)
                    random.shuffle(other_hall_slots)
                    fill_slots(large_hall_slots, current_prof_assignments, current_prof_large_counts, current_prof_workload_score)
                    fill_slots(other_hall_slots, current_prof_assignments, current_prof_large_counts, current_prof_workload_score)
                else: 
                    random.shuffle(all_required_slots)
                    fill_slots(all_required_slots, current_prof_assignments, current_prof_large_counts, current_prof_workload_score)
                
                if balancing_strategy == 'advanced':
                    final_schedule_from_strategy, current_prof_assignments, current_prof_workload_score, current_prof_large_counts = \
                        run_post_processing_swaps(current_schedule, current_prof_assignments, current_prof_workload_score, current_prof_large_counts, settings, all_professors, date_map, swap_attempts)
                elif balancing_strategy == 'phased_polished':
                    final_schedule_from_strategy, current_prof_assignments, current_prof_workload_score, current_prof_large_counts = \
                        run_post_processing_swaps(current_schedule, current_prof_assignments, current_prof_workload_score, current_prof_large_counts, settings, all_professors, date_map, polishing_swaps)
                elif balancing_strategy == 'annealing':
                    final_schedule_from_strategy, current_prof_assignments, current_prof_workload_score, current_prof_large_counts = \
                        run_simulated_annealing(
                            current_schedule, current_prof_assignments, current_prof_workload_score, current_prof_large_counts,
                            settings, all_professors, date_map, duty_patterns, annealing_iterations, annealing_temp, annealing_cooling
                        )
                else: 
                    final_schedule_from_strategy = current_schedule
                
                strategy_success = True

            if not final_schedule_from_strategy:
                log_q.put(f"فشلت الاستراتيجية '{balancing_strategy}' في إيجاد حل في هذه المحاولة. الانتقال للمحاولة التالية...")
                continue
            
            all_exams_in_final_schedule_flat = [exam for date_exams in final_schedule_from_strategy.values() for time_slots_in_day in date_exams.values() for exam in time_slots_in_day]
            
            for exam in all_exams_in_final_schedule_flat:
                num_guards_needed_for_logic = (sum(guards_large_hall_val for h in exam.get('halls', []) if h.get('type') == 'كبيرة') + 
                                               sum(guards_medium_hall_val for h in exam.get('halls', []) if h.get('type') == 'متوسطة') + 
                                               sum(guards_small_hall_val for h in exam.get('halls', []) if h.get('type') == 'صغيرة'))
                
                current_guards = exam.get('guards', [])
                exam['guards'] = [g for g in current_guards if g != "**نقص**"]

                while len(exam['guards']) < num_guards_needed_for_logic:
                    exam['guards'].append("**نقص**")
                
                exam['guards'] = exam['guards'][:num_guards_needed_for_logic]
            
            temp_unfilled_slots_count = 0
            temp_current_failures = []
            temp_current_scheduling_report = []

            for exam in all_exams_in_final_schedule_flat:
                if "**نقص**" in exam.get('guards', []):
                    num_guards_needed_for_logic = (sum(guards_large_hall_val for h in exam.get('halls', []) if h.get('type') == 'كبيرة') + sum(guards_medium_hall_val for h in exam.get('halls', []) if h.get('type') == 'متوسطة') + sum(guards_small_hall_val for h in exam.get('halls', []) if h.get('type') == 'صغيرة'))
                    exam['guards_incomplete'] = True
                    num_actual_shortage = exam['guards'].count("**نقص**") 
                    temp_unfilled_slots_count += num_actual_shortage
                    num_filled = num_guards_needed_for_logic - num_actual_shortage
                    temp_current_scheduling_report.append({"subject": exam['subject'], "level": exam['level'], "reason": f"نقص حراس: {num_filled}/{num_guards_needed_for_logic} (مع {num_actual_shortage} خانة نقص)"})

            current_prof_assignments_for_report = defaultdict(list)
            for exam_report in all_exams_in_final_schedule_flat:
                for guard in exam_report.get('guards', []):
                    if guard != "**نقص**":
                        current_prof_assignments_for_report[guard].append({'date': exam_report['date'], 'time': exam_report['time']})
            
            for prof_name_check, pattern in duty_patterns.items():
                if prof_name_check not in all_professors: continue
                duties_dates_indices = sorted(list({date_map.get(d_item['date']) for d_item in current_prof_assignments_for_report.get(prof_name_check, []) if date_map.get(d_item['date']) is not None}))
                num_unique_duty_days = len(duties_dates_indices)
                if num_unique_duty_days == 0:
                    continue

                if pattern == 'consecutive_strict':
                    if num_unique_duty_days != 2 or (len(duties_dates_indices) > 1 and duties_dates_indices[1] - duties_dates_indices[0] != 1):
                        temp_current_failures.append({"name": prof_name_check, "reason": "فشل في تحقيق شرط اليومين المتتاليين الإلزامي."})
                elif pattern == 'flexible_2_days':
                    if num_unique_duty_days != 2:
                        temp_current_failures.append({"name": prof_name_check, "reason": "فشل في تحقيق شرط اليومين (مرن)."})
                elif pattern == 'flexible_3_days':
                    if num_unique_duty_days not in [2, 3]:
                        temp_current_failures.append({"name": prof_name_check, "reason": "فشل في تحقيق شرط يومين أو 3 أيام (مرن)."})

            # --- هنا يتم حفظ أفضل نتيجة ---
            if (best_result['schedule'] is None or 
               temp_unfilled_slots_count < best_result['unfilled_slots'] or 
               (temp_unfilled_slots_count == best_result['unfilled_slots'] and len(best_result['failures']) > len(temp_current_failures)) or
               (temp_unfilled_slots_count == best_result['unfilled_slots'] and len(best_result['failures']) == len(temp_current_failures) and len(current_all_subjects_to_schedule) < len(best_result['unscheduled_subjects']))):
                
                # ... (بقية الكود تبقى كما هي)
                
                prof_stats = {prof_name: {'large': 0, 'other': 0} for prof_name in all_professors}
                for exam_stat in all_exams_in_final_schedule_flat:
                    guards_for_this_exam = exam_stat.get('guards', [])
                    large_guards_needed_for_exam = sum(guards_large_hall_val for h in exam_stat.get('halls', []) if h.get('type') == 'كبيرة')
                    large_hall_guards = guards_for_this_exam[:large_guards_needed_for_exam]
                    other_hall_guards = guards_for_this_exam[large_guards_needed_for_exam:]
                    for guard in large_hall_guards:
                        if guard != "**نقص**" and guard in prof_stats: prof_stats[guard]['large'] += 1
                    for guard in other_hall_guards:
                        if guard != "**نقص**" and guard in prof_stats: prof_stats[guard]['other'] += 1
                
                prof_targets_map = {}
                if num_professors > 0:
                    if enable_custom_targets and custom_target_patterns:
                        prof_targets_list = []
                        for pattern in custom_target_patterns:
                            for _ in range(pattern.get('count', 0)): prof_targets_list.append({'large': pattern.get('large', 0), 'other': pattern.get('other', 0)})
                        num_to_fill = num_professors - len(prof_targets_list)
                        if num_to_fill > 0:
                            rem_large = sum(s['large'] for s in prof_stats.values()) - sum(p['large'] for p in prof_targets_list)
                            rem_other = sum(s['other'] for s in prof_stats.values()) - sum(p['other'] for p in prof_targets_list)
                            if rem_large >= 0 and rem_other >= 0:
                                prof_targets_list.extend(calculate_balanced_distribution(rem_large, rem_other, num_to_fill, large_hall_weight, other_hall_weight))
                        shuffled_profs = list(prof_stats.keys()); random.shuffle(shuffled_profs)
                        prof_targets_map = {prof: prof_targets_list[i] for i, prof in enumerate(shuffled_profs) if i < len(prof_targets_list)}
                    else:
                        total_large_slots_final = sum(s['large'] for s in prof_stats.values())
                        total_other_slots_final = sum(s['other'] for s in prof_stats.values())
                        prof_targets_list = calculate_balanced_distribution(total_large_slots_final, total_other_slots_final, num_professors, large_hall_weight, other_hall_weight)
                        if prof_targets_list: prof_targets_map = {prof: prof_targets_list[i % len(prof_targets_list)] for i, prof in enumerate(sorted(prof_stats.keys()))}
                
                balance_report = generate_balance_report(prof_stats, prof_targets_map) if prof_targets_map else {'balance_score': 0}
                current_balance_score = balance_report.get('balance_score', 0)
                
                log_q.put(f"✓ [Iteration {i+1}] Found a better solution! (Score: {current_balance_score}%, Nقص: {temp_unfilled_slots_count}, فشل القيود: {len(temp_current_failures)})")
                
                best_result['schedule'] = copy.deepcopy(final_schedule_from_strategy)
                best_result['failures'] = list(temp_current_failures)
                best_result['scheduling_report'] = list(temp_current_scheduling_report)
                best_result['unfilled_slots'] = temp_unfilled_slots_count
                
                # --- إضافة جديدة: تسجيل المواد التي لم تجدول في أفضل حل ---
                best_result['unscheduled_subjects'] = sorted([f"{name} ({level})" for name, level in current_all_subjects_to_schedule])
                
                prof_stats_report = []
                chart_data = { 'labels': [], 'datasets': [ {'label': 'حصص القاعات الأخرى', 'data': [], 'backgroundColor': 'rgba(54, 162, 235, 0.7)'}, {'label': 'حصص القاعة الكبيرة', 'data': [], 'backgroundColor': 'rgba(255, 99, 132, 0.7)'}] }
                
                sorted_prof_names_for_stats = sorted(prof_stats.keys())
                for prof_name in sorted_prof_names_for_stats:
                    stats = prof_stats[prof_name]
                    total = stats['large'] + stats['other']
                    prof_stats_report.append(f"{prof_name}: {total} حصص (قاعة كبيرة: {stats['large']}، قاعات أخرى: {stats['other']})")
                    chart_data['labels'].append(prof_name)
                    chart_data['datasets'][0]['data'].append(stats['other'])
                    chart_data['datasets'][1]['data'].append(stats['large'])
                
                stats_dashboard = {}
                dashboard_total_large = sum(stats['large'] for stats in prof_stats.values())
                dashboard_total_other = sum(stats['other'] for stats in prof_stats.values())
                
                stats_dashboard['total_large_duties'] = dashboard_total_large
                stats_dashboard['total_other_duties'] = dashboard_total_other
                stats_dashboard['total_duties'] = dashboard_total_large + dashboard_total_other
                stats_dashboard['avg_duties_per_prof'] = stats_dashboard['total_duties'] / num_professors if num_professors > 0 else 0
                
                duties_per_day = defaultdict(int)
                for exam_day_count in all_exams_in_final_schedule_flat:
                    for guard in exam_day_count.get('guards', []): 
                        if guard != "**نقص**": duties_per_day[exam_day_count['date']] += 1
                
                if duties_per_day:
                    busiest_day_date = max(duties_per_day, key=duties_per_day.get)
                    stats_dashboard['busiest_day'] = {'date': busiest_day_date, 'duties': duties_per_day[busiest_day_date]}
                else:
                    stats_dashboard['busiest_day'] = {'date': 'N/A', 'duties': 0}

                prof_workload_for_dashboard = {p: (s['large'] * large_hall_weight) + (s['other'] * other_hall_weight) for p, s in prof_stats.items()}
                sorted_profs_by_workload = sorted(prof_workload_for_dashboard.items(), key=lambda item: item[1])
                stats_dashboard['least_burdened_profs'] = [{'name': p[0], 'workload': round(p[1], 2)} for p in sorted_profs_by_workload[:3]]
                stats_dashboard['most_burdened_profs'] = [{'name': p[0], 'workload': round(p[1], 2)} for p in sorted_profs_by_workload[-3:]][::-1]

                best_result['prof_report'] = prof_stats_report
                best_result['chart_data'] = chart_data
                best_result['balance_report'] = balance_report
                
                shortage_report_for_dashboard = []
                for report_item in temp_current_scheduling_report:
                    if "نقص" in report_item.get("reason", ""):
                        shortage_report_for_dashboard.append(
                            f"{report_item['subject']} ({report_item['level']})"
                        )
                stats_dashboard['shortage_reports'] = shortage_report_for_dashboard
                
                # --- إضافة جديدة: إضافة تقرير المواد غير المجدولة إلى لوحة المعلومات ---
                stats_dashboard['unscheduled_subjects_report'] = best_result['unscheduled_subjects']
                best_result['stats_dashboard'] = stats_dashboard

            # if best_result.get('unfilled_slots', float('inf')) == 0 and not best_result.get('failures') and not best_result.get('unscheduled_subjects'):
            #     log_q.put(f">>> [Iteration {i+1}] تم العثور على حل مثالي. إنهاء البحث المكثف.")
            #     break 

        if not best_result.get('schedule'):
            log_q.put("فشل البرنامج في إيجاد أي حل ضمن المحاولات المحددة. الرجاء مراجعة القيود.")
            log_q.put("DONE" + json.dumps({"success": False, "message": "فشل إنشاء أي حل. حاول تخفيف القيود أو زيادة عدد محاولات البحث."}))
            return

        log_q.put("DONE" + json.dumps({ 
            "success": True, 
            "schedule": best_result['schedule'], 
            "failures": best_result.get('failures', []), 
            "scheduling_report": best_result.get('scheduling_report', []),
            "prof_report": best_result.get('prof_report', []),
            "chart_data": best_result.get('chart_data', {}),
            "balance_report": best_result.get('balance_report', {}),
            "stats_dashboard": best_result.get('stats_dashboard', {})
        }, ensure_ascii=False))

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        log_q.put(f"خطأ فادح في معالجة الخلفية: {str(e)}")
        log_q.put(error_details)
        log_q.put("DONE" + json.dumps({"success": False, "message": f"خطأ فادح: {e}"}))


# يمكن إضافتها بعد دوال الخوارزمية وقبل مسارات API

# استبدل هذه الدالة بالكامل في ملف app.py
def create_word_document_with_table(doc, title, headers, data_grid):
    """
    يضيف عنواناً وجدولاً إلى مستند وورد موجود مع دعم كامل للغة العربية (RTL).
    """
    # --- ✅ بداية: الكود الجديد لتصحيح اتجاه العنوان ---
    heading = doc.add_heading(level=2)
    heading.clear() # مسح أي محتوى افتراضي
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # إضافة خاصية اتجاه الفقرة من اليمين لليسار
    pPr = heading._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    
    # إضافة النص كـ "run" مع تفعيل خاصية RTL للخط
    run = heading.add_run(title)
    font = run.font
    font.rtl = True
    font.name = 'Arial'
    # --- ✅ نهاية: الكود الجديد لتصحيح اتجاه العنوان ---

    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.autofit = False

    tbl_pr = table._element.xpath('w:tblPr')[0]
    bidi_visual_element = OxmlElement('w:bidiVisual')
    tbl_pr.append(bidi_visual_element)

    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell_paragraph = hdr_cells[i].paragraphs[0]
        cell_paragraph.text = ""
        run = cell_paragraph.add_run(header)
        font = run.font
        font.rtl = True
        font.name = 'Arial'
        cell_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_paragraph.paragraph_format.rtl = True

    for row_data in data_grid:
        row_cells = table.add_row().cells
        for i, cell_data in enumerate(row_data):
            cell_paragraph = row_cells[i].paragraphs[0]
            cell_paragraph.text = ""
            lines = str(cell_data).split('\n')
            for idx, line in enumerate(lines):
                if idx > 0:
                    cell_paragraph.add_run().add_break()
                run = cell_paragraph.add_run(line)
                font = run.font
                font.rtl = True
                font.name = 'Arial'
            cell_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            cell_paragraph.paragraph_format.rtl = True
            
    doc.add_page_break()

# ================== الجزء الثالث: إعداد تطبيق فلاسك والمسارات (Routes) ==================
app = Flask( __name__, template_folder=get_correct_path('templates'), static_folder=get_correct_path('static') )
app.config['JSON_AS_ASCII'] = False
app.config['JSON_SORT_KEYS'] = False

@app.route('/')
def index(): return render_template('index.html')

@app.route('/api/professors', methods=['GET'])
def get_professors():
    conn = get_db_connection()
    professors = conn.execute('SELECT name FROM professors ORDER BY name').fetchall()
    conn.close()
    return jsonify([dict(row) for row in professors])

@app.route('/api/subjects', methods=['GET'])
def get_subjects():
    conn = get_db_connection()
    subjects = conn.execute('SELECT s.name, l.name as level FROM subjects s JOIN levels l ON s.level_id = l.id ORDER BY l.name, s.name').fetchall()
    conn.close()
    return jsonify([dict(row) for row in subjects])

@app.route('/api/levels', methods=['GET'])
def get_levels():
    conn = get_db_connection()
    levels = conn.execute('SELECT name FROM levels ORDER BY name').fetchall()
    conn.close()
    return jsonify([row['name'] for row in levels])

@app.route('/api/halls', methods=['GET'])
def get_halls():
    conn = get_db_connection()
    halls = conn.execute('SELECT name, type FROM halls ORDER BY name').fetchall()
    conn.close()
    return jsonify([dict(row) for row in halls])

@app.route('/api/assignments', methods=['GET'])
def get_assignments():
    conn = get_db_connection()
    rows = conn.execute('SELECT p.name as professor_name, s.name as subject_name, l.name as level_name FROM assignments a JOIN professors p ON a.professor_id = p.id JOIN subjects s ON a.subject_id = s.id JOIN levels l ON s.level_id = l.id').fetchall()
    conn.close()
    assignments_dict = defaultdict(list)
    for row in rows:
        unique_id = f"{row['subject_name']} ({row['level_name']})"
        assignments_dict[row['professor_name']].append(unique_id)
    return jsonify(assignments_dict)

@app.route('/api/settings', methods=['GET'])
def get_settings():
    conn = get_db_connection()
    setting_row = conn.execute("SELECT value FROM settings WHERE key = 'main_settings'").fetchone()
    conn.close()
    if setting_row: return jsonify(json.loads(setting_row['value']))
    return jsonify({})

@app.route('/api/settings', methods=['POST'])
def save_settings():
    settings_data = request.get_json()
    conn = get_db_connection()
    conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('main_settings', json.dumps(settings_data, ensure_ascii=False)))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "تم حفظ الإعدادات بنجاح."})

@app.route('/api/professors/bulk', methods=['POST'])
def add_professors_bulk():
    data = request.get_json()
    new_names = data.get('names', [])
    if not new_names: return jsonify({"error": "البيانات غير صالحة"}), 400
    conn = get_db_connection()
    cursor = conn.cursor()
    added_count = 0
    for name in new_names:
        try:
            cursor.execute("INSERT INTO professors (name) VALUES (?)", (name,))
            added_count += 1
        except sqlite3.IntegrityError: pass
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"تمت إضافة {added_count} أساتذة بنجاح."}), 201

@app.route('/api/halls/bulk', methods=['POST'])
def add_halls_bulk():
    data = request.get_json()
    new_halls = data.get('halls', [])
    if not new_halls: return jsonify({"error": "البيانات غير صالحة"}), 400
    conn = get_db_connection()
    cursor = conn.cursor()
    added_count = 0
    for hall in new_halls:
        try:
            cursor.execute("INSERT INTO halls (name, type) VALUES (?, ?)", (hall['name'], hall['type']))
            added_count += 1
        except sqlite3.IntegrityError: pass
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"تمت إضافة {added_count} قاعات بنجاح."}), 201

@app.route('/api/levels/bulk', methods=['POST'])
def add_levels_bulk():
    data = request.get_json()
    new_names = data.get('names', [])
    if not new_names: return jsonify({"error": "بيانات غير صالحة"}), 400
    conn = get_db_connection()
    cursor = conn.cursor()
    added_count = 0
    for name in new_names:
        try:
            cursor.execute("INSERT INTO levels (name) VALUES (?)", (name,))
            added_count += 1
        except sqlite3.IntegrityError: pass
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"تمت إضافة {added_count} مستويات بنجاح."}), 201

@app.route('/api/subjects/bulk', methods=['POST'])
def add_subjects_bulk():
    data = request.get_json()
    new_subject_names = data.get('subjects', [])
    level_name = data.get('level')
    if not new_subject_names or not level_name: return jsonify({"error": "البيانات غير مكتملة"}), 400
    conn = get_db_connection()
    cursor = conn.cursor()
    level_row = cursor.execute("SELECT id FROM levels WHERE name = ?", (level_name,)).fetchone()
    if not level_row:
        conn.close()
        return jsonify({"error": f"المستوى '{level_name}' غير موجود."}), 404
    level_id = level_row['id']
    added_count = 0
    for subj_name in new_subject_names:
        exists = cursor.execute("SELECT 1 FROM subjects WHERE name = ? AND level_id = ?", (subj_name, level_id)).fetchone()
        if not exists:
            cursor.execute("INSERT INTO subjects (name, level_id) VALUES (?, ?)", (subj_name, level_id))
            added_count += 1
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"تمت إضافة {added_count} مواد بنجاح."}), 201

@app.route('/api/assign-subjects/bulk', methods=['POST'])
def assign_subjects_bulk():
    data = request.get_json()
    prof_name = data.get('professor')
    subj_unique_ids = data.get('subjects', [])
    if not prof_name or not subj_unique_ids: return jsonify({"error": "بيانات غير كافية"}), 400
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    prof_row = cursor.execute("SELECT id FROM professors WHERE name = ?", (prof_name,)).fetchone()
    if not prof_row:
        conn.close()
        return jsonify({"error": f"خطأ: الأستاذ '{prof_name}' غير موجود في قاعدة البيانات."}), 404
    prof_id = prof_row['id']

    all_levels = [row['name'] for row in cursor.execute("SELECT name FROM levels").fetchall()]
    added_count = 0
    
    for unique_id in subj_unique_ids:
        subj_name, level_name = parse_unique_id(unique_id, all_levels)
        if not subj_name or not level_name:
            conn.close()
            return jsonify({"error": f"خطأ في تحليل اسم المادة: '{unique_id}'"}), 400
            
        subj_row = cursor.execute('SELECT s.id FROM subjects s JOIN levels l ON s.level_id = l.id WHERE s.name = ? AND l.name = ?', (subj_name, level_name)).fetchone()
        
        if subj_row:
            subj_id = subj_row['id']
            try:
                cursor.execute("INSERT INTO assignments (professor_id, subject_id) VALUES (?, ?)", (prof_id, subj_id))
                added_count += 1
            except sqlite3.IntegrityError: pass
        else:
            # ==> هذا هو التعديل المهم: إرجاع رسالة خطأ بدلاً من التجاهل <==
            conn.close()
            return jsonify({"error": f"خطأ: المادة '{subj_name}' في المستوى '{level_name}' غير موجودة. تأكد من تطابق الأسماء."}), 404

    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"تم تخصيص {added_count} مواد للأستاذ '{prof_name}'"})

@app.route('/api/professors', methods=['DELETE'])
def delete_professor():
    data = request.get_json()
    name_to_delete = data.get('name')
    if not name_to_delete: return jsonify({"error": "اسم الأستاذ مفقود"}), 400
    conn = get_db_connection()
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("DELETE FROM professors WHERE name = ?", (name_to_delete,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "تم حذف الأستاذ."})

@app.route('/api/halls', methods=['DELETE'])
def delete_hall():
    data = request.get_json()
    name_to_delete = data.get('name')
    if not name_to_delete: return jsonify({"error": "اسم القاعة مفقود"}), 400
    conn = get_db_connection()
    conn.execute("DELETE FROM halls WHERE name = ?", (name_to_delete,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "تم حذف القاعة."})

@app.route('/api/levels', methods=['DELETE'])
def delete_level():
    data = request.get_json()
    name_to_delete = data.get('name')
    if not name_to_delete: return jsonify({"error": "اسم المستوى مفقود"}), 400
    conn = get_db_connection()
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("DELETE FROM levels WHERE name = ?", (name_to_delete,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "تم حذف المستوى."})

@app.route('/api/subjects', methods=['DELETE'])
def delete_subject():
    data = request.get_json()
    name_to_delete = data.get('name')
    level_to_delete = data.get('level')
    if not name_to_delete or not level_to_delete: return jsonify({"error": "بيانات المادة غير مكتملة"}), 400
    conn = get_db_connection()
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute('DELETE FROM subjects WHERE name = ? AND level_id = (SELECT id FROM levels WHERE name = ?)', (name_to_delete, level_to_delete))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "تم حذف المادة."})

@app.route('/api/unassign-subject', methods=['POST'])
def unassign_subject():
    data = request.get_json()
    unique_id_to_unassign = data.get('subject')
    if not unique_id_to_unassign: return jsonify({"error": "لم يتم تحديد المادة"}), 400
    conn = get_db_connection()
    cursor = conn.cursor()
    all_levels = [row['name'] for row in cursor.execute("SELECT name FROM levels").fetchall()]
    subj_name, level_name = parse_unique_id(unique_id_to_unassign, all_levels)
    if subj_name and level_name:
        cursor.execute('DELETE FROM assignments WHERE subject_id = (SELECT s.id FROM subjects s JOIN levels l ON s.level_id = l.id WHERE s.name = ? AND l.name = ?)', (subj_name, level_name))
        conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"تم إلغاء إسناد المادة '{unique_id_to_unassign}' بنجاح."})

@app.route('/api/professors/edit', methods=['POST'])
def edit_professor():
    data = request.get_json()
    old_name, new_name = data.get('old_name'), data.get('new_name')
    if not old_name or not new_name: return jsonify({"error": "البيانات غير كافية"}), 400
    conn = get_db_connection()
    conn.execute("UPDATE professors SET name = ? WHERE name = ?", (new_name, old_name))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "تم تعديل اسم الأستاذ بنجاح."})

@app.route('/api/halls/edit', methods=['POST'])
def edit_hall():
    data = request.get_json()
    old_name, new_name, new_type = data.get('old_name'), data.get('new_name'), data.get('new_type')
    if not old_name or not new_name or not new_type: return jsonify({"error": "البيانات غير كافية"}), 400
    conn = get_db_connection()
    conn.execute("UPDATE halls SET name = ?, type = ? WHERE name = ?", (new_name, new_type, old_name))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "تم تعديل القاعة بنجاح."})

@app.route('/api/subjects/edit', methods=['POST'])
def edit_subject():
    data = request.get_json()
    old_name, new_name = data.get('old_name'), data.get('new_name')
    old_level, new_level = data.get('old_level'), data.get('new_level')
    if not all([old_name, new_name, old_level, new_level]): return jsonify({"error": "البيانات غير كافية"}), 400
    conn = get_db_connection()
    new_level_id_row = conn.execute("SELECT id FROM levels WHERE name = ?", (new_level,)).fetchone()
    if not new_level_id_row:
        conn.close()
        return jsonify({"error": "المستوى الجديد غير موجود"}), 404
    new_level_id = new_level_id_row['id']
    conn.execute("UPDATE subjects SET name = ?, level_id = ? WHERE name = ? AND level_id = (SELECT id FROM levels WHERE name = ?)", (new_name, new_level_id, old_name, old_level))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "تم تعديل المادة بنجاح."})

@app.route('/api/levels/edit', methods=['POST'])
def edit_level():
    data = request.get_json()
    old_name, new_name = data.get('old_name'), data.get('new_name')
    if not old_name or not new_name: return jsonify({"error": "البيانات غير كافية"}), 400
    conn = get_db_connection()
    conn.execute("UPDATE levels SET name = ? WHERE name = ?", (new_name, old_name))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "تم تعديل المستوى بنجاح."})

# --- Routes for Backup, Restore, Export ---
@app.route('/api/backup', methods=['GET'])
def backup_data():
    conn = get_db_connection()
    professors = [dict(row) for row in conn.execute("SELECT name FROM professors").fetchall()]
    halls = [dict(row) for row in conn.execute("SELECT name, type FROM halls").fetchall()]
    levels = [row['name'] for row in conn.execute("SELECT name FROM levels").fetchall()]
    subjects_rows = conn.execute("SELECT s.name, l.name as level FROM subjects s JOIN levels l ON s.level_id = l.id").fetchall()
    subjects = [dict(row) for row in subjects_rows]
    assignments_rows = conn.execute('SELECT p.name as prof_name, s.name as subj_name, l.name as level_name FROM assignments a JOIN professors p ON a.professor_id = p.id JOIN subjects s ON a.subject_id = s.id JOIN levels l ON s.level_id = l.id').fetchall()
    assignments = defaultdict(list)
    for row in assignments_rows:
        assignments[row['prof_name']].append(f"{row['subj_name']} ({row['level_name']})")
    settings_row = conn.execute("SELECT value FROM settings WHERE key = 'main_settings'").fetchone()
    settings = json.loads(settings_row['value']) if settings_row else {}
    conn.close()
    all_data = {'professors': professors, 'halls': halls, 'levels': levels, 'subjects': subjects, 'assignments': assignments, 'settings': settings}
    json_string = json.dumps(all_data, ensure_ascii=False, indent=4)
    buffer = io.BytesIO(json_string.encode('utf-8'))
    return send_file(buffer, as_attachment=True, download_name="backup_sqlite.json", mimetype="application/json")

@app.route('/api/restore', methods=['POST'])
def restore_data():
    backup_data = request.get_json()
    required_keys = ['professors', 'halls', 'levels', 'subjects', 'assignments', 'settings']
    if not all(key in backup_data for key in required_keys): return jsonify({"error": "ملف النسخة الاحتياطية غير صالح أو تالف."}), 400
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        for table in ['assignments', 'subjects', 'levels', 'halls', 'professors', 'settings']:
            cursor.execute(f"DELETE FROM {table}")
        cursor.executemany("INSERT INTO professors (name) VALUES (?)", [(p['name'],) for p in backup_data['professors']])
        cursor.executemany("INSERT INTO halls (name, type) VALUES (?, ?)", [(h['name'], h['type']) for h in backup_data['halls']])
        cursor.executemany("INSERT INTO levels (name) VALUES (?)", [(l,) for l in backup_data['levels']])
        level_map = {row['name']: row['id'] for row in cursor.execute("SELECT id, name FROM levels").fetchall()}
        subjects_to_insert = [(s['name'], level_map[s['level']]) for s in backup_data['subjects'] if s['level'] in level_map]
        cursor.executemany("INSERT INTO subjects (name, level_id) VALUES (?, ?)", subjects_to_insert)
        prof_map = {row['name']: row['id'] for row in cursor.execute("SELECT id, name FROM professors").fetchall()}
        subj_map = {(row['name'], row['level']): row['id'] for row in cursor.execute("SELECT s.id, s.name, l.name as level FROM subjects s JOIN levels l ON s.level_id = l.id").fetchall()}
        assignments_to_insert = []
        all_levels_list = backup_data['levels']
        for prof_name, unique_ids in backup_data['assignments'].items():
            if prof_name in prof_map:
                prof_id = prof_map[prof_name]
                for uid in unique_ids:
                    subj_name, level_name = parse_unique_id(uid, all_levels_list)
                    if (subj_name, level_name) in subj_map:
                        subj_id = subj_map[(subj_name, level_name)]
                        assignments_to_insert.append((prof_id, subj_id))
        cursor.executemany("INSERT OR IGNORE INTO assignments (professor_id, subject_id) VALUES (?, ?)", assignments_to_insert)
        cursor.execute("INSERT INTO settings (key, value) VALUES (?, ?)", ('main_settings', json.dumps(backup_data['settings'])))
        conn.commit()
        conn.close()
    except Exception as e: return jsonify({"error": f"حدث خطأ أثناء استعادة البيانات: {e}"}), 500
    return jsonify({"success": True, "message": "تم استعادة البيانات بنجاح. سيتم إعادة تحميل الصفحة."})

@app.route('/api/reset-all', methods=['POST'])
def reset_all_data():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        for table in ['assignments', 'subjects', 'levels', 'halls', 'professors', 'settings']:
            cursor.execute(f"DELETE FROM {table}")
        conn.commit()
        conn.close()
        return jsonify({"success": True, "message": "تم مسح جميع البيانات بنجاح. سيتم إعادة تحميل الصفحة."})
    except Exception as e: return jsonify({"error": f"حدث خطأ أثناء مسح البيانات: {e}"}), 500

def build_schedule_table_for_level(ws, start_row, level, all_dates, all_times, day_names, schedule_data, styles, settings):
    header_font, default_font, level_font = styles['header_font'], styles['default_font'], styles['level_font']
    border, header_fill = styles['border'], styles['header_fill']
    guards_large, guards_medium, guards_small = settings['guards_large'], settings['guards_medium'], settings['guards_small']
    current_row = start_row
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(all_dates) + 1)
    title_cell = ws.cell(row=current_row, column=1, value=f"جدول امتحانات: {level}")
    apply_styles_to_cell(title_cell, level_font, Alignment(horizontal='center', vertical='center'), border)
    current_row += 1
    header_row_index = current_row
    headers = ["الفترة"] + [f"{day_names[datetime.strptime(d, '%Y-%m-%d').isoweekday() % 7]}\n{d}" for d in all_dates]
    for col_num, header_value in enumerate(headers, 1):
        cell = ws.cell(row=header_row_index, column=col_num, value=header_value)
        apply_styles_to_cell(cell, header_font, Alignment(horizontal='center', vertical='center', wrap_text=True), border, header_fill)
    ws.row_dimensions[header_row_index].height = 40
    current_row += 1
    for time in all_times:
        data_row_index = current_row
        max_lines_in_row = 1
        time_cell = ws.cell(row=data_row_index, column=1, value=time)
        apply_styles_to_cell(time_cell, header_font, Alignment(horizontal='center', vertical='top', wrap_text=True), border)
        for col_idx, date in enumerate(all_dates, 2):
            exam = next((e for e in schedule_data.get(date, {}).get(time, []) if e['level'] == level), None)
            content = ""
            if exam:
                content = f"{exam['subject']}\nأستاذ المادة: {exam.get('professor', 'N/A')}\n\nالحراسة:"
                halls_by_type = defaultdict(list)
                for h in exam.get('halls', []): halls_by_type[h['type']].append(h['name'])
                guards = exam.get('guards', [])[:]
                for t, n, p in [('القاعة الكبيرة', 'كبيرة', guards_large), ('القاعات المتوسطة', 'متوسطة', guards_medium), ('القاعات الصغيرة', 'صغيرة', guards_small)]:
                    if halls_by_type.get(n):
                        g_list = guards[:len(halls_by_type[n]) * p]
                        guards = guards[len(halls_by_type[n]) * p:]
                        hall_names = ", ".join(halls_by_type[n])
                        guard_text = '\n'.join(g_list) if g_list else '(لا يوجد حراس)'
                        content += f"\n{t}: {hall_names}\n{guard_text}"
            num_lines = content.count('\n') + 1
            max_lines_in_row = max(max_lines_in_row, num_lines)
            data_cell = ws.cell(row=data_row_index, column=col_idx, value=content)
            apply_styles_to_cell(data_cell, default_font, Alignment(horizontal='right', vertical='top', wrap_text=True), border)
        ws.row_dimensions[data_row_index].height = max(40, max_lines_in_row * 15)
        current_row += 1
    ws.column_dimensions['A'].width = 15
    for i in range(2, len(all_dates) + 2):
        ws.column_dimensions[get_column_letter(i)].width = 35
    return current_row

@app.route('/api/export-schedule', methods=['POST'])
def export_schedule():
    schedule_data = request.get_json()
    if not schedule_data: return jsonify({"error": "No schedule data provided"}), 400
    conn = get_db_connection()
    settings_row = conn.execute("SELECT value FROM settings WHERE key = 'main_settings'").fetchone()
    conn.close()
    settings_data = json.loads(settings_row['value']) if settings_row else {}
    settings_for_export = {
        'guards_large': int(settings_data.get('guardsLargeHall', 4)),
        'guards_medium': int(settings_data.get('guardsMediumHall', 2)),
        'guards_small': int(settings_data.get('guardsSmallHall', 1))
    }
    wb = Workbook()
    if 'Sheet' in wb.sheetnames: wb.remove(wb['Sheet'])
    styles = { 'header_font': Font(bold=True, size=12, name='Calibri'), 'default_font': Font(size=11, name='Calibri'), 'level_font': Font(bold=True, size=16, name='Calibri'), 'border': Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin')), 'header_fill': PatternFill(fill_type="solid", start_color="D9D9D9") }
    all_dates = sorted(schedule_data.keys())
    all_times = sorted({time for date_slots in schedule_data.values() for time in date_slots})
    all_levels = sorted({exam['level'] for date_slots in schedule_data.values() for time_slots in date_slots.values() for exam in time_slots})
    day_names = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
    summary_ws = wb.create_sheet(title="الجداول المجمعة", index=0)
    summary_ws.sheet_view.rightToLeft = True
    summary_current_row = 1
    for level in all_levels:
        summary_current_row = build_schedule_table_for_level(summary_ws, summary_current_row, level, all_dates, all_times, day_names, schedule_data, styles, settings_for_export)
        summary_current_row += 2 
        individual_ws = wb.create_sheet(title=level[:31])
        individual_ws.sheet_view.rightToLeft = True
        build_schedule_table_for_level(individual_ws, 1, level, all_dates, all_times, day_names, schedule_data, styles, settings_for_export)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="الجداول_المجمعة.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route('/api/export-prof-schedules', methods=['POST'])
def export_prof_schedules():
    schedule_data = request.get_json()
    if not schedule_data: return jsonify({"error": "No schedule data provided"}), 400
    conn = get_db_connection()
    all_professors = sorted([p['name'] for p in conn.execute("SELECT name FROM professors").fetchall()])
    assignments_rows = conn.execute('SELECT p.name as prof_name, s.name as subj_name, l.name as level_name FROM assignments a JOIN professors p ON a.professor_id = p.id JOIN subjects s ON a.subject_id = s.id JOIN levels l ON s.level_id = l.id').fetchall()
    all_levels_list = [row['name'] for row in conn.execute("SELECT name FROM levels").fetchall()]
    conn.close()
    cleaned_assignments = defaultdict(set)
    for row in assignments_rows:
        cleaned_assignments[row['prof_name']].add((clean_string_for_matching(row['subj_name']), clean_string_for_matching(row['level_name'])))
    wb = Workbook()
    if 'Sheet' in wb.sheetnames: wb.remove(wb['Sheet'])
    header_font = Font(bold=True, name='Calibri', size=11)
    default_font = Font(name='Calibri', size=10)
    bold_font = Font(bold=True, name='Calibri', size=10)
    title_font = Font(bold=True, name='Calibri', size=14)
    header_fill = PatternFill(start_color="D3D3D3", fill_type="solid")
    assigned_subject_fill = PatternFill(start_color="FFEB9C", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    all_dates = sorted(list(schedule_data.keys()))
    all_times = sorted(list(set(time for date_slots in schedule_data.values() for time in date_slots.keys())))
    day_names = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
    exam_map = {}
    for date, time_slots in schedule_data.items():
        for time, exams in time_slots.items(): exam_map[(date, time)] = exams
    for prof_name in all_professors:
        ws = wb.create_sheet(title=prof_name[:31])
        ws.sheet_view.rightToLeft = True
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(all_times) + 1)
        title_cell = ws.cell(row=1, column=1, value=f"جدول الحراسة الخاص بالأستاذ: {prof_name}")
        apply_styles_to_cell(title_cell, title_font, Alignment(horizontal='center', vertical='center'), thin_border)
        ws.row_dimensions[1].height = 30
        headers = ["اليوم/التاريخ"] + all_times
        ws.append(headers)
        header_row_num = ws.max_row
        for cell in ws[header_row_num]: apply_styles_to_cell(cell, header_font, Alignment(horizontal='center', vertical='center', wrap_text=True), thin_border, header_fill)
        for date in all_dates:
            day_name = day_names[datetime.strptime(date, '%Y-%m-%d').isoweekday() % 7]
            row_num = ws.max_row + 1
            date_cell = ws.cell(row=row_num, column=1, value=f"{day_name}\n{date}")
            apply_styles_to_cell(date_cell, header_font, Alignment(horizontal='center', vertical='center', wrap_text=True), thin_border)
            for col_idx, time in enumerate(all_times, 2):
                cell = ws.cell(row=row_num, column=col_idx)
                cell_parts, is_teaching_in_slot, processed_subjects = [], False, set()
                exams_in_slot = exam_map.get((date, time), [])
                prof_assigned_subjects_set = cleaned_assignments.get(prof_name, set())
                for exam in exams_in_slot:
                    exam_tuple = (clean_string_for_matching(exam.get('subject', '')), clean_string_for_matching(exam.get('level', '')))
                    if exam_tuple in prof_assigned_subjects_set:
                        is_teaching_in_slot = True
                        is_guarding = prof_name in exam.get('guards', [])
                        hall_names = ", ".join(sorted([h.get('name', '') for h in exam.get('halls', [])]))
                        if is_guarding: cell_parts.append(f"حراسة: {exam.get('subject')} ({exam.get('level')})\nقاعة: {hall_names}")
                        else: cell_parts.append(f"{exam.get('subject')} ({exam.get('level')})\n(دون حراسة)")
                        processed_subjects.add(exam.get('subject'))
                for exam in exams_in_slot:
                    is_guarding = prof_name in exam.get('guards', [])
                    if is_guarding and exam.get('subject') not in processed_subjects:
                        hall_names = ", ".join(sorted([h.get('name', '') for h in exam.get('halls', [])]))
                        cell_parts.append(f"حراسة: {exam.get('subject')} ({exam.get('level')})\nقاعة: {hall_names}")
                cell_content, cell_fill, cell_font = "", None, default_font
                if cell_parts:
                    cell_content = f"\n{'-'*20}\n".join(cell_parts)
                    if is_teaching_in_slot: cell_fill = assigned_subject_fill
                    if "حراسة:" in cell_content: cell_font = bold_font
                cell.value = cell_content
                apply_styles_to_cell(cell, cell_font, Alignment(horizontal='center', vertical='center', wrap_text=True), thin_border, cell_fill)
        ws.column_dimensions['A'].width = 15
        for i in range(2, len(all_times) + 2): ws.column_dimensions[get_column_letter(i)].width = 30
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="جداول_الحراسة_للأساتذة.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route('/api/data-template', methods=['GET'])
def export_data_template():
    """
    ينشئ ويصدر ملف Excel يحتوي على البيانات الحالية (أساتذة، قاعات، مواد، مستويات).
    إذا كانت قاعدة البيانات فارغة، سيتم تصدير قالب بالرؤوس فقط.
    --- نسخة محدثة مع أعمدة أوسع ---
    """
    try:
        conn = get_db_connection()
        # 1. استعلام لجلب كل البيانات الموجودة من قاعدة البيانات
        professors_data = pd.read_sql_query("SELECT name FROM professors", conn)
        halls_data = pd.read_sql_query("SELECT name, type FROM halls", conn)
        levels_data = pd.read_sql_query("SELECT name FROM levels", conn)
        subjects_data = pd.read_sql_query("SELECT s.name, l.name as level FROM subjects s JOIN levels l ON s.level_id = l.id", conn)
        conn.close()

        # 2. إعادة تسمية الأعمدة لتطابق رؤوس القالب
        professors_data.rename(columns={'name': 'اسم الأستاذ'}, inplace=True)
        halls_data.rename(columns={'name': 'اسم القاعة', 'type': 'نوع القاعة (صغيرة، متوسطة، كبيرة)'}, inplace=True)
        levels_data.rename(columns={'name': 'اسم المستوى الدراسي'}, inplace=True)
        subjects_data.rename(columns={'name': 'اسم المادة', 'level': 'المستوى الدراسي الخاص بها'}, inplace=True)

        # 3. كتابة البيانات إلى ملف إكسل باستخدام Pandas
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            professors_data.to_excel(writer, sheet_name='الأساتذة', index=False)
            halls_data.to_excel(writer, sheet_name='القاعات', index=False)
            levels_data.to_excel(writer, sheet_name='المستويات', index=False)
            subjects_data.to_excel(writer, sheet_name='المواد', index=False)

            # 4. تطبيق التنسيقات (اتجاه الصفحة وعرض الأعمدة)
            workbook = writer.book
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                # تفعيل واجهة من اليمين لليسار
                worksheet.sheet_view.rightToLeft = True
                
                # --- ✅  توسيع عرض الأعمدة ---
                worksheet.column_dimensions['A'].width = 40
                worksheet.column_dimensions['B'].width = 35
                worksheet.column_dimensions['C'].width = 25

        output.seek(0)
        # تغيير اسم الملف المصدّر ليعكس المحتوى
        return send_file(output, as_attachment=True, download_name='بيانات_الحراسة_الحالية.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        # طباعة الخطأ في الطرفية للمساعدة في التصحيح
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"فشل إنشاء الملف: {e}"}), 500

@app.route('/api/import-data', methods=['POST'])
def import_data_from_file():
    if 'file' not in request.files: return jsonify({"error": "لم يتم العثور على ملف."}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "لم يتم تحديد أي ملف."}), 400
    try:
        xls = pd.read_excel(file, sheet_name=None)
        conn = get_db_connection()
        cursor = conn.cursor()
        if 'الأساتذة' in xls:
            df_profs = xls['الأساتذة'].dropna(how='all')
            for index, row in df_profs.iterrows():
                try: cursor.execute("INSERT INTO professors (name) VALUES (?)", (str(row['اسم_الأستاذ']).strip(),))
                except sqlite3.IntegrityError: pass
        if 'القاعات' in xls:
            df_halls = xls['القاعات'].dropna(how='all')
            for index, row in df_halls.iterrows():
                try: cursor.execute("INSERT INTO halls (name, type) VALUES (?, ?)", (str(row['اسم_القاعة']).strip(), str(row['نوع_القاعة']).strip()))
                except sqlite3.IntegrityError: pass
        if 'المستويات' in xls:
            df_levels = xls['المستويات'].dropna(how='all')
            for index, row in df_levels.iterrows():
                try: cursor.execute("INSERT INTO levels (name) VALUES (?)", (str(row['اسم_المستوى_الدراسي']).strip(),))
                except sqlite3.IntegrityError: pass
        conn.commit()
        if 'المواد' in xls:
            level_map = {row['name']: row['id'] for row in cursor.execute("SELECT id, name FROM levels").fetchall()}
            df_subjects = xls['المواد'].dropna(how='all')
            for index, row in df_subjects.iterrows():
                name = str(row['اسم_المادة']).strip()
                level = str(row['المستوى_الدراسي_التابع_له']).strip()
                if name and level and level in level_map:
                    level_id = level_map[level]
                    exists = cursor.execute("SELECT 1 FROM subjects WHERE name = ? AND level_id = ?", (name, level_id)).fetchone()
                    if not exists: cursor.execute("INSERT INTO subjects (name, level_id) VALUES (?, ?)", (name, level_id))
        conn.commit()
        conn.close()
        return jsonify({"success": True, "message": "تم استيراد البيانات بنجاح. سيتم إعادة تحميل الصفحة."}), 201
    except Exception as e:
        return jsonify({"error": f"فشل تحليل الملف. تأكد من أن أسماء الأوراق والأعمدة متطابقة مع القالب. الخطأ: {e}"}), 500

@app.route('/api/generate-guard-schedule', methods=['POST'])
def generate_guard_schedule():
    settings = request.get_json()
    log_queue.put("بدء عملية إنشاء الجدول في الخلفية...")
    executor.submit(_run_schedule_logic_in_background, settings, log_queue)
    return jsonify({"success": True, "message": "بدأت عملية إنشاء الجدول. يرجى متابعة السجل الحي."})

@app.route('/stream-logs')
def stream_logs():
    def generate():
        while True:
            message = log_queue.get()
            if "DONE" in message: 
                yield f"data: {message}\n\n"
                break
            yield f"data: {message}\n\n"
    return Response(stream_with_context(generate()), mimetype='text/event-stream')

@app.route('/shutdown', methods=['POST'])
def shutdown():
    def do_shutdown():
        time.sleep(1)
        print("Shutdown request received. Terminating server.")
        executor.shutdown(wait=False, cancel_futures=True) 
        os.kill(os.getpid(), signal.SIGINT)
    threading.Thread(target=do_shutdown).start()
    return jsonify({"success": True, "message": "Shutdown signal sent."})

# أضف هذا الكود في نهاية ملف app.py قبل قسم التشغيل


# استبدل هذه الدالة بالكامل في ملف app.py
# استبدل هذه الدالة بالكامل في ملف app.py
@app.route('/api/export/word/all-exams', methods=['POST'])
def export_exams_word():
    schedule_data = request.get_json()
    if not schedule_data: return jsonify({"error": "No schedule data provided"}), 400
    
    conn = get_db_connection()
    assignments_rows = conn.execute('SELECT s.name as subj_name, l.name as level_name, p.name as prof_name FROM assignments a JOIN subjects s ON a.subject_id = s.id JOIN levels l ON s.level_id = l.id JOIN professors p ON a.professor_id = p.id').fetchall()
    settings_row = conn.execute("SELECT value FROM settings WHERE key = 'main_settings'").fetchone()
    conn.close()
    
    settings_data = json.loads(settings_row['value']) if settings_row else {}
    guards_large = int(settings_data.get('guardsLargeHall', 4))
    guards_medium = int(settings_data.get('guardsMediumHall', 2))
    guards_small = int(settings_data.get('guardsSmallHall', 1))

    subject_owners = {(row['subj_name'], row['level_name']): row['prof_name'] for row in assignments_rows}
    
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    margin = Cm(0.5)
    section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = margin, margin, margin, margin

    all_dates = sorted(schedule_data.keys())
    all_times = sorted({time for date_slots in schedule_data.values() for time in date_slots})
    all_levels = sorted({exam['level'] for slots in schedule_data.values() for exams in slots.values() for exam in exams})
    day_names = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
    
    headers = ["الفترة"] + [f"{day_names[datetime.strptime(d, '%Y-%m-%d').isoweekday() % 7]}\n{d}" for d in all_dates]

    for level in all_levels:
        data_grid = []
        for time in all_times:
            row_data = [time]
            for date in all_dates:
                exam = next((e for e in schedule_data.get(date, {}).get(time, []) if e['level'] == level), None)
                content = ""
                if exam:
                    owner = subject_owners.get((exam['subject'], exam['level']), "غير محدد")
                    content = f"{exam['subject']}\nأستاذ المادة: {owner}\n\nالحراسة:"

                    halls_by_type = defaultdict(list)
                    for h in exam.get('halls', []): halls_by_type[h['type']].append(h['name'])
                    
                    guards_copy = [g for g in exam.get('guards', []) if g != "**نقص**"]

                    if halls_by_type.get('كبيرة'):
                        num_guards_needed = len(halls_by_type['كبيرة']) * guards_large
                        g_list = guards_copy[:num_guards_needed]
                        guards_copy = guards_copy[num_guards_needed:]
                        hall_names = ", ".join(halls_by_type['كبيرة'])
                        guard_text = '\n'.join(g_list) if g_list else '(لا يوجد)'
                        content += f"\nالقاعة الكبيرة: {hall_names}\n{guard_text}"
                    
                    other_hall_names = halls_by_type.get('متوسطة', []) + halls_by_type.get('صغيرة', [])
                    if other_hall_names:
                        guard_text = '\n'.join(guards_copy) if guards_copy else '(لا يوجد)'
                        content += f"\nالقاعات الأخرى: {', '.join(other_hall_names)}\n{guard_text}"
                
                row_data.append(content)
            data_grid.append(row_data)
        
        create_word_document_with_table(doc, f"جدول امتحانات: {level}", headers, data_grid)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="جداول_الامتحانات.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# استبدل هذه الدالة بالكامل في ملف app.py
@app.route('/api/export/word/all-profs', methods=['POST'])
def export_profs_word():
    schedule_data = request.get_json()
    if not schedule_data: return jsonify({"error": "No schedule data provided"}), 400

    conn = get_db_connection()
    all_professors = sorted([p['name'] for p in conn.execute("SELECT name FROM professors").fetchall()])
    assignments_rows = conn.execute('SELECT p.name as prof_name, s.name as subj_name, l.name as level_name FROM assignments a JOIN professors p ON a.professor_id = p.id JOIN subjects s ON a.subject_id = s.id JOIN levels l ON s.level_id = l.id').fetchall()
    conn.close()

    prof_owned_subjects = defaultdict(set)
    for row in assignments_rows:
        prof_owned_subjects[row['prof_name']].add((row['subj_name'], row['level_name']))

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    margin = Cm(0.5)
    section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = margin, margin, margin, margin

    all_dates = sorted(schedule_data.keys())
    all_times = sorted({time for date_slots in schedule_data.values() for time in date_slots})
    day_names = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
    
    for prof_name in all_professors:
        title = f"جدول الحراسة: {prof_name}"
        headers = ["اليوم/التاريخ"] + all_times
        
        # --- بناء هيكل الجدول مع عنوان سليم ---
        heading = doc.add_heading(level=2); heading.clear()
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pPr = heading._p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi'); bidi.set(qn('w:val'), '1'); pPr.append(bidi)
        run = heading.add_run(title)
        font = run.font; font.rtl = True; font.name = 'Arial'

        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        table.autofit = False
        tbl_pr = table._element.xpath('w:tblPr')[0]
        bidi_visual_element = OxmlElement('w:bidiVisual')
        tbl_pr.append(bidi_visual_element)
        
        hdr_cells = table.rows[0].cells
        for i, header in enumerate(headers):
            p = hdr_cells[i].paragraphs[0]; p.text = ""
            run = p.add_run(header)
            font = run.font; font.rtl = True; font.name = 'Arial'
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.rtl = True

        has_any_duty = False
        for date in all_dates:
            row_cells = table.add_row().cells
            day_name = day_names[datetime.strptime(date, '%Y-%m-%d').isoweekday() % 7]
            
            p = row_cells[0].paragraphs[0]; p.text = ""
            run = p.add_run(f"{day_name}\n{date}"); run.font.rtl = True; run.font.name = 'Arial'; run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.rtl = True

            for i, time in enumerate(all_times, 1):
                cell_content_parts = []
                is_teaching_and_guarding = False
                is_teaching_only = False
                
                exams_in_slot = schedule_data.get(date, {}).get(time, [])
                
                for exam in exams_in_slot:
                    is_guarding = prof_name in exam.get('guards', [])
                    is_owner = (exam['subject'], exam['level']) in prof_owned_subjects.get(prof_name, set())

                    if is_guarding or is_owner:
                        has_any_duty = True
                        if is_guarding:
                            if is_owner: is_teaching_and_guarding = True
                            # ✅ التعديل هنا: استبدال القاعات بـ (حراسة)
                            cell_content_parts.append(f"{exam['subject']} ({exam['level']})\n(حراسة)")
                        elif is_owner:
                            is_teaching_only = True
                            cell_content_parts.append(f"{exam['subject']} ({exam['level']})\n(دون حراسة)")
                
                p = row_cells[i].paragraphs[0]; p.text = ""
                lines = "\n---\n".join(cell_content_parts).split('\n')
                for idx, line in enumerate(lines):
                    if idx > 0: p.add_run().add_break()
                    run = p.add_run(line)
                    font = run.font; font.rtl = True; font.name = 'Arial'
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.rtl = True
                
                shading_elm = OxmlElement('w:shd')
                if is_teaching_and_guarding:
                    shading_elm.set(qn('w:fill'), 'D4EDDA')
                    row_cells[i]._tc.get_or_add_tcPr().append(shading_elm)
                elif is_teaching_only:
                    shading_elm.set(qn('w:fill'), 'FFF3CD')
                    row_cells[i]._tc.get_or_add_tcPr().append(shading_elm)

        if has_any_duty:
             doc.add_page_break()
        else:
            doc._body.remove(table._element)
            doc._body.remove(heading._element)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="جداول_الحراسة_للأساتذة.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# أضف هذه الدالة الجديدة بالكامل في نهاية ملف app.py
@app.route('/api/export/word/all-profs-anonymous', methods=['POST'])
def export_profs_anonymous_word():
    schedule_data = request.get_json()
    if not schedule_data: return jsonify({"error": "No schedule data provided"}), 400

    conn = get_db_connection()
    all_professors = sorted([p['name'] for p in conn.execute("SELECT name FROM professors").fetchall()])
    assignments_rows = conn.execute('SELECT p.name as prof_name, s.name as subj_name, l.name as level_name FROM assignments a JOIN professors p ON a.professor_id = p.id JOIN subjects s ON a.subject_id = s.id JOIN levels l ON s.level_id = l.id').fetchall()
    conn.close()

    prof_owned_subjects = defaultdict(set)
    for row in assignments_rows:
        prof_owned_subjects[row['prof_name']].add((row['subj_name'], row['level_name']))

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    margin = Cm(0.5)
    section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = margin, margin, margin, margin

    all_dates = sorted(schedule_data.keys())
    all_times = sorted({time for date_slots in schedule_data.values() for time in date_slots})
    day_names = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
    
    for prof_name in all_professors:
        title = f"جدول الحراسة (مُبسَّط): {prof_name}"
        headers = ["اليوم/التاريخ"] + all_times
        
        heading = doc.add_heading(level=2); heading.clear()
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pPr = heading._p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi'); bidi.set(qn('w:val'), '1'); pPr.append(bidi)
        run = heading.add_run(title)
        font = run.font; font.rtl = True; font.name = 'Arial'

        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        table.autofit = False
        tbl_pr = table._element.xpath('w:tblPr')[0]
        bidi_visual_element = OxmlElement('w:bidiVisual')
        tbl_pr.append(bidi_visual_element)
        
        hdr_cells = table.rows[0].cells
        for i, header in enumerate(headers):
            p = hdr_cells[i].paragraphs[0]; p.text = ""
            run = p.add_run(header)
            font = run.font; font.rtl = True; font.name = 'Arial'
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.rtl = True

        has_any_duty = False
        for date in all_dates:
            row_cells = table.add_row().cells
            day_name = day_names[datetime.strptime(date, '%Y-%m-%d').isoweekday() % 7]
            
            p = row_cells[0].paragraphs[0]; p.text = ""
            run = p.add_run(f"{day_name}\n{date}"); run.font.rtl = True; run.font.name = 'Arial'; run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.rtl = True

            for i, time in enumerate(all_times, 1):
                cell_content_parts = []
                is_teaching_and_guarding = False
                is_teaching_only = False
                
                exams_in_slot = schedule_data.get(date, {}).get(time, [])
                
                for exam in exams_in_slot:
                    is_guarding = prof_name in exam.get('guards', [])
                    is_owner = (exam['subject'], exam['level']) in prof_owned_subjects.get(prof_name, set())

                    if is_guarding or is_owner:
                        has_any_duty = True
                        if is_guarding:
                            if is_owner:
                                is_teaching_and_guarding = True
                                cell_content_parts.append(f"{exam['subject']} ({exam['level']})\n(حراسة)")
                            else:
                                # ✅ التعديل هنا: إذا كان حارساً وليس صاحب المادة
                                cell_content_parts.append("(تكليف بحراسة)")
                        elif is_owner:
                            is_teaching_only = True
                            cell_content_parts.append(f"{exam['subject']} ({exam['level']})\n(دون حراسة)")
                
                p = row_cells[i].paragraphs[0]; p.text = ""
                lines = "\n---\n".join(cell_content_parts).split('\n')
                for idx, line in enumerate(lines):
                    if idx > 0: p.add_run().add_break()
                    run = p.add_run(line)
                    font = run.font; font.rtl = True; font.name = 'Arial'
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.rtl = True
                
                shading_elm = OxmlElement('w:shd')
                if is_teaching_and_guarding:
                    shading_elm.set(qn('w:fill'), 'D4EDDA')
                    row_cells[i]._tc.get_or_add_tcPr().append(shading_elm)
                elif is_teaching_only:
                    shading_elm.set(qn('w:fill'), 'FFF3CD')
                    row_cells[i]._tc.get_or_add_tcPr().append(shading_elm)

        if has_any_duty:
             doc.add_page_break()
        else:
            doc._body.remove(table._element)
            doc._body.remove(heading._element)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="جداول_الحراسة_المبسطة.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ================== الجزء الرابع: تشغيل البرنامج ==================
if __name__ == '__main__':
    init_db()
    def open_browser():
          webbrowser.open_new("http://127.0.0.1:5000")
    Timer(1, open_browser).start()
    serve(app, host='127.0.0.1', port=5000)
