import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog 
import datetime
import pandas as pd
import calendar 
import json 
import random 
import math
import logging

# ===== [1. 설정 및 상수] ===== #
TOSS_BLUE = '#0066FF'
WORK_DUTIES = ['D', 'E', 'N']
DAILY_LIMITS = {'D': 2, 'E': 2, 'N': 1}
PRESERVED_SHIFTS = ['V', 'v.25', 'v.0.5', 'DH', 'MD']
EDITABLE_SHIFTS = WORK_DUTIES + ['O'] + PRESERVED_SHIFTS + ['']

WINDOW_WIDTH, WINDOW_HEIGHT = 1600, 600
CURRENT_YEAR = datetime.datetime.now().year
CURRENT_MONTH = datetime.datetime.now().month
WORKER_LIST_FILE = 'worker_names.json'

DEFAULT_WORKERS = ["김토스", "이하나", "박우리", "최국민", "정신한", "조농협"]

logging.basicConfig(level=logging.ERROR,
                    format='%(asctime)s %(levelname)s:%(message)s')

# ===== [2. 전역 변수] ===== #
worker_names = []
is_head_nurse_mode = None
CURRENT_SCHEDULE_DF = pd.DataFrame()
CURRENT_SUMMARY_DF = pd.DataFrame()
current_tree = None
global_summary_frame = None
trace_id = None
MANUAL_EDITED_CELLS = set()
month_label_text = None

# ===== [3. 도우미 클래스 및 함수] ===== #

class RoundButton(tk.Canvas):
    """라운드 모서리 사용자 정의 버튼 위젯"""
    def __init__(self, master, text, command, corner_radius, fill_color, text_color, width, height, font_size):
        super().__init__(master, width=width, height=height, bd=0, highlightthickness=0, bg='white')
        self.command = command
        r, w, h, f = corner_radius, width, height, fill_color
        # 네 모서리 및 직사각형 그리기 (라운드 처리)
        self.create_arc(0, 0, r*2, r*2, start=90, extent=90, fill=f, outline=f)
        self.create_arc(w-r*2, 0, w, r*2, start=0, extent=90, fill=f, outline=f)
        self.create_arc(0, h-r*2, r*2, h, start=180, extent=90, fill=f, outline=f)
        self.create_arc(w-r*2, h-r*2, w, h, start=270, extent=90, fill=f, outline=f)
        self.create_rectangle(r, 0, w-r, h, fill=f, outline=f)
        self.create_rectangle(0, r, w, h-r, fill=f, outline=f)
        self.create_text(w/2, h/2, text=text, fill=text_color, font=('Malgun Gothic', font_size, 'bold'))
        self.bind("<Button-1>", self.on_click)
        self.bind("<Enter>", lambda e: self.config(cursor="hand2"))
        self.bind("<Leave>", lambda e: self.config(cursor=""))

    def on_click(self, event):
        if self.command: self.command()

# ===== [4. 데이터 관리 (저장/불러오기)] ===== #
def save_worker_names():
    """근무자 명단을 영구 저장"""
    try:
        with open(WORKER_LIST_FILE, 'w', encoding='utf-8') as f:
            json.dump(worker_names, f, ensure_ascii=False, indent=4)
    except Exception as e:
        logging.error(f"save_worker_names: {e}")

def load_worker_names():
    """근무자 명단 불러오거나 기본값 사용"""
    global worker_names
    try:
        with open(WORKER_LIST_FILE, 'r', encoding='utf-8') as f:
            loaded_names = json.load(f)
            if loaded_names and isinstance(loaded_names, list):
                worker_names = loaded_names
                return
    except Exception as e:
        logging.info(f"[초기값 사용] load_worker_names: {e}")
    worker_names = DEFAULT_WORKERS.copy()
    save_worker_names()

def save_schedule_to_excel():
    """근무표 및 통계 데이터를 엑셀로 저장"""
    global CURRENT_SCHEDULE_DF, CURRENT_SUMMARY_DF
    if CURRENT_SCHEDULE_DF.empty:
        messagebox.showwarning("경고", "근무표를 생성해야 저장할 수 있습니다."); return
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="근무표_데이터.xlsx")
    if not filepath: return
    try:
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            CURRENT_SCHEDULE_DF.to_excel(writer, sheet_name='근무표_스케줄', index=True, header=True)
            CURRENT_SUMMARY_DF.to_excel(writer, sheet_name='근무_통계', index=False)
            pd.DataFrame(worker_names, columns=['근무자 이름']).to_excel(writer, sheet_name='근무자_명단', index=False)
        messagebox.showinfo("저장 완료", f"근무표 및 통계 데이터가 엑셀 파일에 저장되었습니다:\n{filepath}")
    except Exception as e:
        messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다.\n오류: {e}")

def load_workers_from_excel(schedule_frame, year_var, month_var):
    """엑셀에서 근무자 명단만 불러오기"""
    filepath = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="근무자 명단 불러오기")
    if not filepath: return
    try:
        df = pd.read_excel(filepath, engine='openpyxl')
        if df.empty or len(df.columns) == 0:
            messagebox.showwarning("경고", "엑셀 파일에 데이터가 없습니다."); return
        new_workers = df.iloc[:, 0].dropna().astype(str).tolist()
        if not new_workers:
            messagebox.showwarning("경고", "파일에 유효한 근무자 이름이 없습니다."); return
        global worker_names
        worker_names = new_workers
        update_gui_after_worker_change(schedule_frame, year_var, month_var)
        messagebox.showinfo("불러오기 완료", "근무자 명단이 업데이트되었습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"엑셀 파일을 읽는 중 오류가 발생했습니다.\n오류: {e}")

# ===== [5. 근무자 UI 관리] ===== #
def update_gui_after_worker_change(schedule_frame, year_var, month_var):
    """근무자 변경시 UI 및 데이터 동기화"""
    global MANUAL_EDITED_CELLS
    MANUAL_EDITED_CELLS.clear()
    display_initial_schedule_table(schedule_frame, year_var, month_var)
    save_worker_names()

def add_worker(root, schedule_frame, year_var, month_var):
    global worker_names
    new_name = simpledialog.askstring("근무자 추가", "추가할 근무자의 이름을 입력하세요:", parent=root)
    if new_name and new_name.strip():
        fixed_name = new_name.strip()
        if fixed_name not in worker_names:
            worker_names.append(fixed_name)
            messagebox.showinfo("성공", f"근무자 '{fixed_name}'이(가) 추가되었습니다.")
            update_gui_after_worker_change(schedule_frame, year_var, month_var)
        else:
            messagebox.showwarning("중복", f"이미 존재하는 근무자입니다: '{fixed_name}'")

def modify_worker(root, schedule_frame, year_var, month_var):
    global worker_names
    old_name = simpledialog.askstring("근무자 수정", "수정할 근무자의 현재 이름을 입력하세요:", parent=root)
    if old_name and old_name in worker_names:
        new_name = simpledialog.askstring("새 이름 입력", f"'{old_name}'의 새로운 이름을 입력하세요:", parent=root)
        if new_name and new_name.strip():
            worker_names[worker_names.index(old_name)] = new_name.strip()
            messagebox.showinfo("성공", f"'{old_name}'이(가) '{new_name.strip()}'(으)로 수정되었습니다.")
            update_gui_after_worker_change(schedule_frame, year_var, month_var)
    else:
        messagebox.showwarning("오류", "명단에 없는 이름입니다.")

def delete_worker(root, schedule_frame, year_var, month_var):
    global worker_names
    name_to_delete = simpledialog.askstring("근무자 삭제", "삭제할 근무자의 이름을 입력하세요:", parent=root)
    if name_to_delete and name_to_delete in worker_names:
        worker_names.remove(name_to_delete)
        messagebox.showinfo("성공", f"근무자 '{name_to_delete}'이(가) 명단에서 삭제되었습니다.")
        update_gui_after_worker_change(schedule_frame, year_var, month_var)
    else:
        messagebox.showwarning("오류", "명단에 없는 이름입니다.")

def worker_reorder_dialog(root, schedule_frame, year_var, month_var):
    def move_worker(direction):
        try:
            selected_item = worker_tree.selection()[0]
            current_index = worker_tree.index(selected_item)
        except IndexError:
            messagebox.showwarning("경고", "먼저 순서를 바꿀 근무자를 선택해 주세요.", parent=dialog)
            return
        target_index = current_index + direction
        if 0 <= target_index < len(worker_names):
            worker_names.insert(target_index, worker_names.pop(current_index))
            worker_tree.move(selected_item, '', target_index)
            worker_tree.selection_set(selected_item)
        else:
            messagebox.showwarning("경고", "더 이상 이동할 수 없습니다.", parent=dialog)

    dialog = tk.Toplevel(root); dialog.title("근무자 순서 변경 (▲/▼ 이동)"); dialog.geometry("400x400")
    dialog.transient(root); dialog.grab_set()
    dialog.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2) - (dialog.winfo_width() // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (dialog.winfo_height() // 2)
    dialog.geometry(f'+{x}+{y}')
    tk.Label(dialog, text="근무자 목록", font=('Malgun Gothic', 14, 'bold')).pack(pady=10)
    main_frame = ttk.Frame(dialog); main_frame.pack(padx=10, pady=5, fill='both', expand=True)
    worker_tree = ttk.Treeview(main_frame, columns=['Name'], show='headings', selectmode='browse')
    worker_tree.heading('Name', text='근무자 이름')
    worker_tree.column('Name', anchor='center', width=200)
    for name in worker_names: worker_tree.insert('', 'end', values=(name,), tags=(name,))
    worker_tree.pack(side='left', fill='both', expand=True)
    button_frame = ttk.Frame(main_frame); button_frame.pack(side='right', padx=(5, 0))
    ttk.Button(button_frame, text="▲ 위로", command=lambda: move_worker(-1)).pack(pady=5, fill='x')
    ttk.Button(button_frame, text="▼ 아래로", command=lambda: move_worker(1)).pack(pady=5, fill='x')
    def on_ok():
        update_gui_after_worker_change(schedule_frame, year_var, month_var)
        dialog.destroy()
        messagebox.showinfo("순서 변경 완료", "근무자 순서가 성공적으로 적용되었습니다.")
    ttk.Button(dialog, text="확인 및 적용", command=on_ok).pack(side='right', padx=10, pady=10)
    ttk.Button(dialog, text="닫기", command=dialog.destroy).pack(side='right', pady=10)
    root.wait_window(dialog)

# ===== [6. 근무표/통계 및 UI 표시] ===== #
def get_month_days(year, month):
    year, month = int(year), int(month)
    try: _, last_day = calendar.monthrange(year, month)
    except ValueError: last_day = 30
    weekday_names_kr = ["월", "화", "수", "목", "금", "토", "일"]
    day_columns = []
    for day in range(1, last_day + 1):
        try:
            date_obj = datetime.date(year, month, day)
            weekday = weekday_names_kr[date_obj.weekday()]
            day_columns.append(f"{month}/{day} ({weekday})")
        except:
            day_columns.append(f"{month}/{day} (?)")
    return year, month, day_columns

def display_schedule_table(schedule_frame, df, year, month):
    """DataFrame을 Treeview로 표시하며, 헤더 스타일 변경"""
    for widget in schedule_frame.winfo_children(): widget.destroy()
    if df.empty:
        tk.Label(schedule_frame, text="근무표 데이터가 없습니다.", font=('Malgun Gothic', 14)).pack(pady=20); return
    ttk.Style().configure("Custom.Treeview.Heading", background="white", foreground="#A9A9A9", font=('Malgun Gothic', 10, 'bold'))
    tree_frame = ttk.Frame(schedule_frame); tree_frame.pack(fill='both', expand=True)
    tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL); tree_scroll_y.pack(side='right', fill='y')
    tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL); tree_scroll_x.pack(side='bottom', fill='x')
    columns = ["근무자"] + list(df.columns)
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings', yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, style="Custom.Treeview")
    tree_scroll_y.config(command=tree.yview); tree_scroll_x.config(command=tree.xview)
    tree.heading("근무자", text="근무자", anchor='center'); tree.column("근무자", width=100, anchor='center', stretch=tk.NO)
    for col in df.columns:
        day_and_weekday = col.split('/', 1)[-1].strip()
        tree.heading(col, text=day_and_weekday, anchor='center'); tree.column(col, width=60, anchor='center', stretch=tk.NO)
    for worker, row in df.iterrows():
        tree.insert('', 'end', values=[worker] + row.tolist(), tags=(worker,))
    tree.pack(fill='both', expand=True)
    tree.bind("<Button-1>", lambda event: start_schedule_edit(event, tree))
    global month_label_text
    month_label_text.set(f"🗓️ {year}년 {month}월 근무표")

def display_summary_table(summary_frame, summary_df):
    for widget in summary_frame.winfo_children(): widget.destroy()
    if summary_df.empty:
        tk.Label(summary_frame, text="근무표 생성 후\n통계가 표시됩니다.", font=('Malgun Gothic', 12), bg='white').pack(pady=100, padx=50); return
    ttk.Style().configure("Summary.Treeview.Heading", background="#E8F0FE", foreground="#333333", font=('Malgun Gothic', 9, 'bold'))
    tk.Label(summary_frame, text="근무 합산 통계", font=('Malgun Gothic', 12, 'bold'), bg='white').pack(pady=(0, 5))
    tree_frame = ttk.Frame(summary_frame); tree_frame.pack(fill='both', expand=True)
    columns = list(summary_df.columns)
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style="Summary.Treeview")
    column_widths = {'근무자': 80, '전월 연차': 60, '총 연차': 60, '총 근무': 60, 'D': 40, 'DH': 40, 'E': 40, 'MD': 40, 'N': 40, 'Off': 40, 'V': 40, 'v.25': 40, 'v.0.5': 40}
    for col in columns:
        tree.heading(col, text=col.replace('_', ' '), anchor='center')
        tree.column(col, width=column_widths.get(col, 50), anchor='center', stretch=tk.NO)
    for index, row in summary_df.iterrows(): tree.insert('', 'end', values=row.tolist())
    tree.pack(fill='both', expand=True)

def display_initial_schedule_table(schedule_frame, year_var, month_var):
    global CURRENT_SCHEDULE_DF, MANUAL_EDITED_CELLS
    try: selected_year, selected_month = year_var.get(), month_var.get()
    except: return
    if not worker_names:
        for widget in schedule_frame.winfo_children(): widget.destroy()
        tk.Label(schedule_frame, text="근무자 관리 메뉴에서 근무자를 먼저 추가해 주세요.", font=('Malgun Gothic', 14)).pack(pady=100)
        global month_label_text
        month_label_text.set(f"🗓️ {selected_year}년 {selected_month}월 근무표")
        return
    year, month, day_columns = get_month_days(selected_year, selected_month)
    # 데이터 변경시 수동 편집 기록 초기화
    if CURRENT_SCHEDULE_DF.empty or list(CURRENT_SCHEDULE_DF.columns) != day_columns or list(CURRENT_SCHEDULE_DF.index) != worker_names:
        MANUAL_EDITED_CELLS.clear()
        initial_data = {name: [''] * len(day_columns) for name in worker_names}
        df_initial = pd.DataFrame(initial_data).transpose(); df_initial.columns = day_columns
        CURRENT_SCHEDULE_DF = df_initial
    display_schedule_table(schedule_frame, CURRENT_SCHEDULE_DF, year, month)

def generate_schedule_summary(df_schedule):
    """근무표 데이터에서 유형별 집계 통계 생성"""
    if df_schedule.empty: return pd.DataFrame()
    DUTY_TYPES = ['D', 'DH', 'E', 'MD', 'N', 'Off', 'V', 'v.25', 'v.0.5']
    summary_df = df_schedule.stack().groupby(level=0).value_counts().unstack(fill_value=0)
    summary_df['Off'] = summary_df.get('O', 0) + summary_df.get('Off', 0)
    summary_df = summary_df.drop(columns=['O'], errors='ignore')
    for col in DUTY_TYPES:
        if col not in summary_df.columns: summary_df[col] = 0
    summary_df = summary_df.reset_index(names=['근무자'])
    summary_df['총 근무'] = summary_df[['D', 'DH', 'E', 'MD', 'N']].sum(axis=1)
    summary_df['전월 연차'] = 2; summary_df['총 연차'] = 21.5
    final_cols = ['근무자', '전월 연차', '총 연차', '총 근무', 'D', 'DH', 'E', 'MD', 'N', 'Off', 'V', 'v.25', 'v.0.5']
    return summary_df.reindex(columns=final_cols)

def update_schedule_cell(event, tree, combobox, item_id, column_id, col_name, worker_name):
    """Combobox 선택 후 Treeview와 DataFrame을 업데이트, 수동 편집 추적"""
    global CURRENT_SCHEDULE_DF, CURRENT_SUMMARY_DF, global_summary_frame, MANUAL_EDITED_CELLS
    new_value = combobox.get()
    edit_key = (worker_name, col_name)
    if new_value != '':
        MANUAL_EDITED_CELLS.add(edit_key)
    else:
        if edit_key in MANUAL_EDITED_CELLS:
            MANUAL_EDITED_CELLS.remove(edit_key)
    tree.set(item_id, column_id, new_value)
    if not CURRENT_SCHEDULE_DF.empty and worker_name in CURRENT_SCHEDULE_DF.index and col_name in CURRENT_SCHEDULE_DF.columns:
        CURRENT_SCHEDULE_DF.loc[worker_name, col_name] = new_value
        if global_summary_frame:
            CURRENT_SUMMARY_DF = generate_schedule_summary(CURRENT_SCHEDULE_DF)
            display_summary_table(global_summary_frame, CURRENT_SUMMARY_DF)
    combobox.destroy()

def start_schedule_edit(event, tree):
    """Treeview 셀 클릭시 수동 수정 Combobox 띄우기"""
    try:
        if tree.identify_region(event.x, event.y) != "cell" or tree.identify_column(event.x) == '#1': return
        column_id = tree.identify_column(event.x); item_id = tree.identify_row(event.y)
        col_name_full = tree.cget('columns')[int(column_id.replace('#', '')) - 1]
        worker_name = tree.item(item_id, 'values')[0]; current_value = tree.set(item_id, column_id)
        bbox = tree.bbox(item_id, column_id);
        if not bbox: return
        x, y, width, height = bbox
        global current_tree
        if hasattr(current_tree, 'editor_widget') and current_tree.editor_widget.winfo_exists():
            current_tree.editor_widget.destroy()
        combobox = ttk.Combobox(tree, values=EDITABLE_SHIFTS, width=width, font=('Malgun Gothic', 10), state='readonly')
        combobox.set(current_value)
        combobox.place(x=x, y=y, width=width, height=height)
        combobox.bind("<<ComboboxSelected>>", lambda e: update_schedule_cell(e, tree, combobox, item_id, column_id, col_name_full, worker_name))
        combobox.bind("<FocusOut>", lambda e: combobox.destroy() if e.widget == combobox else None)
        combobox.bind("<Return>", lambda e: update_schedule_cell(e, tree, combobox, item_id, column_id, col_name_full, worker_name))
        combobox.focus_set(); current_tree = tree; current_tree.editor_widget = combobox
    except Exception as e: logging.error(f"[start_schedule_edit] {e}")

def generate_monthly_schedule(year, month):
    """자동 근무표 생성 핵심 알고리즘"""
    global CURRENT_SCHEDULE_DF, MANUAL_EDITED_CELLS
    year, month, day_columns = get_month_days(year, month); last_day = len(day_columns)
    if not worker_names: return pd.DataFrame(), year, month
    schedule_data = {name: [''] * last_day for name in worker_names}
    if not CURRENT_SCHEDULE_DF.empty:
        df_temp = CURRENT_SCHEDULE_DF.copy().fillna('').astype(str)
        for worker, col_name in MANUAL_EDITED_CELLS:
            if worker in df_temp.index and col_name in df_temp.columns:
                manual_value = df_temp.loc[worker, col_name]
                try:
                    day_index = day_columns.index(col_name)
                    schedule_data[worker][day_index] = manual_value
                except ValueError: pass
    hn_name = worker_names[0] if worker_names else None
    if is_head_nurse_mode.get() and hn_name and hn_name in schedule_data:
        start_date = datetime.date(year, month, 1)
        for day_index in range(last_day):
            if schedule_data[hn_name][day_index] == '':
                weekday = (start_date + datetime.timedelta(days=day_index)).weekday()
                schedule_data[hn_name][day_index] = 'D' if 0 <= weekday <= 4 else 'O'
    num_workers_for_duty = len(worker_names) - (1 if is_head_nurse_mode.get() and hn_name else 0)
    target_duty_count_per_worker = max(1, math.ceil(last_day / 7 * 5 * 3 / num_workers_for_duty)) if num_workers_for_duty > 0 and WORK_DUTIES else 0
    duty_counts = {name: {d: schedule_data[name].count(d) for d in WORK_DUTIES} for name in worker_names}
    for day_index in range(last_day):
        daily_duty_counts = {'D': 0, 'E': 0, 'N': 0}
        workers_to_schedule = []
        for name in worker_names:
            current_duty = schedule_data[name][day_index]
            if current_duty in WORK_DUTIES:
                daily_duty_counts[current_duty] += 1
            elif current_duty == '':
                workers_to_schedule.append(name)
        if is_head_nurse_mode.get() and hn_name and hn_name in workers_to_schedule:
            workers_to_schedule.remove(hn_name)
        random.shuffle(workers_to_schedule)
        for name in workers_to_schedule:
            assigned_duty = ''; prev_duty = schedule_data[name][day_index - 1] if day_index > 0 else ''
            if prev_duty == 'N': assigned_duty = 'O'
            elif day_index >= 3:
                last_3 = schedule_data[name][day_index-3:day_index]
                if not any(d in ['O', 'V', 'v.25', 'v.0.5', ''] for d in last_3) and all(d in WORK_DUTIES for d in last_3):
                    assigned_duty = 'O'
            if not assigned_duty:
                target_rotation = ''
                if prev_duty == 'D': target_rotation = 'E'
                elif prev_duty == 'E': target_rotation = 'N'
                else: target_rotation = sorted(WORK_DUTIES, key=lambda d: duty_counts[name].get(d, 0))[0]
                under_limit = []
                for duty in WORK_DUTIES:
                    is_daily_full = daily_duty_counts.get(duty, 0) >= DAILY_LIMITS.get(duty, float('inf'))
                    is_worker_full = duty_counts[name][duty] >= target_duty_count_per_worker + 1
                    if not is_daily_full and not is_worker_full: under_limit.append(duty)
                if not under_limit: assigned_duty = 'O'
                elif target_rotation in under_limit: assigned_duty = target_rotation
                else: assigned_duty = random.choice(under_limit)
            schedule_data[name][day_index] = assigned_duty
            if assigned_duty in WORK_DUTIES:
                duty_counts[name][assigned_duty] += 1
                daily_duty_counts[assigned_duty] += 1
    df = pd.DataFrame({name: schedule_data[name] for name in worker_names}).transpose(); df.columns = day_columns
    return df, year, month

def generate_and_display(schedule_frame, summary_frame, year_var, month_var):
    global CURRENT_SCHEDULE_DF, CURRENT_SUMMARY_DF
    if not worker_names:
        messagebox.showwarning("경고", "근무자가 최소 1명 이상 등록되어야 합니다."); return
    try: selected_year, selected_month = year_var.get(), month_var.get()
    except tk.TclError: messagebox.showerror("오류", "올바른 년도와 월을 선택해 주세요."); return
    df_schedule, year, month = generate_monthly_schedule(selected_year, selected_month)
    display_schedule_table(schedule_frame, df_schedule, year, month)
    summary_df = generate_schedule_summary(df_schedule)
    display_summary_table(summary_frame, summary_df)
    CURRENT_SCHEDULE_DF = df_schedule; CURRENT_SUMMARY_DF = summary_df

def clear_schedule(schedule_frame, summary_frame, year_var, month_var):
    global CURRENT_SCHEDULE_DF, CURRENT_SUMMARY_DF, MANUAL_EDITED_CELLS
    if not worker_names: messagebox.showwarning("경고", "초기화할 근무자 명단이 없습니다."); return
    try:
        MANUAL_EDITED_CELLS.clear()
        selected_year, selected_month = year_var.get(), month_var.get()
        year, month, day_columns = get_month_days(selected_year, selected_month)
        initial_data = {name: [''] * len(day_columns) for name in worker_names}
        df_initial = pd.DataFrame(initial_data).transpose(); df_initial.columns = day_columns
        CURRENT_SCHEDULE_DF = df_initial
        display_schedule_table(schedule_frame, CURRENT_SCHEDULE_DF, year, month)
        CURRENT_SUMMARY_DF = pd.DataFrame()
        display_summary_table(summary_frame, CURRENT_SUMMARY_DF)
    except Exception as e:
        messagebox.showerror("초기화 오류", f"근무표 초기화 중 오류가 발생했습니다: {e}")

# ===== [7. 메인 UI 및 이벤트 연결] ===== #
def setup_main_window():
    load_worker_names()
    global month_label_text, is_head_nurse_mode, trace_id
    root = tk.Tk(); root.title("📅 근무표 생성 시스템"); root.configure(bg='white')
    is_head_nurse_mode = tk.BooleanVar(value=True)
    screen_width = root.winfo_screenwidth(); screen_height = root.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (WINDOW_WIDTH / 2))
    y_cordinate = int((screen_height / 2) - (WINDOW_HEIGHT / 2))
    root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x_cordinate}+{y_cordinate}")
    root.protocol("WM_DELETE_WINDOW", lambda: (save_worker_names(), root.destroy()))
    style = ttk.Style(); style.theme_use('default')
    style.configure('Toss.TLabel', font=('Malgun Gothic', 16, 'bold'), background='white', foreground='#333333')
    style.configure('Toss.TFrame', background='white')
    year_var = tk.IntVar(value=CURRENT_YEAR); month_var = tk.IntVar(value=CURRENT_MONTH)
    menu_bar = tk.Menu(root); root.config(menu=menu_bar)
    file_menu = tk.Menu(menu_bar, tearoff=0); menu_bar.add_cascade(label="파일", menu=file_menu)
    file_menu.add_command(label="종료", command=root.destroy)
    worker_menu = tk.Menu(menu_bar, tearoff=0); menu_bar.add_cascade(label="근무자 관리", menu=worker_menu)
    worker_management_submenu = tk.Menu(worker_menu, tearoff=0); worker_menu.add_cascade(label="근무자 관리", menu=worker_management_submenu)
    worker_management_submenu.add_command(label="추가", command=lambda: add_worker(root, schedule_frame, year_var, month_var))
    worker_management_submenu.add_command(label="수정", command=lambda: modify_worker(root, schedule_frame, year_var, month_var))
    worker_management_submenu.add_command(label="삭제", command=lambda: delete_worker(root, schedule_frame, year_var, month_var))
    worker_management_submenu.add_separator()
    worker_management_submenu.add_command(label="순서 변경", command=lambda: worker_reorder_dialog(root, schedule_frame, year_var, month_var))
    def toggle_head_nurse_mode_command():
        if not worker_names:
            messagebox.showwarning("경고", "근무자 명단이 비어있습니다."); is_head_nurse_mode.set(False); return
        display_initial_schedule_table(schedule_frame, year_var, month_var)
        messagebox.showinfo("모드 변경", f"상위 근무자({worker_names[0]}) 주간 근무 모드가 {'적용' if is_head_nurse_mode.get() else '해제'}되었습니다.")
    worker_menu.add_separator()
    worker_menu.add_checkbutton(label="수선생님(상위 1인) 주간 근무 모드", onvalue=True, offvalue=False, variable=is_head_nurse_mode, command=toggle_head_nurse_mode_command)
    data_menu = tk.Menu(menu_bar, tearoff=0); menu_bar.add_cascade(label="데이터", menu=data_menu)
    data_menu.add_command(label="데이터 저장 (.xlsx)", command=save_schedule_to_excel)
    data_menu.add_command(label="데이터 불러오기 (.xlsx)", command=lambda: load_workers_from_excel(schedule_frame, year_var, month_var))
    # 년/월 선택 위젯
    control_frame = ttk.Frame(root, style='Toss.TFrame'); control_frame.pack(pady=(20, 5), padx=20)
    tk.Label(control_frame, text="년도:", font=('Malgun Gothic', 12, 'bold'), bg='white').pack(side='left', padx=(0, 5))
    ttk.Spinbox(control_frame, from_=CURRENT_YEAR - 5, to=CURRENT_YEAR + 5, textvariable=year_var, width=5, font=('Malgun Gothic', 12)).pack(side='left', padx=(0, 15))
    tk.Label(control_frame, text="월:", font=('Malgun Gothic', 12, 'bold'), bg='white').pack(side='left', padx=(0, 5))
    ttk.Spinbox(control_frame, from_=0, to=13, textvariable=month_var, width=3, font=('Malgun Gothic', 12), wrap=True).pack(side='left', padx=(0, 15))
    def on_date_change_cb(*args):
        global trace_id
        try: current_year, current_month = year_var.get(), month_var.get()
        except Exception: return
        if trace_id:
            try: month_var.trace_remove('write', trace_id)
            except tk.TclError: pass
        if current_month > 12: year_var.set(current_year + 1); month_var.set(1)
        elif current_month < 1: year_var.set(current_year - 1); month_var.set(12)
        trace_id = month_var.trace_add('write', on_date_change_cb)
        try: display_initial_schedule_table(schedule_frame, year_var, month_var)
        except Exception as e: logging.error(f"[on_date_change_cb] {e}")
    # 메인 타이틀
    month_label_text = tk.StringVar()
    ttk.Label(root, textvariable=month_label_text, style='Toss.TLabel').pack(pady=5)
    # 좌우 컨테이너
    main_content_frame = ttk.Frame(root, style='Toss.TFrame'); main_content_frame.pack(fill='both', expand=True, padx=20, pady=10)
    global schedule_frame; schedule_frame = ttk.Frame(main_content_frame, relief='flat', borderwidth=0, padding=10, style='Toss.TFrame')
    schedule_frame.pack(side='left', fill='both', expand=True)
    summary_frame = ttk.Frame(main_content_frame, relief='flat', borderwidth=0, padding=10, style='Toss.TFrame')
    summary_frame.pack(side='left', fill='y', padx=(10, 0))
    global global_summary_frame; global_summary_frame = summary_frame
    # 하단 버튼 컨테이너
    button_container = ttk.Frame(root, style='Toss.TFrame'); button_container.pack(pady=30)
    RoundButton(button_container, text="✨ 근무표 생성 및 표시", command=lambda: generate_and_display(schedule_frame, summary_frame, year_var, month_var), corner_radius=20, fill_color=TOSS_BLUE, text_color='white', width=250, height=50, font_size=14).pack(side='left', padx=10)
    RoundButton(button_container, text="🗑️ 근무표 초기화", command=lambda: clear_schedule(schedule_frame, summary_frame, year_var, month_var), corner_radius=20, fill_color='#DCDCDC', text_color='#333333', width=250, height=50, font_size=14).pack(side='left', padx=10)
    # trace 등록 (년/월 변경)
    trace_id = month_var.trace_add("write", on_date_change_cb)
    display_initial_schedule_table(schedule_frame, year_var, month_var)
    tk.Label(summary_frame, text="근무표 생성 후\n통계가 표시됩니다.", font=('Malgun Gothic', 12), bg='white').pack(pady=100, padx=50)
    footer_frame = ttk.Frame(root, style='Toss.TFrame'); footer_frame.pack(side='bottom', fill='x', padx=10, pady=(0, 5))
    tk.Label(footer_frame, text="made by TKㅣver.241124", font=('Malgun Gothic', 9), fg='#AAAAAA', bg='white').pack(side='right', padx=10)
    root.mainloop()

# --- 실행 진입점 --- #
if __name__ == "__main__":
    setup_main_window()