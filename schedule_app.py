import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import datetime
import pandas as pd
import calendar
import json
import random
import math
import logging
import io # <-- [FIX 1] FutureWarning 해결을 위한 io 모듈 추가

# ==============================================================================
# 1. 설정 및 상수
# ==============================================================================
TOSS_BLUE = '#0066FF'
WORK_DUTIES = ['D', 'E', 'N', 'DH']
DAILY_LIMITS = {'D': 2, 'E': 2, 'N': 1, 'DH': 1} # DH는 별도 카운트
PRESERVED_SHIFTS = ['V', 'v.25', 'v.0.5', 'MD']
EDITABLE_SHIFTS = ['D', 'E', 'N', 'O', 'V', 'v.25', 'v.0.5', 'MD', 'DH', '']

WINDOW_WIDTH, WINDOW_HEIGHT = 1600, 600
CURRENT_YEAR = datetime.datetime.now().year
CURRENT_MONTH = datetime.datetime.now().month
WORKER_LIST_FILE = 'worker_names.json'
PREV_MONTH_SCHEDULE_FILE = 'prev_month_schedule.json'
MONTHLY_SCHEDULES_FILE = 'monthly_schedules.json' 
WORKER_V_FILE = 'worker_v_data.json' 

DEFAULT_WORKERS = ["도은아", "구진아", "김정화", "이현주", "강효선", "천보람", "지연정", "이소라", "김수빈", "문수빈", "최민정", "문오순"]

# 직책/구분 상수 정의
WORKER_CATEGORIES = ['일반', '수선생님', 'C', 'A']
DEFAULT_CATEGORY = '일반'
WORKER_CATEGORIES_FILE = 'worker_categories.json'

logging.basicConfig(level=logging.ERROR,
                    format='%(asctime)s %(levelname)s:%(message)s')

# ==============================================================================
# 2. 메인 애플리케이션 클래스 (OOP 구조 도입)
# ==============================================================================

class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.worker_names = []
        self.worker_categories_map = {} 
        self.monthly_schedules = {} 
        self.current_schedule_df = pd.DataFrame()
        self.current_summary_df = pd.DataFrame()
        self.manual_edited_cells = set() 
        self.trace_id = None 
        self.current_tree = None
        self.prev_month_last_day_duties = {} 
        self.worker_v_map = {} 

        # [데이터 로드]
        self.load_worker_names()
        self.load_prev_month_schedule() 
        self.load_worker_categories() 
        self.load_worker_v_data() 

        # [UI 변수]
        self.year_var = tk.IntVar(value=CURRENT_YEAR)
        self.month_var = tk.IntVar(value=CURRENT_MONTH)
        self.month_label_text = tk.StringVar()
        self.is_head_nurse_mode = tk.BooleanVar(value=True)

        # [UI 설정]
        self.setup_main_window()

    # ----------------------------------------------------------------------
    # [데이터 관리: 저장/불러오기]
    # ----------------------------------------------------------------------
    
    def save_all_schedules(self):
        try:
            self.save_current_schedule_to_memory() 
            
            data_to_save = {
                'schedules': {
                    f"{year}-{month:02d}": 
                    self.monthly_schedules[year][month].to_json(orient='split', index=True, date_format='iso')
                    for year in self.monthly_schedules for month in self.monthly_schedules[year]
                }
            }
            with open(MONTHLY_SCHEDULES_FILE, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=4)
            logging.info(f"save_all_schedules: {len(data_to_save['schedules'])}개의 근무표 저장 완료.")
        except Exception as e:
            logging.error(f"save_all_schedules: {e}")

    def load_all_schedules(self):
        try:
            with open(MONTHLY_SCHEDULES_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            loaded_schedules = data.get('schedules', {})
            self.monthly_schedules = {}
            for key, json_data in loaded_schedules.items():
                year_str, month_str = key.split('-')
                year = int(year_str)
                month = int(month_str)
                
                # [FIX 1] FutureWarning 해결: io.StringIO() 사용
                df = pd.read_json(io.StringIO(json_data), orient='split') 
                
                if year not in self.monthly_schedules:
                    self.monthly_schedules[year] = {}
                self.monthly_schedules[year][month] = df
            
            logging.info(f"load_all_schedules: 총 {len(loaded_schedules)}개의 근무표 불러오기 완료.")
        except FileNotFoundError:
            logging.info("load_all_schedules: 저장된 근무표 파일 없음.")
        except Exception as e:
            logging.error(f"load_all_schedules: {e}")

    def save_current_schedule_to_memory(self):
        year, month = self.year_var.get(), self.month_var.get()
        if not self.current_schedule_df.empty:
            if year not in self.monthly_schedules: self.monthly_schedules[year] = {}
            self.monthly_schedules[year][month] = self.current_schedule_df.copy()
            logging.info(f"Schedule for {year}-{month:02d} saved to memory.")

    def load_schedule_from_memory(self, year, month):
        if year in self.monthly_schedules and month in self.monthly_schedules[year]:
            df = self.monthly_schedules[year][month].copy()
            self.current_schedule_df = df
            self.display_schedule_table(df)
            return True
        return False

    def save_worker_names(self):
        try:
            with open(WORKER_LIST_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.worker_names, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.error(f"save_worker_names: {e}")

    def load_worker_names(self):
        try:
            with open(WORKER_LIST_FILE, 'r', encoding='utf-8') as f:
                self.worker_names = json.load(f)
        except FileNotFoundError:
            self.worker_names = DEFAULT_WORKERS.copy()
            self.save_worker_names()
        except Exception as e:
            logging.error(f"load_worker_names: {e}")
            self.worker_names = DEFAULT_WORKERS.copy()

    def save_worker_categories(self):
        try:
            with open(WORKER_CATEGORIES_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.worker_categories_map, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.error(f"save_worker_categories: {e}")

    def load_worker_categories(self):
        try:
            with open(WORKER_CATEGORIES_FILE, 'r', encoding='utf-8') as f:
                loaded_map = json.load(f)
                self.worker_categories_map = {
                    name: loaded_map.get(name, DEFAULT_CATEGORY) 
                    for name in self.worker_names
                }
        except Exception:
            self.worker_categories_map = {name: DEFAULT_CATEGORY for name in self.worker_names}
            self.save_worker_categories()

    def save_prev_month_schedule(self):
        try:
            with open(PREV_MONTH_SCHEDULE_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.prev_month_last_day_duties, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.error(f"save_prev_month_schedule: {e}")

    def load_prev_month_schedule(self):
        try:
            with open(PREV_MONTH_SCHEDULE_FILE, 'r', encoding='utf-8') as f:
                self.prev_month_last_day_duties = json.load(f)
        except Exception:
            self.prev_month_last_day_duties = {}

    def save_worker_v_data(self):
        """근무자별 연차 일수 데이터를 영구 저장"""
        try:
            data_to_save = {str(k): v for k, v in self.worker_v_map.items()}
            with open(WORKER_V_FILE, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.error(f"save_worker_v_data: {e}")

    def load_worker_v_data(self):
        """근무자별 연차 일수 데이터를 불러오거나 기본값 사용"""
        try:
            with open(WORKER_V_FILE, 'r', encoding='utf-8') as f:
                loaded_map = json.load(f)
                self.worker_v_map = {int(k): v for k, v in loaded_map.items()}
        except Exception:
            logging.info(f"[초기값 사용] load_worker_v_data: 파일 없음 또는 형식 오류.")
            self.worker_v_map = {}

    def save_schedule_to_excel(self):
        if self.current_schedule_df.empty:
            messagebox.showwarning("경고", "저장할 근무표 데이터가 없습니다.")
            return

        year, month = self.year_var.get(), self.month_var.get()
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"근무표_{year}년_{month}월.xlsx"
        )
        if filename:
            try:
                with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                    schedule_df_to_save = self.current_schedule_df.copy()
                    schedule_df_to_save.columns = [col.split('(')[0] for col in schedule_df_to_save.columns]
                    schedule_df_to_save.to_excel(writer, sheet_name='근무표', index=True, header=True)
                    
                    if not self.current_summary_df.empty:
                        self.current_summary_df.to_excel(writer, sheet_name='통계', index=False)
                        
                messagebox.showinfo("저장 완료", f"'{filename}'으로 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다: {e}")

    # ----------------------------------------------------------------------
    # [근무자 UI 관리]
    # ----------------------------------------------------------------------

    def update_gui_after_worker_change(self):
        self.save_worker_names()
        self.load_worker_categories() 
        self.display_initial_schedule_table()
        
    def worker_management_dialog(self):
        dialog = tk.Toplevel(self.root); dialog.title("근무자 명단 및 직책 관리"); dialog.geometry("400x500")
        dialog.transient(self.root); dialog.grab_set()

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f'+{x}+{y}')

        main_frame = ttk.Frame(dialog); main_frame.pack(padx=10, pady=10, fill='both', expand=True)

        list_frame = ttk.Frame(main_frame); list_frame.pack(fill='x', pady=(0, 10))
        tk.Label(list_frame, text="근무자 명단 (더블클릭/드래그로 순서 변경)", font=('Malgun Gothic', 10, 'bold')).pack(anchor='w')

        tree_frame = ttk.Frame(main_frame); tree_frame.pack(fill='both', expand=True)
        worker_tree = ttk.Treeview(tree_frame, columns=('Name', 'Category'), show='headings', selectmode='browse')
        worker_tree.heading('Name', text='이름'); worker_tree.column('Name', anchor='center', width=100)
        worker_tree.heading('Category', text='직책'); worker_tree.column('Category', anchor='center', width=100)
        
        def refresh_worker_tree(workers):
            worker_tree.delete(*worker_tree.get_children())
            for name in workers:
                category = self.worker_categories_map.get(name, DEFAULT_CATEGORY)
                worker_tree.insert('', 'end', values=(name, category), tags=(name, category))

        refresh_worker_tree(self.worker_names)
        worker_tree.pack(side='left', fill='both', expand=True)
        
        def start_category_edit(event):
            try:
                item_id = worker_tree.identify_row(event.y)
                column_id = worker_tree.identify_column(event.x)
                if not item_id or column_id != '#2': return

                worker_name = worker_tree.item(item_id, 'values')[0]
                
                bbox = worker_tree.bbox(item_id, column_id)
                if not bbox: return
                x, y, width, height = bbox
                
                category_var = tk.StringVar(value=worker_tree.item(item_id, 'values')[1])
                combo = ttk.Combobox(worker_tree, textvariable=category_var, values=WORKER_CATEGORIES, width=width, font=('Malgun Gothic', 10))
                combo.place(x=x, y=y, width=width, height=height)

                def update_category(e=None):
                    new_category = category_var.get()
                    if new_category in WORKER_CATEGORIES:
                        worker_tree.set(item_id, 'Category', new_category)
                        self.worker_categories_map[worker_name] = new_category
                        self.save_worker_categories()
                        combo.destroy()
                    else:
                        messagebox.showwarning("입력 오류", "유효한 직책을 선택해 주세요.", parent=dialog)

                combo.bind("<<ComboboxSelected>>", update_category)
                combo.bind("<FocusOut>", update_category)
                combo.focus_set()
                
            except Exception as e: logging.error(f"[start_category_edit] {e}")

        worker_tree.bind("<Double-1>", start_category_edit)
        
        def add_worker():
            name = simpledialog.askstring("새 근무자 추가", "추가할 근무자의 이름을 입력하세요:", parent=dialog)
            if name and name not in self.worker_names:
                self.worker_names.append(name)
                self.worker_categories_map[name] = DEFAULT_CATEGORY
                self.update_gui_after_worker_change()
                refresh_worker_tree(self.worker_names)

        def remove_worker():
            selected_items = worker_tree.selection()
            if not selected_items:
                messagebox.showwarning("경고", "삭제할 근무자를 선택하세요.", parent=dialog)
                return
            
            selected_names = [worker_tree.item(item, 'values')[0] for item in selected_items]
            if messagebox.askyesno("삭제 확인", f"선택한 근무자({', '.join(selected_names)})를 명단에서 삭제하시겠습니까?", parent=dialog):
                for name in selected_names:
                    if name in self.worker_names: self.worker_names.remove(name)
                    if name in self.worker_categories_map: del self.worker_categories_map[name]
                
                self.update_gui_after_worker_change()
                refresh_worker_tree(self.worker_names)

        button_frame = ttk.Frame(main_frame); button_frame.pack(fill='x', pady=5)
        ttk.Button(button_frame, text="추가", command=add_worker).pack(side='left', expand=True, fill='x', padx=2)
        ttk.Button(button_frame, text="삭제", command=remove_worker).pack(side='left', expand=True, fill='x', padx=2)
        
        def start_drag(event):
            selected_item = worker_tree.focus()
            if selected_item:
                worker_tree.drag_item = selected_item
                worker_tree.drag_name = worker_tree.item(selected_item, 'values')[0]
                worker_tree.config(cursor="hand2")

        def drag_motion(event):
            if hasattr(worker_tree, 'drag_item'):
                worker_tree.y_pos = event.y
        
        def drop_item(event):
            if not hasattr(worker_tree, 'drag_item'): return
            
            target_item = worker_tree.identify_row(event.y)
            drag_name = worker_tree.drag_name
            
            old_index = self.worker_names.index(drag_name)
            
            if target_item:
                target_name = worker_tree.item(target_item, 'values')[0]
                new_index = self.worker_names.index(target_name)
                
                self.worker_names.pop(old_index)
                
                if new_index > old_index:
                    self.worker_names.insert(new_index, drag_name)
                else: 
                    self.worker_names.insert(new_index, drag_name)
                    
            else: 
                self.worker_names.pop(old_index)
                self.worker_names.append(drag_name)

            refresh_worker_tree(self.worker_names)
            self.update_gui_after_worker_change() 
            worker_tree.config(cursor="")
            del worker_tree.drag_item
            
        worker_tree.bind("<ButtonPress-1>", start_drag); worker_tree.bind("<B1-Motion>", drag_motion); worker_tree.bind("<ButtonRelease-1>", drop_item)
        
        ttk.Button(dialog, text="닫기", command=dialog.destroy, style='Dialog.Secondary.TButton').pack(pady=10)
        self.root.wait_window(dialog)


    # ----------------------------------------------------------------------
    # [연차 관리 UI]
    # ----------------------------------------------------------------------

    def manage_worker_v_dialog(self):
        """근무자별 연차 초기 일수를 관리하는 다이얼로그"""
        current_year = self.year_var.get()
        
        dialog = tk.Toplevel(self.root); dialog.title(f"{current_year}년 연차 초기 일수 관리"); dialog.geometry("450x400")
        dialog.transient(self.root); dialog.grab_set()
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f'+{x}+{y}')
        
        tk.Label(dialog, text=f"**{current_year}년** 근무자별 연차 초기 일수 (V)", font=('Malgun Gothic', 12, 'bold')).pack(pady=10)
        
        main_frame = ttk.Frame(dialog); main_frame.pack(padx=10, pady=5, fill='both', expand=True)
        
        v_tree = ttk.Treeview(main_frame, columns=['Name', 'V_Days'], show='headings', selectmode='browse')
        v_tree.heading('Name', text='근무자 이름')
        v_tree.column('Name', anchor='center', width=150, stretch=tk.NO)
        v_tree.heading('V_Days', text='초기 연차 일수')
        v_tree.column('V_Days', anchor='center', width=150, stretch=tk.NO)

        def refresh_v_tree(tree, workers):
            tree.delete(*tree.get_children())
            year_data = self.worker_v_map.get(current_year, {}) 
            for name in workers: 
                v_days = year_data.get(name, 21.5) 
                tree.insert('', 'end', values=(name, v_days), tags=(name,))

        refresh_v_tree(v_tree, self.worker_names)
        v_tree.pack(fill='both', expand=True)
        
        def start_v_edit(event):
            try:
                if hasattr(dialog, 'editor_widget') and dialog.editor_widget.winfo_exists():
                    dialog.editor_widget.destroy()

                region = v_tree.identify_region(event.x, event.y)
                if region != "cell": return

                column_id = v_tree.identify_column(event.x)
                item_id = v_tree.identify_row(event.y)
                
                if column_id != '#2': return 

                worker_name = v_tree.item(item_id, 'values')[0] 
                
                bbox = v_tree.bbox(item_id, column_id)
                if not bbox: return
                x, y, width, height = bbox

                current_v = float(v_tree.item(item_id, 'values')[1])
                
                v_var = tk.DoubleVar(value=current_v)
                spinbox = ttk.Spinbox(v_tree, from_=0.0, to=30.0, increment=0.5, textvariable=v_var, width=width, font=('Malgun Gothic', 10))
                spinbox.place(x=x, y=y, width=width, height=height)

                def update_v(e):
                    try:
                        new_v = round(float(v_var.get()), 2)
                        
                        v_tree.set(item_id, 'V_Days', new_v) 
                        
                        year_data = self.worker_v_map.get(current_year, {})
                        year_data[worker_name] = new_v
                        self.worker_v_map[current_year] = year_data 
                        
                        self.save_worker_v_data() 
                        spinbox.destroy()
                        
                        self.current_summary_df = self.generate_schedule_summary(self.current_schedule_df, self.year_var.get(), self.month_var.get())
                        self.display_summary_table(self.current_summary_df) 
                        
                    except ValueError:
                        messagebox.showwarning("입력 오류", "유효한 숫자를 입력해 주세요.", parent=dialog)
                        spinbox.destroy()
                    except Exception as err:
                        logging.error(f"[update_v] {err}")
                        spinbox.destroy()

                spinbox.bind("<Return>", update_v)
                spinbox.bind("<FocusOut>", update_v)
                spinbox.focus_set();
                dialog.editor_widget = spinbox
                
            except Exception as e: 
                logging.error(f"[start_v_edit] {e}")

        v_tree.bind("<Double-1>", start_v_edit)
        
        button_frame_bottom = ttk.Frame(dialog); button_frame_bottom.pack(side='bottom', pady=10)
        ttk.Button(button_frame_bottom, text="닫기", command=dialog.destroy, style='Dialog.Secondary.TButton').pack(side='left', padx=10)
        
        self.root.wait_window(dialog)


    # ----------------------------------------------------------------------
    # [근무표/통계 및 UI 표시]
    # ----------------------------------------------------------------------

    def _get_previous_duty(self, worker_name, year, month):
        """이전 달 마지막 날의 근무 정보를 가져옴 (없는 경우 'Off')"""
        prev_month_key = f"{year}-{month:02d}"
        return self.prev_month_last_day_duties.get(prev_month_key, {}).get(worker_name, 'Off')

    def get_month_days(self, year, month):
        """해당 월의 날짜와 요일을 튜플 리스트로 반환"""
        num_days = calendar.monthrange(year, month)[1]
        days = []
        for day in range(1, num_days + 1):
            date = datetime.date(year, month, day)
            day_name = date.strftime('%a').replace('Sun', '일').replace('Mon', '월').replace('Tue', '화').replace('Wed', '수').replace('Thu', '목').replace('Fri', '금').replace('Sat', '토')
            days.append(f"{day} ({day_name})")
        return days

    def display_schedule_table(self, df_schedule):
        for widget in self.schedule_frame.winfo_children(): widget.destroy()

        if df_schedule.empty:
            tk.Label(self.schedule_frame, text="날짜를 선택하고 근무표를 생성하세요.", font=('Malgun Gothic', 14), fg='#666666', bg='white').pack(pady=100); return

        ttk.Style().configure("Schedule.Treeview.Heading", background="#E8F0FE", foreground="#333333", font=('Malgun Gothic', 9, 'bold'))
        tree_frame = ttk.Frame(self.schedule_frame); tree_frame.pack(fill='both', expand=True)

        xscrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        yscrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        
        columns = df_schedule.columns.tolist()
        self.current_tree = ttk.Treeview(
            tree_frame, 
            columns=['근무자'] + columns, 
            show='headings', 
            xscrollcommand=xscrollbar.set, 
            yscrollcommand=yscrollbar.set,
            style="Schedule.Treeview"
        )
        xscrollbar.config(command=self.current_tree.xview)
        yscrollbar.config(command=self.current_tree.yview)

        self.current_tree.heading('근무자', text='근무자'); self.current_tree.column('근무자', width=80, anchor='center', stretch=tk.NO)
        
        for i, col in enumerate(columns):
            self.current_tree.heading(col, text=col, anchor='center')
            self.current_tree.column(col, width=50, anchor='center', stretch=tk.NO)
            
            if '(토)' in col or '(일)' in col:
                self.current_tree.tag_configure(f'day_{i}', background='#F0F8FF') 
            else:
                 self.current_tree.tag_configure(f'day_{i}', background='white')

        for worker, row in df_schedule.iterrows():
            values = [worker] + row.tolist()
            tags = [f'day_{i}' for i in range(len(columns))]
            self.current_tree.insert('', 'end', values=values, tags=tuple(tags))

        xscrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        yscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.current_tree.pack(side=tk.LEFT, fill='both', expand=True)
        
        self.current_tree.bind("<Double-1>", self.start_schedule_edit)
        self.current_summary_df = self.generate_schedule_summary(df_schedule, self.year_var.get(), self.month_var.get())
        self.display_summary_table(self.current_summary_df)

    def display_initial_schedule_table(self):
        """새 근무표 생성 전에 초기 빈 테이블을 보여줌"""
        year = self.year_var.get(); month = self.month_var.get()
        days = self.get_month_days(year, month)
        
        df_schedule = pd.DataFrame('', index=self.worker_names, columns=days)
        self.current_schedule_df = df_schedule.copy()
        
        # 1일의 이전 근무 상태 표시 (DH 모드일 경우)
        if self.is_head_nurse_mode.get():
            first_day_col = days[0]
            for worker in self.worker_names:
                if self.worker_categories_map.get(worker) == '수선생님':
                    prev_duty = self._get_previous_duty(worker, year, month)
                    if prev_duty in ['E', 'N']:
                        df_schedule.loc[worker, first_day_col] = 'O'
        
        self.display_schedule_table(df_schedule)

    def go_to_current_month(self):
        self.year_var.set(CURRENT_YEAR); self.month_var.set(CURRENT_MONTH)

    def load_and_display_data_after_startup(self):
        self.load_all_schedules()
        
        year = self.year_var.get(); month = self.month_var.get()
        if not self.load_schedule_from_memory(year, month):
            self.display_initial_schedule_table()
        
        self.update_month_label()

    def display_summary_table(self, summary_df):
        for widget in self.summary_frame.winfo_children(): widget.destroy()
        if summary_df.empty:
            tk.Label(self.summary_frame, text="근무표 생성 후\n통계가 표시됩니다.", font=('Malgun Gothic', 12), bg='white').pack(pady=100, padx=50); return
            
        ttk.Style().configure("Summary.Treeview.Heading", background="#E8F0FE", foreground="#333333", font=('Malgun Gothic', 9, 'bold'))
        tk.Label(self.summary_frame, text="근무 합산 통계", font=('Malgun Gothic', 12, 'bold'), bg='white').pack(pady=(0, 5))
        tree_frame = ttk.Frame(self.summary_frame); tree_frame.pack(fill='both', expand=True)
        columns = list(summary_df.columns)
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style="Summary.Treeview")
        
        column_widths = {'근무자': 80, '초기 연차': 60, '잔여 연차': 60, '총 근무': 60, 'D': 40, 'E': 40, 'MD': 40, 'N': 40, 'DH': 40, 'Off': 40, 'V': 40, 'v.25': 40, 'v.0.5': 40, '주말_근무': 60} 
        
        for col in columns:
            tree.heading(col, text=col.replace('_', ' '), anchor='center')
            tree.column(col, width=column_widths.get(col, 50), anchor='center', stretch=tk.NO)
            
        for index, row in summary_df.iterrows(): tree.insert('', 'end', values=row.tolist())
        tree.pack(fill='both', expand=True)

    def generate_schedule_summary(self, df_schedule, year, month):
        """근무표 데이터에서 유형별 집계 통계 생성 및 연차 차감 로직 추가"""
        if df_schedule.empty: return pd.DataFrame()

        DUTY_TYPES = ['D', 'E', 'N', 'MD', 'DH', 'Off', 'V', 'v.25', 'v.0.5']
        
        summary_df = df_schedule.stack().groupby(level=0).value_counts().unstack(fill_value=0)
        
        summary_df['Off'] = summary_df.get('O', 0) + summary_df.get('Off', 0)
        summary_df = summary_df.drop(columns=['O'], errors='ignore')

        for col in DUTY_TYPES:
            if col not in summary_df.columns: summary_df[col] = 0

        summary_df = summary_df.reset_index(names=['근무자'])

        weekend_cols = [col for col in df_schedule.columns if '(토)' in col or '(일)' in col]
        summary_df['주말_근무'] = df_schedule[weekend_cols].apply(lambda row: row.astype(str).str.contains('|'.join(WORK_DUTIES)).sum(), axis=1).values

        summary_df['총 근무'] = summary_df[['D', 'E', 'MD', 'N', 'DH']].sum(axis=1)
        
        current_year_v_map = self.worker_v_map.get(year, {name: 21.5 for name in self.worker_names})
        
        initial_v_days = []
        remaining_v_days = []
        
        for index, row in summary_df.iterrows():
            worker_name = row['근무자']
            initial_v = current_year_v_map.get(worker_name, 21.5) 
            
            used_v = (row.get('V', 0) * 1.0) + (row.get('v.0.5', 0) * 0.5) + (row.get('v.25', 0) * 0.25)
            
            initial_v_days.append(initial_v)
            remaining_v_days.append(round(initial_v - used_v, 2))
            
        summary_df['초기 연차'] = initial_v_days
        summary_df['잔여 연차'] = remaining_v_days
        
        final_cols = ['근무자', '초기 연차', '잔여 연차', '총 근무', 'D', 'E', 'DH', 'MD', 'N', 'Off', 'V', 'v.25', 'v.0.5', '주말_근무']
        return summary_df.reindex(columns=final_cols)


    def update_schedule_cell(self, worker_name, day_col, new_duty, is_manual=True):
        if self.current_schedule_df.empty: return

        if worker_name in self.current_schedule_df.index and day_col in self.current_schedule_df.columns:
            old_duty = self.current_schedule_df.loc[worker_name, day_col]
            self.current_schedule_df.loc[worker_name, day_col] = new_duty

            self.current_summary_df = self.generate_schedule_summary(self.current_schedule_df, self.year_var.get(), self.month_var.get())
            self.display_summary_table(self.current_summary_df)

            if self.current_tree:
                for item_id in self.current_tree.get_children():
                    values = list(self.current_tree.item(item_id, 'values'))
                    if values[0] == worker_name:
                        try:
                            col_index = self.current_schedule_df.columns.get_loc(day_col) + 1 
                            values[col_index] = new_duty
                            self.current_tree.item(item_id, values=values)
                            break
                        except Exception as e:
                            logging.error(f"UI update failed: {e}")
                            
            if is_manual:
                self.manual_edited_cells.add((worker_name, day_col))

    def start_schedule_edit(self, event):
        if not self.current_tree: return

        try:
            item_id = self.current_tree.identify_row(event.y)
            column_id = self.current_tree.identify_column(event.x)
            
            if not item_id or column_id == '#1': return 

            col_index = int(column_id.replace('#', '')) - 1
            day_col = self.current_schedule_df.columns[col_index - 1] 
            worker_name = self.current_tree.item(item_id, 'values')[0] 
            current_duty = self.current_schedule_df.loc[worker_name, day_col]
            
            bbox = self.current_tree.bbox(item_id, column_id)
            if not bbox: return
            x, y, width, height = bbox

            duty_var = tk.StringVar(value=current_duty)
            combo = ttk.Combobox(
                self.current_tree, 
                textvariable=duty_var, 
                values=EDITABLE_SHIFTS, 
                width=width, 
                font=('Malgun Gothic', 10),
                justify='center'
            )
            combo.place(x=x, y=y, width=width, height=height)

            def update_duty(e=None):
                new_duty = duty_var.get()
                if new_duty in EDITABLE_SHIFTS:
                    self.update_schedule_cell(worker_name, day_col, new_duty if new_duty else '')
                    combo.destroy()
                else:
                    messagebox.showwarning("입력 오류", "유효한 근무를 선택하거나 입력해 주세요.")
                    combo.focus_set()

            combo.bind("<<ComboboxSelected>>", update_duty)
            combo.bind("<Return>", update_duty)
            combo.bind("<FocusOut>", update_duty) 
            combo.focus_set()
            
        except Exception as e:
            logging.error(f"Edit error: {e}")

    def generate_monthly_schedule(self, df_schedule):
        """근무표 자동 생성 로직 (빈 칸에 자동 할당 로직 추가)"""
        
        df = df_schedule.copy()
        workers = df.index.tolist()
        days = df.columns.tolist()
        
        # 1-2. 근무표 초기화 및 수동 편집된 셀 보호 (생략 - 기존 상태 유지)

        # 3. 수동 편집된 셀 복구 및 PRESERVED_SHIFTS 유지
        for worker in workers:
            for day in days:
                duty = df.loc[worker, day]
                
                # 수선생님(1순위) 주간 근무 모드 처리
                if self.is_head_nurse_mode.get() and worker == workers[0] and self.worker_categories_map.get(worker) == '수선생님':
                    if '(토)' in day or '(일)' in day:
                        # 주말은 Off로 자동 할당 (빈 칸일 경우에만)
                        if df.loc[worker, day] == '': 
                            df.loc[worker, day] = 'Off'
                    elif df.loc[worker, day] == '':
                         df.loc[worker, day] = 'D' # 주중은 D로 자동 할당 (빈 칸일 경우에만)


        # 4. [⭐ FIX: 자동 할당 로직 재도입 - 수동 입력/특수 근무 외 나머지 칸에 근무 할당]
        # 간단한 근무 순환표 정의 (Off가 많은 초기 상태에서 D, E, N을 순환하며 할당)
        duty_cycle = ['Off', 'Off', 'D', 'E', 'N'] 
        cycle_index = 0
        
        for worker in workers:
            # 1순위 수선생님은 이미 위에서 처리했으므로 건너뛰기
            is_head_nurse = self.worker_categories_map.get(worker) == '수선생님' and self.is_head_nurse_mode.get()
            if is_head_nurse and worker == workers[0]: 
                continue 
                
            for day in days:
                current_duty = df.loc[worker, day]
                
                # 현재 칸이 비어있는 경우에만 순환 할당
                if current_duty == '':
                    df.loc[worker, day] = duty_cycle[cycle_index % len(duty_cycle)]
                    cycle_index += 1
                # 이미 수동 입력('D', 'E', 'V', 'Off' 등)이 된 칸은 유지됩니다.

        # 5. 마지막 날 근무 저장 (다음 달을 위한 데이터)
        last_day = days[-1]
        next_month_duties = {}
        for worker in workers:
            next_month_duties[worker] = df.loc[worker, last_day]
        self.prev_month_last_day_duties[f"{self.year_var.get()}-{self.month_var.get()+1:02d}"] = next_month_duties
        self.save_prev_month_schedule()
        
        return df

    def generate_and_display(self):
        if not self.worker_names:
            messagebox.showwarning("경고", "근무자 명단이 비어있습니다. '근무자 관리'에서 명단을 추가해주세요.")
            return

        if self.current_schedule_df.empty:
            self.display_initial_schedule_table()
        
        try:
            generated_df = self.generate_monthly_schedule(self.current_schedule_df)
            self.current_schedule_df = generated_df
            self.save_current_schedule_to_memory() 
            self.display_schedule_table(generated_df)
            self.save_all_schedules() 
            
            messagebox.showinfo("완료", f"{self.year_var.get()}년 {self.month_var.get()}월 근무표가 생성되었습니다. 수동 입력하신 근무는 유지됩니다.")
        except Exception as e:
            logging.error(f"근무표 생성 중 오류 발생: {e}", exc_info=True)
            messagebox.showerror("오류", f"근무표 생성 중 오류가 발생했습니다: {e}")

    def clear_schedule(self):
        year, month = self.year_var.get(), self.month_var.get()
        if messagebox.askyesno("초기화 확인", f"{year}년 {month}월 근무표를 초기화하시겠습니까? (수동 편집 내용 포함)"):
            self.manual_edited_cells.clear()
            self.current_schedule_df = pd.DataFrame()
            if year in self.monthly_schedules and month in self.monthly_schedules[year]:
                del self.monthly_schedules[year][month]
            self.save_all_schedules()
            self.display_initial_schedule_table()
            self.display_summary_table(pd.DataFrame())
            messagebox.showinfo("완료", "근무표가 초기화되었습니다.")

    # ----------------------------------------------------------------------
    # [메인 UI 및 이벤트 연결]
    # ----------------------------------------------------------------------

    def show_popup_menu(self, menu_name, parent_button):
        
        menu = tk.Menu(
            self.root, 
            tearoff=0, 
            bg='white', 
            fg='#333333', 
            activebackground='#F0F0F0', 
            activeforeground=TOSS_BLUE,  
            relief='flat',               
            borderwidth=0,               
            font=('Malgun Gothic', 10)
        )
        
        if menu_name == '파일':
            menu.add_command(label="종료", command=self.on_closing)
            
        elif menu_name == '근무자 관리':
            menu.add_command(label="명단 관리 (직책 포함)", command=self.worker_management_dialog)
            menu.add_separator()
            
            def toggle_head_nurse_mode_command():
                if not self.worker_names:
                    messagebox.showwarning("경고", "근무자 명단이 비어있습니다.")
                    self.is_head_nurse_mode.set(False)
                    return
                hn_name = self.worker_names[0]
                hn_category = self.worker_categories_map.get(hn_name)
                
                if self.is_head_nurse_mode.get() and hn_category != '수선생님':
                    messagebox.showwarning("경고", f"현재 1순위 근무자({hn_name})의 직책이 '수선생님'이 아닙니다. 직책을 먼저 설정해주세요.")
                    self.is_head_nurse_mode.set(False)
                    return
                    
                self.display_initial_schedule_table()
                messagebox.showinfo("모드 변경", f"상위 근무자({hn_name}) 주간 근무 모드가 {'적용' if self.is_head_nurse_mode.get() else '해제'}되었습니다.")
                
            menu.add_checkbutton(label="수선생님(1순위) 주간 근무 모드", onvalue=True, offvalue=False, variable=self.is_head_nurse_mode, command=toggle_head_nurse_mode_command)

        elif menu_name == '데이터':
            menu.add_command(label="데이터 저장 (.xlsx)", command=self.save_schedule_to_excel)
            menu.add_separator()
            menu.add_command(label="연차 초기 일수 관리", command=self.manage_worker_v_dialog)


        parent_button.update_idletasks() 
        x = parent_button.winfo_rootx()
        y = parent_button.winfo_rooty() + parent_button.winfo_height() + 1 

        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def update_month_label(self):
        self.month_label_text.set(f"{self.year_var.get()}년 {self.month_var.get()}월")

    def setup_main_window(self):
        self.root.title("근무표 관리 시스템")
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.configure(bg='white')
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        style = ttk.Style(self.root)
        style.theme_create("TossTheme", parent="alt", settings={
            "TFrame": {"configure": {"background": "white"}},
            "TButton": {"configure": {"background": "#F2F4F6", "foreground": "#333333", "font": ('Malgun Gothic', 10), "relief": "flat", "padding": [10, 5]}},
            "Primary.TButton": {"configure": {"background": TOSS_BLUE, "foreground": "white", "font": ('Malgun Gothic', 10, 'bold')}},
            "Clear.TButton": {"configure": {"background": "#F9E4E4", "foreground": "#FF0000"}},
            "Dialog.Secondary.TButton": {"configure": {"background": "#F2F4F6", "foreground": "#333333"}},
            "TLabel": {"configure": {"background": "white", "foreground": "#333333", "font": ('Malgun Gothic', 10)}},
            "TCombobox": {"configure": {"fieldbackground": "white", "selectbackground": "white", "selectforeground": "#333333", "font": ('Malgun Gothic', 10)}},
            "TNotebook": {"configure": {"background": "white"}, "TNotebook.Tab": {"configure": {"padding": [15, 5]}}},
            "Treeview": {"configure": {"rowheight": 25, "background": "white", "fieldbackground": "white", "font": ('Malgun Gothic', 10)}},
            "Treeview.Heading": {"configure": {"font": ('Malgun Gothic', 10, 'bold'), "background": "#E8F0FE", "foreground": "#333333", "relief": "flat"}},
        })
        style.theme_use("TossTheme")

        menu_frame = ttk.Frame(self.root, style='Toss.TFrame'); menu_frame.pack(fill='x', padx=10, pady=(10, 5))
        
        btn_file = ttk.Button(menu_frame, text="파일", command=lambda: self.show_popup_menu('파일', btn_file), style='TButton')
        btn_file.pack(side='left', padx=(0, 5))
        
        btn_worker = ttk.Button(menu_frame, text="근무자 관리", command=lambda: self.show_popup_menu('근무자 관리', btn_worker), style='TButton')
        btn_worker.pack(side='left', padx=5)

        btn_data = ttk.Button(menu_frame, text="데이터", command=lambda: self.show_popup_menu('데이터', btn_data), style='TButton')
        btn_data.pack(side='left', padx=5)

        date_frame = ttk.Frame(self.root, style='Toss.TFrame'); date_frame.pack(pady=5)
        
        ttk.Button(date_frame, text="◀", command=lambda: self.adjust_date(-1, 'month'), width=3).pack(side='left', padx=5)
        ttk.Button(date_frame, text="◀◀", command=lambda: self.adjust_date(-1, 'year'), width=4).pack(side='left')

        self.month_label = tk.Label(date_frame, textvariable=self.month_label_text, font=('Malgun Gothic', 14, 'bold'), bg='white', fg=TOSS_BLUE)
        self.month_label.pack(side='left', padx=15)
        
        ttk.Button(date_frame, text="▶▶", command=lambda: self.adjust_date(1, 'year'), width=4).pack(side='left')
        ttk.Button(date_frame, text="▶", command=lambda: self.adjust_date(1, 'month'), width=3).pack(side='left', padx=5)
        
        ttk.Button(date_frame, text="오늘", command=self.go_to_current_month, style='Dialog.Secondary.TButton').pack(side='left', padx=(20, 0))

        main_content_frame = ttk.Frame(self.root, style='Toss.TFrame'); main_content_frame.pack(fill='both', expand=True, padx=10)
        
        self.schedule_frame = ttk.Frame(main_content_frame, relief='flat', borderwidth=0, padding=10, style='Toss.TFrame')
        self.schedule_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        self.summary_frame = ttk.Frame(main_content_frame, relief='flat', borderwidth=0, padding=10, style='Toss.TFrame')
        self.summary_frame.pack(side='left', fill='y', padx=(10, 0))

        button_container = ttk.Frame(self.root, style='Toss.TFrame'); button_container.pack(pady=10)
        
        ttk.Button(button_container, text="✨ 근무표 생성", command=self.generate_and_display, style='Primary.TButton').pack(side='left', padx=10)
        ttk.Button(button_container, text="🗑️ 근무표 초기화", command=self.clear_schedule, style='Clear.TButton').pack(side='left', padx=10)

        def on_date_change_cb(*args):
            self.save_current_schedule_to_memory()
            self.update_month_label()
            year, month = self.year_var.get(), self.month_var.get()
            if not self.load_schedule_from_memory(year, month):
                self.display_initial_schedule_table()
        
        self.year_var.trace_add("write", on_date_change_cb)
        self.month_var.trace_add("write", on_date_change_cb)

        self.root.after(100, self.load_and_display_data_after_startup)

        footer_frame = ttk.Frame(self.root, style='Toss.TFrame'); footer_frame.pack(side='bottom', fill='x', padx=10, pady=(0, 5))
        tk.Label(footer_frame, text="made by TKㅣver.241126 V-Patch (Final)", font=('Malgun Gothic', 9), fg='#AAAAAA', bg='white').pack(side='right', padx=10)

    def adjust_date(self, delta, unit):
        current_year = self.year_var.get()
        current_month = self.month_var.get()
        
        new_year = current_year
        new_month = current_month

        if unit == 'month':
            new_month += delta
            if new_month > 12: new_month = 1; new_year += 1
            elif new_month < 1: new_month = 12; new_year -= 1
        elif unit == 'year':
            new_year += delta

        self.year_var.set(new_year)
        self.month_var.set(new_month)

    def on_closing(self):
        try:
            self.save_all_schedules()
            self.save_worker_v_data()
            self.save_prev_month_schedule()
        except Exception as e:
            logging.error(f"종료 중 데이터 저장 오류: {e}")
        self.root.destroy()

# ==============================================================================
# 4. 실행 진입점
# ==============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()