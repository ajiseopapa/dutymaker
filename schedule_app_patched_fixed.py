import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import datetime
import pandas as pd
import calendar
import json
import random
import math
import logging

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

        # [데이터 로드]
        self.load_worker_names()
        self.load_prev_month_schedule() 
        self.load_worker_categories() 

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

    def load_all_schedules(self):
        """저장된 모든 월별 스케줄 데이터를 JSON 파일에서 불러옴"""
        try:
            with open(MONTHLY_SCHEDULES_FILE, 'r', encoding='utf-8') as f:
                self.monthly_schedules = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.monthly_schedules = {}
            logging.info(f"[{MONTHLY_SCHEDULES_FILE}] 파일이 없거나 형식 오류. 새 딕셔너리 생성.")
        except Exception as e:
            logging.error(f"load_all_schedules: {e}")

    def save_all_schedules(self):
        """모든 월별 스케줄 데이터를 JSON 파일에 저장"""
        try:
            with open(MONTHLY_SCHEDULES_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.monthly_schedules, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.error(f"save_all_schedules: {e}")

    def save_current_schedule_to_memory(self, df_schedule, year, month):
        """현재 근무표를 내부 딕셔너리에 저장하고 파일에 반영"""
        key = f"{year}-{month:02d}"
        
        data_to_save = {
            'columns': df_schedule.columns.tolist(),
            'index': df_schedule.index.tolist(),
            'data': df_schedule.values.tolist(),
            'manual_edits': list(self.manual_edited_cells)
        }
        self.monthly_schedules[key] = data_to_save
        self.save_all_schedules()
        
    def load_schedule_from_memory(self, year, month):
        """특정 월의 근무표를 내부 딕셔너리에서 불러옴 (DataFrame 및 수동 편집 목록 반환)"""
        key = f"{year}-{month:02d}"
        if key in self.monthly_schedules:
            data = self.monthly_schedules[key]
            df = pd.DataFrame(data['data'], index=data['index'], columns=data['columns'])
            
            manual_edits_list = data.get('manual_edits', [])
            manual_edits_set = set(tuple(item) for item in manual_edits_list)
            
            return df, manual_edits_set
        return None, set()

    def save_worker_names(self):
        """근무자 명단을 영구 저장"""
        try:
            with open(WORKER_LIST_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.worker_names, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.error(f"save_worker_names: {e}")

    def load_worker_names(self):
        """근무자 명단 불러오거나 기본값 사용"""
        try:
            with open(WORKER_LIST_FILE, 'r', encoding='utf-8') as f:
                loaded_names = json.load(f)
                if loaded_names and isinstance(loaded_names, list):
                    self.worker_names = loaded_names
                    return
        except Exception as e:
            logging.info(f"[초기값 사용] load_worker_names: {e}")
        self.worker_names = DEFAULT_WORKERS.copy()
        self.save_worker_names()

    def save_worker_categories(self):
        """근무자별 직책/구분 정보를 영구 저장"""
        try:
            with open(WORKER_CATEGORIES_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.worker_categories_map, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.error(f"save_worker_categories: {e}")

    def load_worker_categories(self):
        """근무자별 직책/구분 정보를 불러오거나 기본값 사용"""
        try:
            with open(WORKER_CATEGORIES_FILE, 'r', encoding='utf-8') as f:
                loaded_map = json.load(f)
                if loaded_map and isinstance(loaded_map, dict):
                    self.worker_categories_map = loaded_map
        except Exception:
            logging.info(f"[초기값 사용] load_worker_categories: 파일 없음.")
        
        updated_map = {}
        for name in self.worker_names:
            updated_map[name] = self.worker_categories_map.get(name, DEFAULT_CATEGORY)
            
        self.worker_categories_map = updated_map

    def save_prev_month_schedule(self):
        """현재 근무표의 마지막 5일 근무를 이전 달 데이터로 저장하여 다음 달 N-Set 연속성에 활용"""
        if self.current_schedule_df.empty: return
        try:
            df = self.current_schedule_df
            # 마지막 5일의 근무만 저장
            last_5_days_df = df.iloc[:, -5:] 
            
            last_day_duties_list = {
                worker: row.tolist() for worker, row in last_5_days_df.iterrows()
            }
            
            with open(PREV_MONTH_SCHEDULE_FILE, 'w', encoding='utf-8') as f:
                json.dump(last_day_duties_list, f, ensure_ascii=False, indent=4)
                
        except Exception as e:
            logging.error(f"save_prev_month_schedule: {e}")

    def load_prev_month_schedule(self):
        """이전 달의 마지막 5일 근무 기록을 불러옴"""
        try:
            with open(PREV_MONTH_SCHEDULE_FILE, 'r', encoding='utf-8') as f:
                self.prev_month_last_day_duties = json.load(f)
        except Exception as e:
            logging.info(f"load_prev_month_schedule: 이전 달 근무표 없음. {e}")
            self.prev_month_last_day_duties = {}

    def save_schedule_to_excel(self):
        """
        [🚨 FIX] 현재 근무표와 통계를 엑셀 파일로 저장하는 기능을 구현
        근무표와 통계를 각각 '근무표' 시트와 '근무_통계' 시트로 저장
        """
        if self.current_schedule_df.empty:
            messagebox.showwarning("경고", "저장할 근무표 데이터가 없습니다. 근무표를 먼저 생성해 주세요.")
            return

        year, month = self.year_var.get(), self.month_var.get()
        default_filename = f"{year}년_{month}월_근무표.xlsx"
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel files", "*.xlsx")],
            title="근무표 및 통계 저장"
        )
        
        if not file_path:
            return # 사용자가 취소함

        try:
            # pandas.ExcelWriter를 사용하여 여러 시트에 데이터를 저장
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 1. 근무표 시트 저장
                # 인덱스(근무자 이름)를 첫 번째 열로 포함
                df_schedule_reset = self.current_schedule_df.reset_index(names=['근무자'])
                df_schedule_reset.to_excel(writer, sheet_name='근무표', index=False)
                
                # 2. 통계 시트 저장 (있는 경우)
                if not self.current_summary_df.empty:
                    df_summary_reset = self.current_summary_df.copy()
                    df_summary_reset.to_excel(writer, sheet_name='근무_통계', index=False)

            messagebox.showinfo("저장 성공", f"근무표와 통계를 '{file_path}'에 성공적으로 저장했습니다.")
            logging.info(f"Excel saved to: {file_path}")

        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 파일 저장 중 오류가 발생했습니다.\n오류: {e}")
            logging.error(f"Error saving to Excel: {e}")

    # ----------------------------------------------------------------------
    # [근무자 UI 관리]
    # ----------------------------------------------------------------------

    def update_gui_after_worker_change(self):
        """근무자 추가/삭제/순서 변경시 UI 및 데이터 동기화 (과거 데이터 초기화 포함)"""
        self.monthly_schedules.clear()
        self.save_all_schedules()
        self.manual_edited_cells.clear()
        self.display_initial_schedule_table()
        self.save_worker_names()
        self.load_worker_categories() 
        self.save_worker_categories() 

    def worker_management_dialog(self):
        """근무자 추가/수정/삭제/순서 변경을 통합한 다이얼로그"""
        
        dialog = tk.Toplevel(self.root); dialog.title("근무자 명단 및 순서 관리"); dialog.geometry("700x500")
        dialog.transient(self.root); dialog.grab_set()
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f'+{x}+{y}')
        
        tk.Label(dialog, text="근무자 목록 (직책/구분 셀 더블 클릭하여 수정)", font=('Malgun Gothic', 14, 'bold')).pack(pady=10)
        
        main_frame = ttk.Frame(dialog); main_frame.pack(padx=10, pady=5, fill='both', expand=True)
        
        worker_tree = ttk.Treeview(main_frame, columns=['Name', 'Category'], show='headings', selectmode='browse')
        worker_tree.heading('Name', text='근무자 이름')
        worker_tree.column('Name', anchor='center', width=200, stretch=tk.NO)
        worker_tree.heading('Category', text='직책/구분')
        worker_tree.column('Category', anchor='center', width=150, stretch=tk.NO)

        def refresh_worker_tree(tree, workers):
            """Treeview를 현재 self.worker_names 리스트로 업데이트"""
            tree.delete(*tree.get_children())
            for name in workers: 
                category = self.worker_categories_map.get(name, DEFAULT_CATEGORY)
                tree.insert('', 'end', values=(name, category), tags=(name,))

        refresh_worker_tree(worker_tree, self.worker_names)
        worker_tree.pack(side='left', fill='both', expand=True)
        
        # 직책 수정 기능 추가 (더블 클릭 이벤트)
        def start_category_edit(event):
            try:
                if hasattr(dialog, 'editor_widget') and dialog.editor_widget.winfo_exists():
                    dialog.editor_widget.destroy()

                region = worker_tree.identify_region(event.x, event.y)
                if region != "cell": return

                column_id = worker_tree.identify_column(event.x)
                item_id = worker_tree.identify_row(event.y)
                
                if column_id != '#2': return 

                worker_name = worker_tree.item(item_id, 'values')[0] 
                
                bbox = worker_tree.bbox(item_id, column_id)
                if not bbox: return
                x, y, width, height = bbox

                combobox = ttk.Combobox(worker_tree, values=WORKER_CATEGORIES, width=width, font=('Malgun Gothic', 10), state='readonly')
                
                current_category = self.worker_categories_map.get(worker_name, DEFAULT_CATEGORY)
                combobox.set(current_category)
                
                combobox.place(x=x, y=y, width=width, height=height)

                def update_category(e):
                    new_category = combobox.get()
                    worker_tree.set(item_id, 'Category', new_category) 
                    self.worker_categories_map[worker_name] = new_category 
                    self.save_worker_categories() 
                    combobox.destroy()

                combobox.bind("<<ComboboxSelected>>", update_category)
                combobox.bind("<Return>", update_category)
                combobox.bind("<FocusOut>", lambda e: combobox.destroy() if e.widget == combobox else None)
                combobox.focus_set();
                dialog.editor_widget = combobox
                
            except Exception as e: 
                logging.error(f"[start_category_edit] {e}")

        worker_tree.bind("<Double-1>", start_category_edit)
        
        
        # 순서 변경 버튼 프레임
        reorder_button_frame = ttk.Frame(main_frame); reorder_button_frame.pack(side='right', padx=(5, 0))
        
        def move_worker(direction):
            try:
                selected_item = worker_tree.selection()[0]
                current_name = worker_tree.item(selected_item, 'values')[0]
                current_index = self.worker_names.index(current_name)
            except (IndexError, ValueError):
                messagebox.showwarning("경고", "먼저 순서를 바꿀 근무자를 선택해 주세요.", parent=dialog)
                return
            
            target_index = current_index + direction
            
            if 0 <= target_index < len(self.worker_names):
                self.worker_names.insert(target_index, self.worker_names.pop(current_index))
                refresh_worker_tree(worker_tree, self.worker_names)
                for item_id in worker_tree.get_children():
                    if worker_tree.item(item_id, 'values')[0] == current_name:
                        worker_tree.selection_set(item_id)
                        break
            else:
                messagebox.showwarning("경고", "더 이상 이동할 수 없습니다.", parent=dialog)

        ttk.Button(reorder_button_frame, text="🔼 위로 이동", command=lambda: move_worker(-1), style='Small.TButton').pack(pady=5, fill='x')
        ttk.Button(reorder_button_frame, text="🔽 아래로 이동", command=lambda: move_worker(1), style='Small.TButton').pack(pady=5, fill='x')

        # 추가/수정/삭제 버튼 프레임
        crud_frame = ttk.Frame(dialog); crud_frame.pack(pady=10)

        def _add_worker_in_dialog():
            new_name = simpledialog.askstring("근무자 추가", "추가할 근무자의 이름을 입력하세요:", parent=dialog)
            if new_name and new_name.strip():
                fixed_name = new_name.strip()
                if fixed_name not in self.worker_names:
                    self.worker_names.append(fixed_name)
                    self.worker_categories_map[fixed_name] = DEFAULT_CATEGORY
                    self.save_worker_categories()
                    refresh_worker_tree(worker_tree, self.worker_names)
                    messagebox.showinfo("성공", f"근무자 '{fixed_name}'이(가) 추가되었습니다.", parent=dialog)
                else:
                    messagebox.showwarning("중복", f"이미 존재하는 근무자입니다: '{fixed_name}'", parent=dialog)

        def _modify_worker_in_dialog():
            try:
                selected_item = worker_tree.selection()[0]
                old_name = worker_tree.item(selected_item, 'values')[0]
            except IndexError:
                messagebox.showwarning("경고", "먼저 수정할 근무자를 선택해 주세요.", parent=dialog); return
            
            new_name = simpledialog.askstring("새 이름 입력", f"'{old_name}'의 새로운 이름을 입력하세요:", parent=dialog)
            if new_name and new_name.strip() and new_name.strip() != old_name:
                fixed_new_name = new_name.strip()
                if fixed_new_name in self.worker_names:
                    messagebox.showwarning("중복", f"이미 존재하는 근무자 이름입니다: '{fixed_new_name}'", parent=dialog)
                    return
                
                self.worker_names[self.worker_names.index(old_name)] = fixed_new_name
                
                # 카테고리 맵의 키도 업데이트
                if old_name in self.worker_categories_map:
                    category = self.worker_categories_map.pop(old_name)
                    self.worker_categories_map[fixed_new_name] = category
                    self.save_worker_categories()
                    
                refresh_worker_tree(worker_tree, self.worker_names)
                for item_id in worker_tree.get_children():
                    if worker_tree.item(item_id, 'values')[0] == fixed_new_name:
                        worker_tree.selection_set(item_id)
                        break
                        
                messagebox.showinfo("성공", f"'{old_name}'이(가) '{fixed_new_name}'(으)로 수정되었습니다.", parent=dialog)
            
        def _delete_worker_in_dialog():
            try:
                selected_item = worker_tree.selection()[0]
                name_to_delete = worker_tree.item(selected_item, 'values')[0]
            except IndexError:
                messagebox.showwarning("경고", "먼저 삭제할 근무자를 선택해 주세요.", parent=dialog); return
            
            if messagebox.askyesno("삭제 확인", f"근무자 '{name_to_delete}'을(를) 정말 삭제하시겠습니까?", parent=dialog):
                self.worker_names.remove(name_to_delete)
                if name_to_delete in self.worker_categories_map:
                    del self.worker_categories_map[name_to_delete]
                    self.save_worker_categories()
                    
                refresh_worker_tree(worker_tree, self.worker_names)
                messagebox.showinfo("성공", f"근무자 '{name_to_delete}'이(가) 명단에서 삭제되었습니다.", parent=dialog)
        
        ttk.Button(crud_frame, text="➕ 근무자 추가", command=_add_worker_in_dialog, style='Small.TButton').pack(side='left', padx=5)
        ttk.Button(crud_frame, text="📝 선택 근무자 수정", command=_modify_worker_in_dialog, style='Small.TButton').pack(side='left', padx=5)
        ttk.Button(crud_frame, text="🗑️ 선택 근무자 삭제", command=_delete_worker_in_dialog, style='Small.TButton').pack(side='left', padx=5)

        # 하단 적용/닫기 버튼
        def on_ok():
            self.save_worker_categories() 
            self.update_gui_after_worker_change()
            dialog.destroy()
            messagebox.showinfo("적용 완료", "변경된 근무자 명단 및 순서가 메인 화면에 적용되었습니다.")
            
        button_frame_bottom = ttk.Frame(dialog); button_frame_bottom.pack(side='bottom', pady=10)
        ttk.Button(button_frame_bottom, text="확인 및 적용", command=on_ok, style='Dialog.Primary.TButton').pack(side='left', padx=10)
        ttk.Button(button_frame_bottom, text="닫기", command=dialog.destroy, style='Dialog.Secondary.TButton').pack(side='left', padx=10)
        
        self.root.wait_window(dialog)
        
    def start_worker_name_edit(self, event):
        """근무자 이름 Treeview에서 이름 더블 클릭시 수정 (고정 열)"""
        tree = self.current_tree
        try:
            region = tree.identify_region(event.x, event.y)
            if region != "heading": 
                column_id = tree.identify_column(event.x)
                item_id = tree.identify_row(event.y)
            else: return
            
            if column_id != '#1': return

            old_name = tree.item(item_id, 'values')[0]
            
            new_name = simpledialog.askstring(
                "근무자 이름 수정", 
                f"'{old_name}'의 새로운 이름을 입력하세요:", 
                parent=self.root
            )
            
            if new_name and new_name.strip() and new_name.strip() != old_name:
                fixed_new_name = new_name.strip()
                if fixed_new_name in self.worker_names:
                    messagebox.showwarning("중복", f"이미 존재하는 근무자 이름입니다: '{fixed_new_name}'")
                    return
                
                old_index = self.worker_names.index(old_name)
                self.worker_names[old_index] = fixed_new_name
                self.save_worker_names() 
                
                # worker_categories_map 업데이트
                if old_name in self.worker_categories_map:
                    category = self.worker_categories_map.pop(old_name)
                    self.worker_categories_map[fixed_new_name] = category
                    self.save_worker_categories() 
                
                # DataFrame Index 업데이트
                if not self.current_schedule_df.empty:
                    self.current_schedule_df = self.current_schedule_df.rename(index={old_name: fixed_new_name})
                    year, month = self.year_var.get(), self.month_var.get()
                    self.save_current_schedule_to_memory(self.current_schedule_df, year, month)
                
                # 이전 달 근무 기록 업데이트
                if old_name in self.prev_month_last_day_duties:
                    duty = self.prev_month_last_day_duties.pop(old_name)
                    self.prev_month_last_day_duties[fixed_new_name] = duty
                
                # UI 새로고침
                self.display_initial_schedule_table()
                messagebox.showinfo("성공", f"'{old_name}'이(가) '{fixed_new_name}'(으)로 수정되었습니다.")

        except IndexError:
            pass
        except Exception as e: 
            logging.error(f"[start_worker_name_edit] {e}")
            messagebox.showerror("오류", f"근무자 이름 수정 중 오류가 발생했습니다.\n오류: {e}")


    # ----------------------------------------------------------------------
    # [근무표/통계 및 UI 표시]
    # ----------------------------------------------------------------------
    
    def _get_previous_duty(self, worker_name, day_index, day_offset, schedule_data):
        """지정된 오프셋만큼 이전 근무를 가져옵니다. 전월 데이터도 확인합니다."""
        if day_index >= day_offset:
            # 당월 데이터에서 확인
            return schedule_data[worker_name][day_index - day_offset]
        
        # 전월 데이터에서 확인 (최대 5일)
        prev_duties = self.prev_month_last_day_duties.get(worker_name, [])
        prev_index = len(prev_duties) - day_offset + day_index
        
        if 0 <= prev_index < len(prev_duties):
            return prev_duties[prev_index]
            
        return ''

    def get_month_days(self, year, month):
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
        return year, month, last_day, day_columns 

    def display_schedule_table(self, df, year, month):
        """DataFrame을 Treeview로 표시하며, 헤더 스타일 변경"""
        for widget in self.schedule_frame.winfo_children(): widget.destroy()
        if df.empty:
            tk.Label(self.schedule_frame, text="근무표 데이터가 없습니다.", font=('Malgun Gothic', 14)).pack(pady=20); return
        
        ttk.Style().configure("Custom.Treeview.Heading", background="white", foreground="#A9A9A9", font=('Malgun Gothic', 10, 'bold'))
        
        tree_frame = ttk.Frame(self.schedule_frame); tree_frame.pack(fill='both', expand=True)
        
        tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        tree_scroll_y.pack(side='right', fill='y')
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.pack(side='bottom', fill='x')
        
        columns = ["근무자"] + list(df.columns)
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                            yscrollcommand=tree_scroll_y.set, 
                            xscrollcommand=tree_scroll_x.set)

        def remove_editor_widget(*args):
            if hasattr(tree, 'editor_widget') and tree.editor_widget.winfo_exists():
                tree.editor_widget.destroy()

        tree_scroll_y.config(command=lambda *args: (remove_editor_widget(), tree.yview(*args))) 
        tree_scroll_x.config(command=lambda *args: (remove_editor_widget(), tree.xview(*args)))
        
        tree.bind("<MouseWheel>", lambda e: (remove_editor_widget(), tree.yview_scroll(int(-1*(e.delta/120)), "units")))

        for col in df.columns:
            day_and_weekday = col.split('/', 1)[-1].strip()
            tree.heading(col, text=day_and_weekday, anchor='center'); tree.column(col, width=60, anchor='center', stretch=tk.NO)
        tree.heading("근무자", text="근무자", anchor='center'); tree.column("근무자", width=100, anchor='center', stretch=tk.NO)

        for worker, row in df.iterrows():
            tree.insert('', 'end', values=[worker] + row.tolist(), tags=(worker,))
        tree.pack(fill='both', expand=True)
        
        tree.bind("<Button-1>", self.start_schedule_edit) 
        tree.bind("<Double-1>", self.start_worker_name_edit) 

        self.month_label_text.set(f"🗓️ {year}년 {month}월 근무표")
        self.current_tree = tree 

    def display_summary_table(self, summary_df):
        for widget in self.summary_frame.winfo_children(): widget.destroy()
        if summary_df.empty:
            tk.Label(self.summary_frame, text="근무표 생성 후\n통계가 표시됩니다.", font=('Malgun Gothic', 12), bg='white').pack(pady=100, padx=50); return
        ttk.Style().configure("Summary.Treeview.Heading", background="#E8F0FE", foreground="#333333", font=('Malgun Gothic', 9, 'bold'))
        tk.Label(self.summary_frame, text="근무 합산 통계", font=('Malgun Gothic', 12, 'bold'), bg='white').pack(pady=(0, 5))
        tree_frame = ttk.Frame(self.summary_frame); tree_frame.pack(fill='both', expand=True)
        columns = list(summary_df.columns)
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style="Summary.Treeview")
        column_widths = {'근무자': 80, '전월 연차': 60, '총 연차': 60, '총 근무': 60, 'D': 40, 'E': 40, 'MD': 40, 'N': 40, 'DH': 40, 'Off': 40, 'V': 40, 'v.25': 40, 'v.0.5': 40, '주말_근무': 60} 
        for col in columns:
            tree.heading(col, text=col.replace('_', ' '), anchor='center')
            tree.column(col, width=column_widths.get(col, 50), anchor='center', stretch=tk.NO)
        for index, row in summary_df.iterrows(): tree.insert('', 'end', values=row.tolist())
        tree.pack(fill='both', expand=True)

    def display_initial_schedule_table(self):
        try: selected_year, selected_month = self.year_var.get(), self.month_var.get()
        except: return
        if not self.worker_names:
            for widget in self.schedule_frame.winfo_children(): widget.destroy()
            tk.Label(self.schedule_frame, text="근무자 관리 메뉴에서 근무자를 먼저 추가해 주세요.", font=('Malgun Gothic', 14)).pack(pady=100)
            self.month_label_text.set(f"🗓️ {selected_year}년 {selected_month}월 근무표")
            return
        
        loaded_df, loaded_manual_edits = self.load_schedule_from_memory(selected_year, selected_month)
        year, month, last_day, day_columns = self.get_month_days(selected_year, selected_month)

        if loaded_df is not None and not loaded_df.empty:
            self.current_schedule_df = loaded_df
            self.manual_edited_cells = loaded_manual_edits 
            
            # 근무자 목록이 변경된 경우, 이전 근무표는 폐기
            if list(self.current_schedule_df.index) != self.worker_names:
                loaded_df = None
                self.current_schedule_df = pd.DataFrame() 
                self.manual_edited_cells.clear() 
            
        if loaded_df is None:
            self.manual_edited_cells.clear() 
            initial_data = {name: [''] * len(day_columns) for name in self.worker_names}
            df_initial = pd.DataFrame(initial_data).transpose(); df_initial.columns = day_columns
            self.current_schedule_df = df_initial

        self.display_schedule_table(self.current_schedule_df, year, month)
        
        if not self.current_schedule_df.empty and loaded_df is not None:
            self.current_summary_df = self.generate_schedule_summary(self.current_schedule_df, selected_year, selected_month)
        else:
            self.current_summary_df = pd.DataFrame()
            
        self.display_summary_table(self.current_summary_df)

    def go_to_current_month(self):
        """현재 년/월로 이동하고 근무표를 갱신합니다."""
        now = datetime.datetime.now()
        
        self.year_var.set(now.year)
        self.month_var.set(now.month)
        
        messagebox.showinfo("이동 완료", f"현재 날짜인 {now.year}년 {now.month}월로 이동했습니다.")
        
    def load_and_display_data_after_startup(self):
        """창이 뜬 후, 늦게 로드해도 되는 데이터를 처리"""
        self.load_all_schedules() 
        self.display_initial_schedule_table()


    def generate_schedule_summary(self, df_schedule, year, month):
        """근무표 데이터에서 유형별 집계 통계 생성 및 디테일 강화 (DH 카운트 포함)"""
        if df_schedule.empty: return pd.DataFrame()

        DUTY_TYPES = ['D', 'E', 'N', 'MD', 'DH', 'Off', 'V', 'v.25', 'v.0.5']
        summary_df = df_schedule.stack().groupby(level=0).value_counts().unstack(fill_value=0)
        summary_df['Off'] = summary_df.get('O', 0) + summary_df.get('Off', 0)
        summary_df = summary_df.drop(columns=['O'], errors='ignore')

        for col in DUTY_TYPES:
            if col not in summary_df.columns: summary_df[col] = 0

        summary_df = summary_df.reset_index(names=['근무자'])

        # 1. 주말 근무 카운트
        weekend_cols = [col for col in df_schedule.columns if '(토)' in col or '(일)' in col]
        summary_df['주말_근무'] = df_schedule[weekend_cols].apply(lambda row: row.astype(str).str.contains('|'.join(WORK_DUTIES)).sum(), axis=1).values

        # 2. 총 근무, 연차 정보
        summary_df['총 근무'] = summary_df[['D', 'E', 'MD', 'N', 'DH']].sum(axis=1)
        summary_df['전월 연차'] = 2.0 
        summary_df['총 연차'] = 21.5 - (summary_df.get('V', 0) + summary_df.get('v.25', 0) * 0.25 + summary_df.get('v.0.5', 0) * 0.5)

        final_cols = ['근무자', '전월 연차', '총 연차', '총 근무', 'D', 'E', 'DH', 'MD', 'N', 'Off', 'V', 'v.25', 'v.0.5', '주말_근무']
        return summary_df.reindex(columns=final_cols)

    def update_schedule_cell(self, event, tree, combobox, item_id, column_id, col_name, worker_name):
        """Combobox 선택 후 Treeview와 DataFrame을 업데이트, 수동 편집 추적"""
        new_value = combobox.get()
        col_name_df = tree.heading(column_id)['text'] 
        if col_name_df == '근무자': return 
        
        edit_key = (worker_name, col_name)

        if new_value != '': 
            self.manual_edited_cells.add(edit_key)
        else:
            if edit_key in self.manual_edited_cells:
                self.manual_edited_cells.remove(edit_key)

        tree.set(item_id, column_id, new_value)
        
        if not self.current_schedule_df.empty and worker_name in self.current_schedule_df.index and col_name in self.current_schedule_df.columns:
            self.current_schedule_df.loc[worker_name, col_name] = new_value
            
            year, month = self.year_var.get(), self.month_var.get()
            self.save_current_schedule_to_memory(self.current_schedule_df, year, month)
            
            # 수동 편집이 당월의 마지막 5일에 해당되면, 다음 달 연속성을 위해 prev_month_schedule에도 즉시 반영
            if col_name in self.current_schedule_df.columns[-5:]:
                self.save_prev_month_schedule() 
            
            if self.summary_frame:
                self.current_summary_df = self.generate_schedule_summary(self.current_schedule_df, year, month)
                self.display_summary_table(self.current_summary_df)
                
        combobox.destroy()

    def start_schedule_edit(self, event):
        """근무 내용 Treeview 셀 클릭시 수동 수정 Combobox 띄우기"""
        tree = self.current_tree 
        try:
            if hasattr(tree, 'editor_widget') and tree.editor_widget.winfo_exists():
                tree.editor_widget.destroy()

            region = tree.identify_region(event.x, event.y)
            if region != "cell": return

            column_id = tree.identify_column(event.x)
            item_id = tree.identify_row(event.y)
            
            if column_id == '#1': return 

            col_index = int(column_id.replace('#', '')) - 2
            worker_name = tree.item(item_id, 'values')[0] 
            
            col_name_full = self.current_schedule_df.columns[col_index] 
            current_value = tree.set(item_id, column_id)
            bbox = tree.bbox(item_id, column_id)
            if not bbox: return
            x, y, width, height = bbox

            combobox = ttk.Combobox(tree, values=EDITABLE_SHIFTS, width=width, font=('Malgun Gothic', 10), state='readonly')
            combobox.set(current_value)
            combobox.place(x=x, y=y, width=width, height=height)
            combobox.bind("<<ComboboxSelected>>", lambda e: self.update_schedule_cell(e, tree, combobox, item_id, column_id, col_name_full, worker_name))
            combobox.bind("<FocusOut>", lambda e: combobox.destroy() if e.widget == combobox and not combobox.winfo_ismapped() else None)
            combobox.bind("<Return>", lambda e: self.update_schedule_cell(e, tree, combobox, item_id, column_id, col_name_full, worker_name))
            combobox.focus_set();
            tree.editor_widget = combobox
        except Exception as e: logging.error(f"[start_schedule_edit] {e}")

    def generate_monthly_schedule(self, year, month):
        """자동 근무표 생성 핵심 알고리즘 (N-N-N-O-O 및 N근무 연속성 최우선 강제 적용)"""
        year, month, last_day, day_columns = self.get_month_days(year, month);
        if not self.worker_names: return pd.DataFrame(), year, month

        schedule_data = {name: [''] * last_day for name in self.worker_names}
        hn_name = self.worker_names[0] if self.worker_names else None
        start_date = datetime.date(year, month, 1) 
        
        # 1. 수동 편집 데이터 절대 보존 및 초기 반영 
        # "근무표 생성" 버튼은 이전에 자동 생성된 값은 모두 무시하고, 
        # 수동으로 편집된 값(self.manual_edited_cells)만 보존합니다.
        if not self.current_schedule_df.empty:
            df_temp = self.current_schedule_df.copy().fillna('').astype(str)
            for worker in self.worker_names:
                for day_index, col_name in enumerate(day_columns):
                    edit_key = (worker, col_name)
                    if edit_key in self.manual_edited_cells and \
                       worker in df_temp.index and col_name in df_temp.columns:
                        
                        manual_value = df_temp.loc[worker, col_name]
                        schedule_data[worker][day_index] = manual_value

        # 2. 근무 할당 기준 초기화 (수동 편집 데이터 포함하여 카운트)
        daily_n_usage = [0] * last_day
        duty_counts = {name: {d: 0 for d in WORK_DUTIES} for name in self.worker_names}

        for name in self.worker_names:
            for day_index in range(last_day):
                duty = schedule_data[name][day_index]
                if duty == 'N':
                    daily_n_usage[day_index] += 1
                if duty in WORK_DUTIES:
                    duty_counts[name][duty] += 1

        # 3. 수선생님 주간 근무 (HN duties) - 빈 셀에만 할당
        if self.is_head_nurse_mode.get() and hn_name and hn_name in schedule_data and self.worker_categories_map.get(hn_name) == '수선생님':
            for day_index in range(last_day):
                if schedule_data[hn_name][day_index] == '':
                    weekday = (start_date + datetime.timedelta(days=day_index)).weekday()
                    assigned_duty = 'D' if 0 <= weekday <= 4 else 'O'
                    schedule_data[hn_name][day_index] = assigned_duty
                    if assigned_duty == 'D': 
                         duty_counts[hn_name]['D'] += 1

        # 4. N-Set (N-N-N-O-O) 우선 할당 (빈 셀에만 할당)
        
        # '수선생님' 직책 근무자는 N-Set 할당에서 제외
        workers_for_n = [w for w in self.worker_names if self.worker_categories_map.get(w) != '수선생님']
        random.shuffle(workers_for_n)
        n_set_counts = {name: 0 for name in workers_for_n}
        prev_month_duties = self.prev_month_last_day_duties


        # 4-A. ⭐전월 N-Set 연속성 강조 (N-N-N-O-O) 강제 할당⭐
        N_PATTERN = ['N', 'N', 'N', 'O', 'O']

        for name in workers_for_n:
            last_5 = prev_month_duties.get(name, [])
            if not last_5: continue

            # 마지막부터 연속된 N 개수 파악
            n_count = 0
            for d in reversed(last_5):
                if d == 'N': n_count += 1
                else: break

            duties_to_continue = []

            # 1. 전월이 N-Set 연속선상(N, N-N, N-N-N)이었을 경우
            if 1 <= n_count <= 3:
                duties_to_continue = N_PATTERN[n_count:]
            # 2. 전월이 N-N-N-O로 끝났을 경우 (당월 1일은 두 번째 O가 되어야 함)
            elif len(last_5) >= 4 and last_5[-1] == 'O' and last_5[-2] == 'N' and last_5[-3] == 'N' and last_5[-4] == 'N':
                duties_to_continue = ['O']

            # 강제 할당 실행
            for day_idx, duty in enumerate(duties_to_continue):
                if day_idx >= last_day: break
                
                existing_duty = schedule_data[name][day_idx]

                if existing_duty != '':
                    # 💡 FIX: 수동 편집된 셀이 연속성에 필요한 근무와 일치하면, 할당은 건너뛰고 연속성은 유지함
                    if existing_duty == duty:
                        continue 
                    # 수동 편집된 값이 다른 근무라면, N-Set 연속성은 깨진 것으로 간주하고 중단
                    else:
                        break # 연속성 할당 중단
                    
                # 빈 셀인 경우에만 강제 할당
                schedule_data[name][day_idx] = duty
                if duty == 'N':
                    daily_n_usage[day_idx] += 1
                    duty_counts[name]['N'] += 1

        # 4-B. 당월 N-Set 할당 (기존 로직 유지)
        for start_day in range(last_day): 
            if daily_n_usage[start_day] >= DAILY_LIMITS['N'] * 2 or (last_day - start_day) < 3: 
                continue

            n_len = 3 
            o_start_day = start_day + n_len
            o_len = min(2, last_day - o_start_day) 
            block_len = n_len + o_len
            
            available_workers = [
                worker for worker in workers_for_n
                if n_set_counts[worker] < 2 and 
                all(schedule_data[worker][d] == '' for d in range(start_day, start_day + block_len)) and
                daily_n_usage[start_day] < DAILY_LIMITS['N'] and
                (start_day + 1 >= last_day or daily_n_usage[start_day+1] < DAILY_LIMITS['N']) and
                (start_day + 2 >= last_day or daily_n_usage[start_day+2] < DAILY_LIMITS['N'])
            ]
            
            if not available_workers: continue
                
            worker_to_assign = min(available_workers, key=lambda w: n_set_counts[w])
            
            for d in range(start_day, start_day + n_len):
                if d < last_day:
                    schedule_data[worker_to_assign][d] = 'N'
                    daily_n_usage[d] += 1
                    duty_counts[worker_to_assign]['N'] += 1
            
            for d in range(o_start_day, o_start_day + o_len):
                if d < last_day:
                    schedule_data[worker_to_assign][d] = 'O'
            
            n_set_counts[worker_to_assign] += 1
            
        
        # 4-C. ⭐월말 N 근무 강제 할당⭐ (다음 달 N-Set 시작을 위한 N 확보)
        last_day_index = last_day - 1
        n_on_last_day = any(schedule_data[name][last_day_index] == 'N' for name in workers_for_n)
        
        if not n_on_last_day and daily_n_usage[last_day_index] < DAILY_LIMITS['N']: 
            eligible_workers = [
                name for name in workers_for_n
                if schedule_data[name][last_day_index] == ''
            ]
            
            if eligible_workers:
                worker_to_assign = min(eligible_workers, key=lambda w: n_set_counts.get(w, 0))
                schedule_data[worker_to_assign][last_day_index] = 'N'
                daily_n_usage[last_day_index] += 1
                duty_counts[worker_to_assign]['N'] += 1
        

        # 5. 근무 할당 기준 계산
        num_workers_for_duty = len(self.worker_names) - (1 if self.is_head_nurse_mode.get() and hn_name else 0)
        num_work_days = sum(1 for col_str in day_columns if col_str[-2] not in ['토', '일'])
        duties_for_auto_allocation = ['D', 'E', 'N'] 
        target_duty_count_per_worker = max(1, math.ceil(num_work_days * len(duties_for_auto_allocation) / num_workers_for_duty)) if num_workers_for_duty > 0 else 0


        # 6. 일자별 근무 할당 (D, E, Off) - 빈 셀에만 할당
        DAILY_ASSIGNABLE_DUTIES = ['D', 'E'] 

        for day_index in range(last_day):
            date_obj = start_date + datetime.timedelta(days=day_index)
            weekday = date_obj.weekday() 
            current_daily_limits = DAILY_LIMITS.copy()
            if weekday >= 5: current_daily_limits['E'] = 1
            
            daily_duty_counts = {'D': 0, 'E': 0, 'N': 0, 'DH': 0}

            workers_to_schedule = []

            for name in self.worker_names:
                current_duty = schedule_data[name][day_index]
                if current_duty in WORK_DUTIES:
                    daily_duty_counts[current_duty] += 1
                elif current_duty == '':
                    workers_to_schedule.append(name)

            if self.is_head_nurse_mode.get() and hn_name and hn_name in workers_to_schedule and self.worker_categories_map.get(hn_name) == '수선생님':
                workers_to_schedule.remove(hn_name)

            random.shuffle(workers_to_schedule)

            for name in workers_to_schedule:
                assigned_duty = ''

                # 6-1. 이전 근무 확인
                prev_duty = self._get_previous_duty(name, day_index, 1, schedule_data)
                prev_2_duty = self._get_previous_duty(name, day_index, 2, schedule_data)
                prev_3_duty = self._get_previous_duty(name, day_index, 3, schedule_data)
                prev_4_duty = self._get_previous_duty(name, day_index, 4, schedule_data)
                
                # 6-2. 강제 Off 규칙 적용
                if prev_duty == 'N': assigned_duty = 'O'
                elif prev_duty == 'O': 
                    if prev_2_duty == 'N' and prev_3_duty == 'N' and prev_4_duty == 'N':
                         assigned_duty = 'O'
                elif day_index >= 5: 
                    last_5 = schedule_data[name][day_index-5:day_index]
                    if len(last_5) == 5 and all(d in WORK_DUTIES for d in last_5):
                        assigned_duty = 'O'
                
                # 6-3. 자동 순환 및 할당 (D, E만 고려)
                if not assigned_duty:
                    target_rotation = ''
                    forbidden_duties = set()
                    
                    if prev_duty == 'E': forbidden_duties.add('D'); target_rotation = 'E' 
                    if prev_2_duty == 'N' and prev_duty == 'O': forbidden_duties.add('D')
                        
                    if not target_rotation: 
                        if prev_duty == 'D': target_rotation = 'E'
                        else: target_rotation = sorted(DAILY_ASSIGNABLE_DUTIES, key=lambda d: duty_counts[name].get(d, 0))[0]

                    if target_rotation not in DAILY_ASSIGNABLE_DUTIES:
                        target_rotation = sorted(DAILY_ASSIGNABLE_DUTIES, key=lambda d: duty_counts[name].get(d, 0))[0]

                    duties_to_check = [target_rotation] + [d for d in DAILY_ASSIGNABLE_DUTIES if d != target_rotation]
                    
                    for duty_to_check in duties_to_check:
                        if duty_to_check in forbidden_duties: continue 
                        is_daily_full = daily_duty_counts.get(duty_to_check, 0) >= current_daily_limits.get(duty_to_check, float('inf'))
                        if duty_to_check in ['D', 'E', 'N']:
                             is_worker_full = duty_counts[name][duty_to_check] >= target_duty_count_per_worker + 1 
                        else:
                            is_worker_full = False
                        
                        if not is_daily_full and not is_worker_full:
                            assigned_duty = duty_to_check
                            break
                    
                    if not assigned_duty: assigned_duty = 'O'

                schedule_data[name][day_index] = assigned_duty

                if assigned_duty in WORK_DUTIES:
                    duty_counts[name][assigned_duty] += 1
                    daily_duty_counts[assigned_duty] += 1

        df = pd.DataFrame({name: schedule_data[name] for name in self.worker_names}).transpose(); df.columns = day_columns
        return df, year, month

    def generate_and_display(self):
        if not self.worker_names:
            messagebox.showwarning("경고", "근무자가 최소 1명 이상 등록되어야 합니다."); return
        try: selected_year, selected_month = self.year_var.get(), self.month_var.get()
        except tk.TclError: messagebox.showerror("오류", "올바른 년도와 월을 선택해 주세요."); return

        # ⭐ 핵심 수정: 근무표 생성 전에 항상 전월 근무 데이터를 파일에서 새로 불러옵니다.
        self.load_prev_month_schedule() 

        df_schedule, year, month = self.generate_monthly_schedule(selected_year, selected_month)

        self.save_current_schedule_to_memory(df_schedule, year, month)

        self.display_schedule_table(df_schedule, year, month)
        summary_df = self.generate_schedule_summary(df_schedule, year, month)
        self.display_summary_table(summary_df)

        self.current_schedule_df = df_schedule
        self.current_summary_df = summary_df

        self.save_prev_month_schedule() # 새로 생성된 근무표의 마지막 5일을 다음 달을 위해 다시 저장


    def clear_schedule(self):
        if not self.worker_names: messagebox.showwarning("경고", "초기화할 근무자 명단이 없습니다."); return
        try:
            year, month = self.year_var.get(), self.month_var.get()
            if not messagebox.askyesno("확인", f"{year}년 {month}월 근무표를 초기화하시겠습니까? (수동 편집 내용 포함)"): return
            
            key = f"{year}-{month:02d}"
            if key in self.monthly_schedules:
                del self.monthly_schedules[key]
                self.save_all_schedules()
                
            self.manual_edited_cells.clear()
            year, month, last_day, day_columns = self.get_month_days(year, month)
            initial_data = {name: [''] * len(day_columns) for name in self.worker_names}
            df_initial = pd.DataFrame(initial_data).transpose(); df_initial.columns = day_columns
            self.current_schedule_df = df_initial
            self.display_schedule_table(self.current_schedule_df, year, month)
            self.current_summary_df = pd.DataFrame()
            self.display_summary_table(self.current_summary_df)
            self.save_prev_month_schedule() 
        except Exception as e:
            messagebox.showerror("초기화 오류", f"근무표 초기화 중 오류가 발생했습니다: {e}")

    # ----------------------------------------------------------------------
    # [메인 UI 및 이벤트 연결]
    # ----------------------------------------------------------------------

    def show_popup_menu(self, menu_name, parent_button):
        """상단 버튼 클릭 시 팝업 메뉴를 표시"""
        
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
            # 💡 FIX: 엑셀 저장 기능을 self.save_schedule_to_excel 함수에 연결
            menu.add_command(label="데이터 저장 (.xlsx)", command=self.save_schedule_to_excel)

        parent_button.update_idletasks() 
        x = parent_button.winfo_rootx()
        y = parent_button.winfo_rooty() + parent_button.winfo_height() + 1 

        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def setup_main_window(self):
        self.root.title("📅 근무표 생성 시스템"); self.root.configure(bg='white')
        
        try: self.root.iconbitmap('favicon.ico') 
        except tk.TclError: pass
        
        screen_width = self.root.winfo_screenwidth(); screen_height = self.root.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (WINDOW_WIDTH / 2))
        y_cordinate = int((screen_height / 2) - (WINDOW_HEIGHT / 2))
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x_cordinate}+{y_cordinate}")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        style = ttk.Style(); style.theme_use('default')
        style.configure('Toss.TLabel', font=('Malgun Gothic', 16, 'bold'), background='white', foreground='#333333')
        style.configure('Toss.TFrame', background='white')
        
        style.configure('Primary.TButton', font=('Malgun Gothic', 14, 'bold'), foreground='white', background=TOSS_BLUE, padding=[30, 15], relief='flat')
        style.map('Primary.TButton', background=[('active', '#004A99')], foreground=[('active', 'white')])
        
        style.configure('Clear.TButton', 
            font=('Malgun Gothic', 14, 'bold'), 
            foreground='#333333', 
            background='white', 
            padding=[30, 15], 
            relief='solid',     
            borderwidth=2       
        )
        style.map('Clear.TButton', 
            background=[('active', '#F0F0F0')], 
            foreground=[('active', '#333333')]
        )
        
        style.configure('Dialog.Primary.TButton', font=('Malgun Gothic', 11, 'bold'), foreground='white', background=TOSS_BLUE, padding=[15, 8], relief='flat')
        style.map('Dialog.Primary.TButton', background=[('active', '#004A99')], foreground=[('active', 'white')])
        style.configure('Dialog.Secondary.TButton', font=('Malgun Gothic', 11, 'bold'), foreground='#333333', background='#EFEFEF', padding=[15, 8], relief='flat')
        style.map('Dialog.Secondary.TButton', background=[('active', '#DCDCDC')], foreground=[('active', '#333333')])
        
        style.configure('Small.TButton', font=('Malgun Gothic', 12), padding=[10, 5], background='#F0F0F0', relief='flat')
        style.map('Small.TButton', background=[('active', '#E0E0E0')], foreground=[('active', '#333333')])
        style.configure('Menu.TButton', font=('Malgun Gothic', 10, 'bold'), foreground='#333333', background='white', padding=[10, 5], relief='flat')
        style.map('Menu.TButton', background=[('active', '#F0F0F0')], foreground=[('active', TOSS_BLUE)])

        top_bar_frame = ttk.Frame(self.root, style='Toss.TFrame'); top_bar_frame.pack(fill='x', padx=20, pady=(10, 0)) 
        
        file_button = ttk.Button(top_bar_frame, text="파일", style='Menu.TButton')
        file_button.config(command=lambda btn=file_button: self.show_popup_menu('파일', btn))
        file_button.pack(side='left', padx=(0, 5))
        
        worker_button = ttk.Button(top_bar_frame, text="근무자 관리", style='Menu.TButton')
        worker_button.config(command=lambda btn=worker_button: self.show_popup_menu('근무자 관리', btn))
        worker_button.pack(side='left', padx=5)
        
        data_button = ttk.Button(top_bar_frame, text="데이터", style='Menu.TButton')
        data_button.config(command=lambda btn=data_button: self.show_popup_menu('데이터', btn))
        data_button.pack(side='left', padx=5)


        control_frame = ttk.Frame(self.root, style='Toss.TFrame'); control_frame.pack(pady=(5, 5), padx=20)
        tk.Label(control_frame, text="년도:", font=('Malgun Gothic', 12, 'bold'), bg='white').pack(side='left', padx=(0, 5))
        
        # 💡 FIX: 년도 최대값 설정을 5년 후에서 100년 후로 변경 (2030년 -> 2130년)
        ttk.Spinbox(control_frame, from_=CURRENT_YEAR - 5, to=CURRENT_YEAR + 100, textvariable=self.year_var, width=5, font=('Malgun Gothic', 12)).pack(side='left', padx=(0, 15))
        
        tk.Label(control_frame, text="월:", font=('Malgun Gothic', 12, 'bold'), bg='white').pack(side='left', padx=(0, 5))
        ttk.Spinbox(control_frame, from_=0, to=13, textvariable=self.month_var, width=3, font=('Malgun Gothic', 12), wrap=True).pack(side='left', padx=(0, 15))
        
        ttk.Button(control_frame, text="오늘", command=self.go_to_current_month, style='Small.TButton').pack(side='left', padx=(15, 0))

        def on_date_change_cb(*args):
            # 월 오버플로우/언더플로우 로직 (년도 자동 변경)
            try: current_year, current_month = self.year_var.get(), self.month_var.get()
            except Exception: return
            
            if current_month > 12: 
                self.year_var.set(current_year + 1)
                self.month_var.set(1)
            elif current_month < 1: 
                self.year_var.set(current_year - 1)
                self.month_var.set(12)
                
            # 근무표/타이틀 갱신
            try: self.display_initial_schedule_table()
            except Exception as e: logging.error(f"[on_date_change_cb] {e}")

        ttk.Label(self.root, textvariable=self.month_label_text, style='Toss.TLabel').pack(pady=5)

        main_content_frame = ttk.Frame(self.root, style='Toss.TFrame'); main_content_frame.pack(fill='both', expand=True, padx=20, pady=10)
        self.schedule_frame = ttk.Frame(main_content_frame, relief='flat', borderwidth=0, padding=10, style='Toss.TFrame')
        self.schedule_frame.pack(side='left', fill='both', expand=True)
        self.summary_frame = ttk.Frame(main_content_frame, relief='flat', borderwidth=0, padding=10, style='Toss.TFrame')
        self.summary_frame.pack(side='left', fill='y', padx=(10, 0))

        button_container = ttk.Frame(self.root, style='Toss.TFrame'); button_container.pack(pady=30)
        
        ttk.Button(button_container, text="✨ 근무표 생성", command=self.generate_and_display, style='Primary.TButton').pack(side='left', padx=10)
        ttk.Button(button_container, text="🗑️ 근무표 초기화", command=self.clear_schedule, style='Clear.TButton').pack(side='left', padx=10)

        # 년도/월 변수 추적 연결
        self.year_var.trace_add("write", on_date_change_cb)
        self.month_var.trace_add("write", on_date_change_cb)

        self.root.after(100, self.load_and_display_data_after_startup)

        footer_frame = ttk.Frame(self.root, style='Toss.TFrame'); footer_frame.pack(side='bottom', fill='x', padx=10, pady=(0, 5))
        # 💡 FIX: 버전 정보 업데이트
        tk.Label(footer_frame, text="made by TKㅣver.24112542", font=('Malgun Gothic', 9), fg='#AAAAAA', bg='white').pack(side='right', padx=10)

    def on_closing(self):
        """종료 시 근무자 명단 및 카테고리 저장"""
        self.save_worker_names()
        self.save_worker_categories() 
        self.root.destroy()

# ==============================================================================
# 4. 실행 진입점
# ==============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()