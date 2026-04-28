import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
import queue
from src.logic import OrderParser, CatalogParser, PriceMatcher
from src.writer import StatementWriter
from src.utils import extract_year_from_text, parse_date_from_sheet_name

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("거래명세서 자동화 프로그램")
        self.root.geometry("600x500")
        
        # Variables
        self.order_path = tk.StringVar()
        self.catalog_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.selected_sheet = tk.StringVar()
        self.status_text = tk.StringVar(value="대기 중")
        self.sheets = []
        self.events = queue.Queue()
        self.processing = False
        
        self.create_widgets()
        self.root.after(100, self.process_events)
        
    def create_widgets(self):
        # Frame for File Selection
        frame_input = tk.LabelFrame(self.root, text="파일 선택", padx=10, pady=10)
        frame_input.pack(fill="x", padx=10, pady=5)
        
        # Order File
        tk.Label(frame_input, text="발주서 파일:").grid(row=0, column=0, sticky="w")
        tk.Entry(frame_input, textvariable=self.order_path, width=40).grid(row=0, column=1)
        tk.Button(frame_input, text="선택", command=self.select_order_file).grid(row=0, column=2, padx=5)
        
        # Catalog File
        tk.Label(frame_input, text="카탈로그 파일:").grid(row=1, column=0, sticky="w")
        tk.Entry(frame_input, textvariable=self.catalog_path, width=40).grid(row=1, column=1)
        tk.Button(frame_input, text="선택", command=self.select_catalog_file).grid(row=1, column=2, padx=5)
        
        # Sheet Selection
        frame_sheet = tk.Frame(self.root)
        frame_sheet.pack(fill="x", padx=10, pady=5)
        tk.Label(frame_sheet, text="날짜(시트) 선택:").pack(side="left")
        self.sheet_combo = ttk.Combobox(frame_sheet, textvariable=self.selected_sheet, state="readonly", width=30)
        self.sheet_combo.pack(side="left", padx=5)
        
        # Output Folder
        frame_output = tk.LabelFrame(self.root, text="출력 폴더", padx=10, pady=10)
        frame_output.pack(fill="x", padx=10, pady=5)
        tk.Entry(frame_output, textvariable=self.output_path, width=40).grid(row=0, column=0)
        tk.Button(frame_output, text="폴더 선택", command=self.select_output_folder).grid(row=0, column=1, padx=5)
        
        # Action Button
        self.action_button = tk.Button(self.root, text="거래명세서 생성", command=self.start_process, bg="lightblue", height=2)
        self.action_button.pack(fill="x", padx=10, pady=10)

        frame_status = tk.Frame(self.root)
        frame_status.pack(fill="x", padx=10, pady=2)
        tk.Label(frame_status, textvariable=self.status_text, anchor="w").pack(side="left", fill="x", expand=True)
        self.progress = ttk.Progressbar(frame_status, mode="indeterminate", length=140)
        self.progress.pack(side="right")
        
        # Log Area
        self.log_area = scrolledtext.ScrolledText(self.root, height=10)
        self.log_area.pack(fill="both", padx=10, pady=5, expand=True)

    def select_order_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.order_path.set(path)
            self.log(f"발주서 파일 선택됨: {os.path.basename(path)}")
            if not self.output_path.get():
                self.output_path.set(os.path.join(os.path.dirname(path), "output"))
            # Load sheets
            try:
                parser = OrderParser(path)
                self.sheets = parser.get_sheet_names()
                self.sheet_combo['values'] = self.sheets
                if self.sheets:
                    self.sheet_combo.current(0)
                self.log(f"시트 로드 완료: {len(self.sheets)}개 발견")
            except Exception as e:
                messagebox.showerror("Error", f"발주서 로드 실패: {e}")

    def select_catalog_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.catalog_path.set(path)
            self.log(f"카탈로그 파일 선택됨: {os.path.basename(path)}")

    def select_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_path.set(path)

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def post_event(self, event_type, payload=None):
        self.events.put((event_type, payload))

    def process_events(self):
        try:
            while True:
                event_type, payload = self.events.get_nowait()
                if event_type == "log":
                    self.log(payload)
                elif event_type == "status":
                    self.status_text.set(payload)
                elif event_type == "done":
                    self.finish_process(payload)
                elif event_type == "error":
                    self.fail_process(payload)
        except queue.Empty:
            pass
        self.root.after(100, self.process_events)

    def start_process(self):
        if self.processing:
            return
        if not self.order_path.get() or not self.catalog_path.get() or not self.selected_sheet.get() or not self.output_path.get():
            messagebox.showwarning("입력 확인", "모든 항목을 선택해주세요.")
            return

        self.processing = True
        self.action_button.configure(state="disabled")
        self.progress.start(10)
        self.status_text.set("처리 중")
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            self.post_event("log", "작업 시작...")
            
            # 1. Parse Catalog
            self.post_event("status", "카탈로그 읽는 중")
            self.post_event("log", "카탈로그 읽는 중...")
            catalog_parser = CatalogParser(self.catalog_path.get())
            price_map = catalog_parser.parse()
            if not price_map:
                raise ValueError("카탈로그에서 제품명/단가 헤더를 찾지 못했습니다. 제품명, 품목명, 상품명, 단가, 가격 등의 헤더를 확인해주세요.")
            matcher = PriceMatcher(price_map)
            self.post_event("log", f"카탈로그 로드 완료. {len(price_map)}개 품목 식별됨.")
            
            # 2. Parse Order Sheet
            sheet_name = self.selected_sheet.get()
            self.post_event("status", "발주서 읽는 중")
            self.post_event("log", f"발주서 시트 '{sheet_name}' 읽는 중...")
            order_parser = OrderParser(self.order_path.get())
            order_data = order_parser.parse_sheet(sheet_name)
            total_items = sum(len(items) for items in order_data.values())
            if total_items == 0:
                raise ValueError("선택한 시트에서 발주 품목을 찾지 못했습니다. 품목명/제품명, 규격/용량, 수량 헤더가 있는지 확인해주세요.")
            
            # Date Parsing
            fallback_year = (
                extract_year_from_text(os.path.basename(self.order_path.get()))
                or extract_year_from_text(os.path.basename(self.catalog_path.get()))
            )
            date_obj = parse_date_from_sheet_name(sheet_name, fallback_year=fallback_year)
            
            # 3. Match and Write
            writer = StatementWriter(self.output_path.get())
            
            summary = []
            all_items = []
            totals = {"matched": 0, "review": 0, "unmatched": 0}
            
            for section_name, items in order_data.items():
                if not items:
                    continue
                    
                self.post_event("status", f"{section_name} 매칭 중")
                self.post_event("log", f"[{section_name}] 섹션 처리 중 ({len(items)}개 품목)...")
                
                # Match Prices
                section_counts = {"matched": 0, "review": 0, "unmatched": 0}
                for item in items:
                    match = matcher.match(item.get('name'), item.get('spec'), item.get('category'))
                    status = match.get("status", "unmatched")
                    section_counts[status] = section_counts.get(status, 0) + 1
                    totals[status] = totals.get(status, 0) + 1

                    if status == "review":
                        self.post_event(
                            "log",
                            f"  [검토필요] {item.get('name')} ({item.get('spec')}) -> "
                            f"{match.get('catalog_name')} ({match.get('catalog_spec')}) / {match.get('confidence')}점"
                        )
                    elif status == "unmatched":
                        self.post_event("log", f"  [미매칭] {item.get('name')} ({item.get('spec')})")

                    item['match'] = match
                    item['price'] = match.get("price")
                    all_items.append(item)
                
                # Write to File
                out_file = writer.write_statement(items, section_name, date_obj)
                summary.append(
                    f"{section_name}: {out_file} "
                    f"(확정 {section_counts.get('matched', 0)}건, 검토 {section_counts.get('review', 0)}건, 미매칭 {section_counts.get('unmatched', 0)}건)"
                )
            
            report_file = writer.write_review_report(all_items, date_obj)
            summary.append(f"검토 리포트: {report_file}")
            summary.append(
                f"전체 요약: 확정 {totals.get('matched', 0)}건, "
                f"검토 {totals.get('review', 0)}건, 미매칭 {totals.get('unmatched', 0)}건"
            )

            self.post_event("done", summary)
            
        except Exception as e:
            self.post_event("error", str(e))

    def finish_process(self, summary):
        self.processing = False
        self.progress.stop()
        self.action_button.configure(state="normal")
        self.status_text.set("완료")
        self.log("작업 완료!")
        for line in summary:
            self.log(line)
        messagebox.showinfo("완료", "거래명세서 생성이 완료되었습니다.\n" + "\n".join(summary))

    def fail_process(self, message):
        self.processing = False
        self.progress.stop()
        self.action_button.configure(state="normal")
        self.status_text.set("오류")
        self.log(f"오류 발생: {message}")
        messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{message}")
