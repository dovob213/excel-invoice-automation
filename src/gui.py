import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
from src.logic import OrderParser, CatalogParser, PriceMatcher
from src.writer import StatementWriter
from src.utils import parse_date_from_sheet_name

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
        self.sheets = []
        
        self.create_widgets()
        
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
        tk.Button(self.root, text="거래명세서 생성", command=self.start_process, bg="lightblue", height=2).pack(fill="x", padx=10, pady=10)
        
        # Log Area
        self.log_area = scrolledtext.ScrolledText(self.root, height=10)
        self.log_area.pack(fill="both", padx=10, pady=5, expand=True)

    def select_order_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.order_path.set(path)
            self.log(f"발주서 파일 선택됨: {os.path.basename(path)}")
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

    def start_process(self):
        if not self.order_path.get() or not self.catalog_path.get() or not self.selected_sheet.get() or not self.output_path.get():
            messagebox.showwarning("입력 확인", "모든 항목을 선택해주세요.")
            return
            
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            self.log("작업 시작...")
            
            # 1. Parse Catalog
            self.log("카탈로그 읽는 중...")
            catalog_parser = CatalogParser(self.catalog_path.get())
            price_map = catalog_parser.parse()
            matcher = PriceMatcher(price_map)
            self.log(f"카탈로그 로드 완료. {len(price_map)}개 품목 식별됨.")
            
            # 2. Parse Order Sheet
            sheet_name = self.selected_sheet.get()
            self.log(f"발주서 시트 '{sheet_name}' 읽는 중...")
            order_parser = OrderParser(self.order_path.get())
            order_data = order_parser.parse_sheet(sheet_name)
            
            # Date Parsing
            date_obj = parse_date_from_sheet_name(sheet_name)
            
            # 3. Match and Write
            writer = StatementWriter(self.output_path.get())
            
            summary = []
            
            for section_name, items in order_data.items():
                if not items:
                    continue
                    
                self.log(f"[{section_name}] 섹션 처리 중 ({len(items)}개 품목)...")
                
                # Match Prices
                unmatched_count = 0
                for item in items:
                    price = matcher.get_price(item['name'], item['spec'])
                    if price is None:
                        unmatched_count += 1
                        self.log(f"  [미매칭] {item['name']} ({item['spec']})")
                    item['price'] = price
                
                # Write to File
                out_file = writer.write_statement(items, section_name, date_obj)
                summary.append(f"{section_name}: {out_file} (미매칭 {unmatched_count}건)")
            
            self.log("작업 완료!")
            for s in summary:
                self.log(s)
                
            messagebox.showinfo("완료", "거래명세서 생성이 완료되었습니다.\n" + "\n".join(summary))
            
        except Exception as e:
            self.log(f"오류 발생: {e}")
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{e}")
