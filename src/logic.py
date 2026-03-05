import openpyxl
from src.utils import normalize_string

class OrderParser:
    def __init__(self, filepath):
        self.workbook = openpyxl.load_workbook(filepath, data_only=True)

    def get_sheet_names(self):
        return self.workbook.sheetnames

    def parse_sheet(self, sheet_name):
        sheet = self.workbook[sheet_name]
        
        # Sections columns (1-indexed for openpyxl)
        # Section 1 (Default): Cols A-K (1 - 11)
        # Section 2 (Employee): Cols L-V (12 - 22)
        # Section 3 (Dairy): Cols W-AG (23 - 33)
        
        sections = [
            {'name': 'default', 'start_col': 1, 'end_col': 7},    # A~G
            {'name': 'employee', 'start_col': 13, 'end_col': 24}, # M~X
            {'name': 'dairy', 'start_col': 25, 'end_col': 36}     # Y~AJ
        ]
        
        parsed_data = {
            'default': [],
            'employee': [],
            'dairy': []
        }
        
        # Data starts from row 6 to 34
        
        for section in sections:
            start_col = section['start_col']
            end_col = section['end_col']
            
            for row in sheet.iter_rows(min_row=6, max_row=34, min_col=start_col, max_col=end_col):
                # Check No. column (index 0 relative to the sliced row tuple)
                no_val = row[0].value
                if not no_val:
                    continue
                
                item = {
                    'no': no_val,
                    'name': row[1].value,
                    'spec': row[2].value,
                    'unit': row[3].value,
                    'qty': row[4].value,
                }
                
                # Filter out header rows that might be captured
                if item['name'] and str(item['name']).replace(" ", "") in ["식품명", "품목명"]:
                    continue
                
                # Check for empty name
                if not item['name']:
                    continue

                parsed_data[section['name']].append(item)
                
        return parsed_data

class CatalogParser:
    def __init__(self, filepath):
        self.workbook = openpyxl.load_workbook(filepath, data_only=True)
        self.price_map = {} # (Name, Spec) -> Price

    def parse(self):
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            
            # Find headers dynamically in the first 10 rows
            header_row_idx = -1
            headers = {}
            for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=10, values_only=True), 1):
                temp_headers = {}
                for col_idx, cell_value in enumerate(row, 1):
                    if cell_value and isinstance(cell_value, str):
                        clean_header = normalize_string(cell_value)
                        temp_headers[clean_header] = col_idx
                
                name_key = "식품명" if "식품명" in temp_headers else "품목명" if "품목명" in temp_headers else None
                if name_key and "단가" in temp_headers:
                    headers = temp_headers
                    header_row_idx = row_idx
                    break
                    
            if header_row_idx == -1:
                continue
                
            col_name = headers.get("식품명") or headers.get("품목명")
            col_spec = headers.get("규격") # Might be missing
            col_price = headers.get("단가")
            
            if not col_price or not col_name:
                continue

            # Data starts row after header
            for row in sheet.iter_rows(min_row=header_row_idx + 1, values_only=True):
                # row is tuple. Index = col_idx - 1
                try:
                    name_val = row[col_name - 1]
                    if not name_val:
                        continue
                        
                    spec_val = row[col_spec - 1] if col_spec else ""
                    price_val = row[col_price - 1]
                    
                    if price_val is None:
                        continue
                        
                    clean_name = normalize_string(name_val)
                    clean_spec = normalize_string(spec_val)
                    
                    if clean_name not in self.price_map:
                        self.price_map[clean_name] = []
                    
                    self.price_map[clean_name].append({
                        'spec': clean_spec, 
                        'price': price_val,
                        'original_spec': spec_val,
                        'original_name': name_val
                    })
                    
                except IndexError:
                    continue
        
        return self.price_map

class PriceMatcher:
    def __init__(self, price_catalogs):
        # price_catalogs: dict { Name -> [ {spec, price} ] }
        self.catalogs = price_catalogs

    def _tokens_match(self, spec1, spec2):
        # Split by comma or slash?
        # spec1, spec2 are already normalized (no spaces)
        # But for token matching, we need original separators?
        # Actually normalize_string removes spaces. 
        # "10kg,중국산" -> "10kg,중국산".
        # Let's rely on `in` check or set check if comma exists.
        
        def get_tokens(s):
            return set(s.replace(',', ' ').split())
            
        # Refined normalization for token matching might be needed
        # But simple version: if one is subset of another?
        if spec1 == spec2:
            return True
        return False

    def get_price(self, name, spec):
        clean_name = normalize_string(name)
        clean_spec = normalize_string(spec)
        
        # 1. Name Match
        if clean_name in self.catalogs:
            candidates = self.catalogs[clean_name]
            
            # 2. Strict Spec Match
            for cand in candidates:
                if cand['spec'] == clean_spec:
                    return cand['price']
            
            # 3. Token/Permutation Match (e.g. "중국산,10kg" vs "10kg,중국산")
            # Problem: normalize_string removes spaces. "10kg,중국산" -> "10kg,중국산"
            # It preserves commas.
            for cand in candidates:
                # Simple check: Sort tokens separated by comma
                def sort_tokens(s):
                    return sorted(s.split(','))
                
                if sort_tokens(cand['spec']) == sort_tokens(clean_spec):
                    return cand['price']

            # 4. Blank Spec Match (If Order spec is empty, check candidates)
            if not clean_spec:
                # If there's only one candidate, it's safe to assume.
                # If multiple, it's ambiguous -> return None (User requested strictness/safety)
                if len(candidates) == 1:
                     return candidates[0]['price']
                # If multiple, we don't know which one.
                return None
            
            # 5. Partial Spec Containment (Dangerous? "200g" vs "200g, box")
            # Let's try: if Order Spec is contained in Catalog Spec
            # Ex: Order "200g", Catalog "200g,봉" -> Match?
            for cand in candidates:
                if clean_spec in cand['spec'] or cand['spec'] in clean_spec:
                    # Provide this match
                    return cand['price']
            
            return None
            
        return None
