import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from datetime import datetime, timedelta
import os
import shutil

from src.utils import get_resource_path

class StatementWriter:
    def __init__(self, output_dir):
        self.output_dir = output_dir
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
    def _load_template(self, template_type):
        # Determine template file based on type
        # template_type: 'default', 'employee', 'dairy'
        filename_map = {
            'default': 'transaction_statement.xlsx',
            'employee': 'transaction_statement_employee.xlsx',
            'dairy': 'transaction_statement_dairy.xlsx'
        }
        
        # Use resource path handler
        template_path = get_resource_path(os.path.join('templates', filename_map.get(template_type, 'transaction_statement.xlsx')))
        
        if not os.path.exists(template_path):
            # Fallback: check if it's in the current directory (for some Windows cases)
            fallback = os.path.join(os.getcwd(), 'templates', filename_map.get(template_type, 'transaction_statement.xlsx'))
            if os.path.exists(fallback):
                template_path = fallback
            else:
                raise FileNotFoundError(f"Template not found: {template_path}")
            
        return openpyxl.load_workbook(template_path)

    def write_statement(self, data_list, template_type, date_obj):
        """
        data_list: list of dicts {'no', 'name', 'spec', 'unit', 'qty', 'price', ...}
        template_type: 'default', 'employee', 'dairy'
        date_obj: datetime object for the "Receiving Date" (Input Date)
        """
        wb = self._load_template(template_type)
        ws = wb.active
        
        # Write Dates
        # Valid locations based on dummy template: F5 (Date)
        # However, requirements say: "Input Date (selected)" and "Sending Date (Input - 1)"
        # Let's put Input Date in "입고일" and Sending Date in "발송일" if slots exist.
        # User said: "Header: Sender/Receiver/Contact/SendingDate/ReceivingDate"
        # My dummy template has A4..F5.
        # Let's assume standard locations or just fill "Date" with Input Date for now.
        # User said: "Sending date = Input - 1".
        
        sending_date = date_obj - timedelta(days=1) if date_obj else datetime.now()
        date_str = date_obj.strftime("%Y-%m-%d") if date_obj else ""
        sending_str = sending_date.strftime("%Y-%m-%d")
        
        # In dummy template, F5 is "Date: ". Let's append there? 
        # Or just find the cell and set it.
        # "발송일", "입고일" exact location might need config unless template is fixed.
        # I'll just put date in F5 for now as a placeholder.
        ws['F5'] = f"날짜: {date_str} (발송: {sending_str})"
        
        # Write Data starting Row 9
        start_row = 9
        
        # Style definition for missing price
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        
        for i, item in enumerate(data_list):
            current_row = start_row + i
            
            # Map item keys to columns (1-indexed)
            # 1: No, 2: Name, 3: Spec, 4: Unit, 5: Qty, 6: Price, 7: Amount, 8: Note
            
            ws.cell(row=current_row, column=1, value=item.get('no'))
            ws.cell(row=current_row, column=2, value=item.get('name'))
            ws.cell(row=current_row, column=3, value=item.get('spec'))
            ws.cell(row=current_row, column=4, value=item.get('unit'))
            ws.cell(row=current_row, column=5, value=item.get('qty'))
            
            price = item.get('price')
            qty = item.get('qty')
            
            # Handle Price
            if price is not None:
                ws.cell(row=current_row, column=6, value=price)
                # Calculate Amount
                try:
                    amount = float(price) * float(qty) if qty else 0
                    ws.cell(row=current_row, column=7, value=amount)
                except (ValueError, TypeError):
                    ws.cell(row=current_row, column=7, value=0)
            else:
                # Missing Price -> Yellow Highlight + Empty
                ws.cell(row=current_row, column=6).fill = yellow_fill
                ws.cell(row=current_row, column=6).value = None # Leave blank
                ws.cell(row=current_row, column=7).value = None # Leave blank
            
            # Apply styling to all cells in row
            for col in range(1, 9):
                cell = ws.cell(row=current_row, column=col)
                cell.border = thin_border
                # Alignment?
                if col in [1, 4, 5, 6, 7]: # Numbers centric
                    cell.alignment = Alignment(horizontal='center')
        
        # Generate Output Filename
        # Format: "거래명세서_[Type]_[Date].xlsx"
        name_map = {
            'employee': '_직원용',
            'dairy': '_유제품'
        }
        type_suffix = name_map.get(template_type, "")
        filename = f"거래명세서{type_suffix}_{date_str}.xlsx"
        output_path = os.path.join(self.output_dir, filename)
        
        wb.save(output_path)
        return output_path
