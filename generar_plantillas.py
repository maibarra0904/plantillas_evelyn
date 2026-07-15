import os
import xlrd
import xlwt
from datetime import datetime

# Input source files directory and target output directories
SOURCE_DIR = "."
OUTPUT_DIR = "PLANTILLAS"
TEMPLATE_REF_DIR = "PLANTILLAS_REFERENCE"

# Make sure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Formatting styles for xlwt
style_header = xlwt.easyxf('font: bold on; align: horiz left')
style_general = xlwt.easyxf('align: horiz left')
style_date = xlwt.easyxf('align: horiz left', num_format_str='dd/mm/yyyy')
style_number = xlwt.easyxf('align: horiz left', num_format_str='#,##0.00')

def clean_amount(val):
    if isinstance(val, (int, float)):
        return abs(float(val))
    val_str = str(val).replace('$', '').replace(' ', '').replace(',', '').strip()
    if not val_str:
        return 0.0
    val_str = val_str.replace('-', '')
    
    # Handle multiple dots (e.g., Huancavilca '4.413.66' -> 4413.66)
    if val_str.count('.') > 1:
        parts = val_str.split('.')
        val_str = "".join(parts[:-1]) + "." + parts[-1]
        
    try:
        return abs(float(val_str))
    except:
        return 0.0

def parse_date(val, datemode, bank_name):
    # Determine expected month based on directory path (5 for Mayo, 6 for June/default)
    expected_month = 6
    cwd = os.getcwd().lower()
    script_dir = os.path.dirname(os.path.abspath(__file__)).lower()
    if "mayo" in cwd or "may" in cwd or "mayo" in script_dir or "may" in script_dir:
        expected_month = 5

    if isinstance(val, (int, float)):
        try:
            dt_tuple = xlrd.xldate_as_tuple(val, datemode)
            dt = datetime(dt_tuple[0], dt_tuple[1], dt_tuple[2])
            # Check if we need to swap day and month
            if dt.month != expected_month:
                if dt.day <= 12:
                    swapped = datetime(dt.year, dt.day, dt.month)
                    if swapped.month == expected_month:
                        dt = swapped
            return dt
        except Exception:
            pass
            
    val_str = str(val).strip()
    if not val_str:
        return None
        
    # Remove time part if present
    val_str = val_str.split(' ')[0]
    
    # Try parsing formats. Prefer the one that matches expected_month
    parsed_dt = None
    formats = ['%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%d-%m-%Y']
    for fmt in formats:
        try:
            dt = datetime.strptime(val_str, fmt)
            if dt.month == expected_month:
                parsed_dt = dt
                break
        except ValueError:
            continue
            
    if parsed_dt is None:
        # Fallback to first format that parses
        for fmt in formats:
            try:
                parsed_dt = datetime.strptime(val_str, fmt)
                break
            except ValueError:
                continue
                
    if parsed_dt is None:
        return None
        
    # If month doesn't match expected_month and day is <= 12, swap them
    if parsed_dt.month != expected_month:
        try:
            if parsed_dt.day <= 12:
                swapped = datetime(parsed_dt.year, parsed_dt.day, parsed_dt.month)
                if swapped.month == expected_month:
                    parsed_dt = swapped
        except ValueError:
            pass
            
    return parsed_dt

def write_template(output_filename, rows):
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Hoja1')
    
    headers = [
        "NUMERO DE ESTADO DE CUENTA",
        "CLASE DE DOCUMENTO",
        "NUMERO DE DOCUMENTO",
        "FECHA DE DOCUMENTO",
        "VALOR",
        "TIPO DE MOVIMIENTO",
        "DESCRIPCION"
    ]
    
    # Write headers
    for c, h in enumerate(headers):
        ws.write(0, c, h, style_header)
        
    # Write data rows
    for r, row in enumerate(rows, start=1):
        acc_num, doc_class, seq_num, date_obj, val, mov_type, desc = row
        
        ws.write(r, 0, acc_num, style_general)
        ws.write(r, 1, doc_class, style_general)
        ws.write(r, 2, seq_num, style_general)
        
        if isinstance(date_obj, datetime):
            ws.write(r, 3, date_obj, style_date)
        else:
            ws.write(r, 3, date_obj, style_general)
            
        ws.write(r, 4, val, style_number)
        ws.write(r, 5, mov_type, style_general)
        ws.write(r, 6, desc, style_general)
        
    wb.save(output_filename)

def process_bolivariano(src_sheet, datemode):
    # Dynamic header search
    header_row_idx = 12
    for r in range(min(25, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'fecha' in row_str and 'secuencial' in row_str:
            header_row_idx = r
            break

    rows = []
    seq = 1.0
    for r in range(header_row_idx + 1, src_sheet.nrows):
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 11 or not str(row_vals[0]).strip():
            continue
        
        sign = str(row_vals[6]).strip()
        desc = str(row_vals[10]).strip()
        
        if 'DEPOSITO' in desc.upper():
            doc_class = 'DP'
        else:
            doc_class = 'NC' if sign == '+' else 'ND'
            
        mov_type = 'H' if sign == '+' else 'D'
        
        date_obj = parse_date(row_vals[1], datemode, 'bolivariano')
        val = clean_amount(row_vals[7])
        
        rows.append(['NO_ENCONTRADO', doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_central(src_sheet, datemode):
    # Dynamic account extraction
    acc_num = ''
    for r in range(min(15, src_sheet.nrows)):
        for c in range(src_sheet.ncols):
            val_str = str(src_sheet.cell_value(r, c)).strip().upper()
            if val_str == 'CUENTA CORRIENTE:':
                if c + 1 < src_sheet.ncols:
                    next_val = str(src_sheet.cell_value(r, c + 1)).strip()
                    acc_num = next_val.split(' ')[0]
                    break
        if acc_num:
            break
    if not acc_num:
        acc_num = '69310007'

    # Dynamic header search
    header_row_idx = 4
    for r in range(min(15, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'fecha' in row_str and 'contracuenta' in row_str:
            header_row_idx = r
            break
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 2, src_sheet.nrows): # skip SALDO ANTERIOR row too
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 11:
            continue
            
        tipo = str(row_vals[2]).strip()
        if 'SALDO ANTERIOR' in tipo or not tipo:
            continue
            
        # skip rows with empty debits and credits
        debit = clean_amount(row_vals[6])
        credit = clean_amount(row_vals[7])
        if debit == 0.0 and credit == 0.0:
            continue
            
        doc_class = 'NC' if 'CREDITO' in tipo else 'ND'
        mov_type = 'H' if 'CREDITO' in tipo else 'D'
        val = credit if 'CREDITO' in tipo else debit
        
        date_obj = parse_date(row_vals[0], datemode, 'central')
        desc = str(row_vals[10]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_guayaquil(src_sheet, datemode):
    # Dynamic account extraction
    acc_num = ''
    for r in range(min(15, src_sheet.nrows)):
        for c in range(src_sheet.ncols):
            val_str = str(src_sheet.cell_value(r, c)).strip().upper()
            if val_str == 'CUENTA:':
                acc_num = val_str.split(' ')[-1]
                break
        if acc_num:
            break
    if not acc_num:
        acc_num = '0011148395'

    # Dynamic header search
    header_row_idx = 13
    for r in range(min(20, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if any('fecha de transacción' in x or 'fecha de transaccion' in x for x in row_str):
            header_row_idx = r
            break
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 1, src_sheet.nrows):
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 14 or not str(row_vals[0]).strip():
            continue
            
        mov_name = str(row_vals[3]).strip().lower()
        if 'crédito' in mov_name or 'credito' in mov_name:
            doc_class = 'NC'
        elif 'débito' in mov_name or 'debito' in mov_name:
            doc_class = 'ND'
        elif 'depósito' in mov_name or 'deposito' in mov_name:
            doc_class = 'DP'
        else:
            doc_class = 'NC'
            
        sign = str(row_vals[13]).strip()
        mov_type = 'H' if sign == '+' else 'D'
        
        # Use col 1 (Fecha de transacción) instead of col 2 (Fecha contable)
        date_obj = parse_date(row_vals[1], datemode, 'guayaquil')
        val = clean_amount(row_vals[7])
        desc = str(row_vals[5]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_huancavilca(src_sheet, datemode):
    # Dynamic account extraction
    acc_num = ''
    for r in range(min(15, src_sheet.nrows)):
        for c in range(src_sheet.ncols):
            val_str = str(src_sheet.cell_value(r, c)).strip().upper()
            if val_str == 'CUENTA:':
                if c + 1 < src_sheet.ncols:
                    next_val = str(src_sheet.cell_value(r, c + 1)).strip()
                    acc_num = next_val.replace('XXXXXXXX', '56000000')
                    break
        if acc_num:
            break
    if not acc_num:
        acc_num = '560000000008'

    # Dynamic header search
    header_row_idx = 15
    for r in range(min(20, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'orden' in row_str and 'nro.' in row_str:
            header_row_idx = r
            break
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 1, src_sheet.nrows):
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 10 or not str(row_vals[0]).strip():
            continue
            
        doc_type = str(row_vals[4]).strip()
        doc_class = 'ND' if doc_type == 'N/D' else 'NC'
        mov_type = 'D' if doc_type == 'N/D' else 'H'
        
        date_obj = parse_date(row_vals[1], datemode, 'huancavilca')
        
        if doc_class == 'ND':
            val = clean_amount(row_vals[8]) # Retiros
        else:
            val = clean_amount(row_vals[7]) # Depósitos
            
        # Skip rows with zero value
        if val == 0.0:
            continue
            
        desc = str(row_vals[5]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_jep(src_sheet, datemode):
    # Dynamic account extraction
    acc_num = ''
    for r in range(min(15, src_sheet.nrows)):
        for c in range(src_sheet.ncols):
            val_str = str(src_sheet.cell_value(r, c)).strip().upper()
            if val_str.startswith('CUENTA:'):
                acc_num = val_str.split(':')[-1]
                break
        if acc_num:
            break
    if not acc_num:
        acc_num = '560096696570'

    # Dynamic header search
    header_row_idx = 6
    for r in range(min(15, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'fecha' in row_str and 'd/c' in row_str:
            header_row_idx = r
            break
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 1, src_sheet.nrows):
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 7 or not str(row_vals[0]).strip():
            continue
            
        dc = str(row_vals[1]).strip().upper()
        doc_class = 'ND' if dc == 'DEBITO' else 'NC'
        mov_type = 'D' if dc == 'DEBITO' else 'H'
        
        date_obj = parse_date(row_vals[0], datemode, 'jep')
        val = clean_amount(row_vals[5])
        desc = str(row_vals[3]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_machala(src_sheet, datemode):
    # Dynamic account extraction
    acc_num = ''
    for r in range(min(15, src_sheet.nrows)):
        for c in range(src_sheet.ncols):
            val_str = str(src_sheet.cell_value(r, c)).strip().upper()
            if val_str == 'CUENTA':
                if c + 1 < src_sheet.ncols:
                    next_val = str(src_sheet.cell_value(r, c + 1)).strip()
                    acc_num = next_val.replace('XXXXXXX', '1260022')
                    break
        if acc_num:
            break
    if not acc_num:
        acc_num = '1260022195'

    # Dynamic header search
    header_row_idx = 11
    for r in range(min(20, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'fecha' in row_str and ('código transacción' in row_str or 'codigo transaccion' in row_str):
            header_row_idx = r
            break
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 1, src_sheet.nrows):
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 12 or not str(row_vals[0]).strip():
            continue
            
        tipo = str(row_vals[5]).strip()
        if tipo in ['ND', 'NDI']:
            doc_class = 'ND'
            mov_type = 'D'
            val = clean_amount(row_vals[9]) # Débito
        elif tipo == 'CRD':
            doc_class = 'NC'
            mov_type = 'H'
            val = clean_amount(row_vals[10]) # Crédito
        elif tipo == 'DP':
            doc_class = 'DP'
            mov_type = 'H'
            val = clean_amount(row_vals[10]) # Crédito
        else:
            doc_class = 'NC'
            mov_type = 'H'
            val = clean_amount(row_vals[10])
            
        date_obj = parse_date(row_vals[2], datemode, 'machala')
        desc = str(row_vals[6]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_pacifico(src_sheet, datemode):
    # Dynamic account extraction
    acc_num = ''
    for r in range(min(30, src_sheet.nrows)):
        for c in range(src_sheet.ncols):
            val_str = str(src_sheet.cell_value(r, c)).strip().upper()
            if 'CUENTA CORRIENTE' in val_str:
                acc_num = val_str.split(' ')[-1]
                break
        if acc_num:
            break
            
    if not acc_num:
        acc_num = '7503911'
    acc_num = acc_num.replace('#', '')

    # Dynamic header search
    header_row_idx = 5
    for r in range(min(30, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'fecha contable' in row_str and 'nut' in row_str:
            header_row_idx = r
            break
            
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 1, src_sheet.nrows):
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 11 or not str(row_vals[0]).strip():
            continue
            
        if str(row_vals[0]).strip().lower() == 'fecha contable':
            continue
            
        tipo = str(row_vals[3]).strip()
        if tipo == 'N/C':
            doc_class = 'NC'
            mov_type = 'H'
        elif tipo == 'DEP':
            doc_class = 'DP'
            mov_type = 'H'
        elif tipo == 'N/D':
            doc_class = 'ND'
            mov_type = 'D'
        else:
            doc_class = 'NC'
            mov_type = 'H'
            
        date_obj = parse_date(row_vals[0], datemode, 'pacifico')
        val = clean_amount(row_vals[5])
        desc = str(row_vals[7]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_pichincha(src_sheet, datemode):
    # Dynamic header search
    header_row_idx = 8
    for r in range(min(20, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'codigo' in row_str and 'documento' in row_str:
            header_row_idx = r
            break
            
    acc_num = '2100109732'
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 1, src_sheet.nrows):
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 8 or not str(row_vals[0]).strip():
            continue
            
        concept = str(row_vals[3]).strip().upper()
        tipo = str(row_vals[4]).strip().upper()
        
        if 'DEPOSITO' in concept:
            doc_class = 'DP'
            mov_type = 'H'
        elif tipo == 'C':
            doc_class = 'NC'
            mov_type = 'H'
        elif tipo == 'D':
            doc_class = 'ND'
            mov_type = 'D'
        else:
            doc_class = 'NC'
            mov_type = 'H'
            
        date_obj = parse_date(row_vals[0], datemode, 'pichincha')
        val = clean_amount(row_vals[5])
        desc = str(row_vals[3]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_produbanco(src_sheet, datemode):
    # Dynamic account extraction
    acc_num = ''
    for r in range(min(15, src_sheet.nrows)):
        for c in range(src_sheet.ncols):
            val_str = str(src_sheet.cell_value(r, c)).strip().upper()
            if val_str == 'CUENTA:':
                if c + 1 < src_sheet.ncols:
                    acc_num = str(src_sheet.cell_value(r, c + 1)).strip()
                    break
        if acc_num:
            break
    if not acc_num:
        acc_num = '72005000106'

    # Dynamic header search
    header_row_idx = 9
    for r in range(min(15, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'referencia' in row_str and 'saldo contable' in row_str:
            header_row_idx = r
            break
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 1, src_sheet.nrows):
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 8 or not str(row_vals[0]).strip():
            continue
            
        sign = str(row_vals[3]).strip()
        doc_class = 'NC' if sign == '+' else 'ND'
        mov_type = 'H' if sign == '+' else 'D'
        
        date_obj = parse_date(row_vals[0], datemode, 'produbanco')
        val = clean_amount(row_vals[4])
        desc = str(row_vals[2]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_reporne(src_sheet, datemode):
    # Dynamic header search
    header_row_idx = 4
    for r in range(min(10, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'fecha' in row_str and 'contracuenta' in row_str:
            header_row_idx = r
            break
            
    acc_num = '69310007'
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 2, src_sheet.nrows): # skip SALDO ANTERIOR row too
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 11:
            continue
            
        tipo = str(row_vals[2]).strip()
        if 'SALDO ANTERIOR' in tipo or not tipo:
            continue
            
        debit = clean_amount(row_vals[6])
        credit = clean_amount(row_vals[7])
        if debit == 0.0 and credit == 0.0:
            continue
            
        doc_class = 'NC' if 'CREDITO' in tipo else 'ND'
        mov_type = 'H' if 'CREDITO' in tipo else 'D'
        val = credit if 'CREDITO' in tipo else debit
        
        date_obj = parse_date(row_vals[0], datemode, 'reporne')
        desc = str(row_vals[10]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_sagrario(src_sheet, datemode):
    # Dynamic header search
    header_row_idx = 5
    for r in range(min(15, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'tipo de comprobante' in row_str and 'detalle' in row_str:
            header_row_idx = r
            break
            
    acc_num = ''
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 1, src_sheet.nrows):
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 7 or not str(row_vals[0]).strip():
            continue
            
        debit = clean_amount(row_vals[4])
        credit = clean_amount(row_vals[5])
        
        doc_class = 'NC' if credit > 0 else 'ND'
        mov_type = 'H' if credit > 0 else 'D'
        val = credit if credit > 0 else debit
        
        date_obj = parse_date(row_vals[0], datemode, 'sagrario')
        desc = str(row_vals[3]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def process_transferunion(src_sheet, datemode):
    # Dynamic header search
    header_row_idx = 4
    for r in range(min(10, src_sheet.nrows)):
        row_str = [str(x).lower() for x in src_sheet.row_values(r)]
        if 'fecha' in row_str and 'contracuenta' in row_str:
            header_row_idx = r
            break
            
    acc_num = '69310007'
    
    rows = []
    seq = 1.0
    for r in range(header_row_idx + 2, src_sheet.nrows): # skip SALDO ANTERIOR row too
        row_vals = src_sheet.row_values(r)
        if len(row_vals) < 11:
            continue
            
        tipo = str(row_vals[2]).strip()
        if 'SALDO ANTERIOR' in tipo or not tipo:
            continue
            
        debit = clean_amount(row_vals[6])
        credit = clean_amount(row_vals[7])
        if debit == 0.0 and credit == 0.0:
            continue
            
        doc_class = 'NC' if 'CREDITO' in tipo else 'ND'
        mov_type = 'H' if 'CREDITO' in tipo else 'D'
        val = credit if 'CREDITO' in tipo else debit
        
        date_obj = parse_date(row_vals[0], datemode, 'transferunion')
        desc = str(row_vals[10]).strip()
        
        rows.append([acc_num, doc_class, seq, date_obj, val, mov_type, desc])
        seq += 1.0
    return rows

def main():
    banks = [
        "bolivariano", "central", "guayaquil", "huancavilca", "jep",
        "machala", "pacifico", "pichincha", "produbanco", "reporne",
        "sagrario", "transferunion"
    ]
    
    print("=== Generating templates into PLANTILLAS_ENTRENAMIENTO ===")
    for bank in banks:
        source_file = os.path.join(SOURCE_DIR, f"banco_{bank}.xls")
        output_file = os.path.join(OUTPUT_DIR, f"plantilla_banco_{bank}.xls")
        
        if not os.path.exists(source_file):
            print(f"[-] Source file not found: {source_file}")
            continue
            
        print(f"[+] Processing {bank}...")
        try:
            wb = xlrd.open_workbook(source_file)
            sheet = wb.sheet_by_index(0)
            datemode = wb.datemode
            
            func_name = f"process_{bank}"
            if func_name in globals():
                rows = globals()[func_name](sheet, datemode)
                write_template(output_file, rows)
                print(f"    Saved {len(rows)} rows to {output_file}")
            else:
                print(f"    [-] No process function found for {bank}")
        except Exception as e:
            print(f"    [-] Error processing {bank}: {e}")
            import traceback
            traceback.print_exc()

    # Run verification against PLANTILLAS
    print("\n=== Verifying generated templates against PLANTILLAS ===")
    mismatches = 0
    matches = 0
    for bank in banks:
        ref_file = os.path.join(TEMPLATE_REF_DIR, f"plantilla_banco_{bank}.xls")
        gen_file = os.path.join(OUTPUT_DIR, f"plantilla_banco_{bank}.xls")
        
        if not os.path.exists(ref_file):
            print(f"[-] Reference file not found: {ref_file}")
            continue
        if not os.path.exists(gen_file):
            print(f"[-] Generated file not found: {gen_file}")
            continue
            
        try:
            ref_wb = xlrd.open_workbook(ref_file)
            ref_sheet = ref_wb.sheet_by_index(0)
            
            gen_wb = xlrd.open_workbook(gen_file)
            gen_sheet = gen_wb.sheet_by_index(0)
            
            if ref_sheet.nrows != gen_sheet.nrows:
                print(f"[-] {bank}: Row count mismatch. Ref={ref_sheet.nrows}, Gen={gen_sheet.nrows}")
                mismatches += 1
                continue
                
            if ref_sheet.ncols != gen_sheet.ncols:
                print(f"[-] {bank}: Col count mismatch. Ref={ref_sheet.ncols}, Gen={gen_sheet.ncols}")
                mismatches += 1
                continue
                
            bank_matched = True
            for r in range(ref_sheet.nrows):
                for c in range(ref_sheet.ncols):
                    ref_val = ref_sheet.cell_value(r, c)
                    gen_val = gen_sheet.cell_value(r, c)
                    
                    if isinstance(ref_val, float) and isinstance(gen_val, float):
                        if abs(ref_val - gen_val) > 1e-6:
                            print(f"  [-] Mismatch in {bank} @ Row {r}, Col {c}. Ref={repr(ref_val)}, Gen={repr(gen_val)}")
                            bank_matched = False
                    else:
                        if str(ref_val).strip() != str(gen_val).strip():
                            print(f"  [-] Mismatch in {bank} @ Row {r}, Col {c}. Ref={repr(ref_val)}, Gen={repr(gen_val)}")
                            bank_matched = False
                            
            if bank_matched:
                print(f"[SUCCESS] {bank}: Matches reference exactly!")
                matches += 1
            else:
                mismatches += 1
                
        except Exception as e:
            print(f"[-] Error verifying {bank}: {e}")
            mismatches += 1
            
    print(f"\nVerification summary: {matches} matched, {mismatches} mismatched.")

if __name__ == "__main__":
    main()
