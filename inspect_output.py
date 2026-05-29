import openpyxl

file_path = 'output_results/EMP0001.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    print(f"Checking {file_path} (Active Sheet: {ws.title})")
    
    # Check some key cells from the template
    # Title is in B2:E3, Employee ID in C5, Name in E5
    cells_to_check = ['B2', 'C5', 'E5', 'C6', 'E6', 'C7', 'E7', 'C8']
    for addr in cells_to_check:
        val = ws[addr].value
        print(f"Cell {addr}: {val}")
    
    wb.close()
except Exception as e:
    print(f"Error: {e}")
