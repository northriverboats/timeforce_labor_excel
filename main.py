"""
    Oh yes, the classic "Hello World". The problem is that you
    can do it so many ways using PySimpleGUI
"""

from pathlib import Path
import PySimpleGUI as sg
import sys
from decimal import Decimal
from excelopen import ExcelOpenDocument
from xlsxwriter import Workbook # type: ignore

TITLES = [
    ["Employee Name", 24],
    ["Job Name", 10],
    ["Task Name", 23],
    ['Total Hours', 12],
    ['Fabrication', 12],
    ['Paint', 12],
    ['Canvas', 12],
    ['Floorboards', 12],
    ['Outfitting', 12],
]

TASKS = {
    "1 Boat Builder": "Fabrication",
    "4 Paint": "Paint",
    "2 Canvas and Upholstery": "Canvas",
    "5 Outfitting - Floorboard": "Floorboard",
    "5 Outfitting": "Outfitting",
}

TASK_COLUMN = {
    "Fabrication": 4,
    "Paint": 5,
    "Canvas": 6,
    "Floorboard": 7,
    "Outfitting": 8,
}

DEPTS = [
    "Fabrication",
    "Paint",
    "Canvas",
    "Outfitting",
    "Floorboard",
]

def is_excel(file_name):
    """convert filename to path if valid excel file"""
    excel_file = None
    path = Path(file_name)
    if path.is_file() and path.suffix == '.xlsx':
        excel_file = path
    return excel_file

def dept_ref(dept, row):
    """convert deept/row to cell address"""
    return chr(TASK_COLUMN[dept] +65) + str(row + 1)

def read_sheet(original):
    """read spreadsheet"""
    xlsx = ExcelOpenDocument()
    xlsx.open(original)
    labor = {}
    for employee, _, job, task, _, _, hours in xlsx.rows(min_row=2, min_col=3, max_col=9):
        if task.value not in TASKS:
            continue
        if job.value not in labor:
            labor[job.value] = {}
        if task.value not in labor[job.value]:
            labor[job.value][task.value] = {}
        labor[job.value][task.value][employee.value] = Decimal(hours.value)
    xlsx.close()
    return labor

def write_headers(formats, xlsx):
    """write spreadsheet headrs"""
    for column, title in enumerate(TITLES):
        col = chr(column + 65)
        xlsx.set_column(col + ":" + col, title[1])
        xlsx.write(0, column, title[0], formats['bold'])

def write_task(formats, xlsx, employees, boat, task_name, task, row):
    """write out task of one boat to spreadsheet"""
    hours = Decimal(0.0)
    start_row = row + 1
    for employee in employees:
        hour = employees[employee]
        hours += hour
        xlsx.write(row, 0, employee)
        xlsx.write(row, 1, boat)
        xlsx.write(row, 2, task_name)
        xlsx.write(row, 3, hour, formats['decimal'])
        end_row = row
        row +=1
    text = f"=SUM(D{start_row}:D{end_row + 1})"
    xlsx.write(end_row, TASK_COLUMN[task], text, formats['decimal'])
    return row

def write_boat(formats, xlsx, boat, boat_name, row):
    """write out one boat to spreadsheet"""
    for task in boat:
        row = write_task(formats, xlsx, boat[task], boat_name, task, TASKS[task], row)
    return row

def write_totals(formats, xlsx, row):
    """write out totals on sheet"""
    row -= 1
    for task in DEPTS:
        text = f"=SUM({dept_ref(task, 1)}:{dept_ref(task, row-1)})"
        xlsx.write(row, TASK_COLUMN[task], text, formats['totals'])
    xlsx.write(row, 3, f"=SUM(D2:D{row})", formats['totals'])
    xlsx.freeze_panes(1, 0, row-20, 1)
    xlsx.set_selection(row, 2, row, 2)


def write_boats(formats, xlsx, labor):
    """write out all boats"""
    row = 1
    for boat in labor:
        row = write_boat(formats, xlsx, labor[boat], boat, row)
        row += 1
    write_totals(formats, xlsx, row)

def write_sheet(file_path, labor):
    """write spreadsheet"""
    with Workbook(file_path) as workbook:
        xlsx = workbook.add_worksheet('Labor')
        formats = {}
        formats['bold'] = workbook.add_format({'bold': True})
        formats['decimal'] = workbook.add_format({'num_format': '#,##0.00'})
        formats['totals'] = workbook.add_format({'bold': True, 'num_format': '#,##0.00'})
        formats['totals'].set_top(6)
        write_headers(formats, xlsx)
        write_boats(formats, xlsx, labor)

def gui(excel_file):
    """build/show gui and handle event loop"""
    layout = [
        [sg.Text('Spreadsheet to process:')],
        [sg.Input('', key='-TEXT-', readonly=True)],
    ]
    window = sg.Window('Format TimeForce Labor Report', layout, finalize=True)

    timeout = thread = None
    if excel_file:
        window.write_event_value('-READSHEET-', True)
    else:
        window.write_event_value('-OPENFILE-', True)
    # --------------------- EVENT LOOP ---------------------
    while True:
        event, values = window.read(timeout=timeout)
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        elif event == '-READSHEET-':
            window['-TEXT-'].update(excel_file.name)
        elif event == '-OPENFILE-':
            file_name = sg.popup_get_file(
                message="Select Labor Spreadsheet",
                file_types=(('Excel File','*.xlsx'),),
                no_window=True)
            excel_file = is_excel(file_name)
            if excel_file:
                window.write_event_value('-READSHEET-', True)
            else:
                sg.Popup('No spreadsheet to process..........', title='Closing Program', keep_on_top=True)
                break

    window.close()

def main():
    """main function show gui"""
    args = sys.argv
    if len(args) == 2:
        excel_file = is_excel(args[1])
    gui(excel_file)
    sys.exit(0)

    """
    original = Path("May 2022.xlsx")
    file_path = "test.xlsx"
    labor = read_sheet(original)
    write_sheet(file_path, labor)
    """

        




if __name__ == "__main__":
    main()  # pylint: disable=no-value-for-parameter