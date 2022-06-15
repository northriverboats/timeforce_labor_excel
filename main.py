"""
    Oh yes, the classic "Hello World". The problem is that you
    can do it so many ways using PySimpleGUI
"""

from pathlib import Path
import PySimpleGUI as sg
import sys
from decimal import Decimal
from excelopen import ExcelOpenDocument


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
    "Fabrication": 5,
    "Paint": 6,
    "Canvas": 7,
    "Floorboard": 8,
    "Outfitting": 9,
}

DEPTS = [
    "Fabrication",
    "Paint",
    "Canvas",
    "Outfitting",
    "Floorboard",
]

def dept_ref(dept, row):
    """convert deept/row to cell address"""
    return chr(TASK_COLUMN[dept] +64) + str(row)

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
    return labor

def write_headers(xlsx):
    """write spreadsheet headrs"""
    bold = xlsx.font(bold=True)
    xlsx.freeze_panes("A2")
    for column, title in enumerate(TITLES, start=1):
        xlsx.set_width(chr(column + 64), title[1])
        cell = xlsx.cell(column=column, row=1)
        cell.font = bold
        cell.value = title[0]

def write_task(xlsx, employees, boat, task_name, task, row):
    """write out task of one boat to spreadsheet"""
    hours = Decimal(0.0)
    start_row = row
    for employee in employees:
        hour = employees[employee]
        hours += hour
        xlsx.cell(row=row, column=1).value = employee
        xlsx.cell(row=row, column=2).value = boat
        xlsx.cell(row=row, column=3).value = task_name
        xlsx.cell(row=row, column=4).value = hour
        xlsx.cell(row=row, column=4).number_format = "#,##0.00"
        end_row = row
        row +=1
    text = f"=SUM(D{start_row}:D{end_row})"
    xlsx.cell(row=end_row, column=TASK_COLUMN[task]).value = text
    xlsx.cell(row=end_row, column=TASK_COLUMN[task]).number_format = "#,##0.00"
    return row

def write_boat(xlsx, boat, boat_name, row):
    """write out one boat to spreadsheet"""
    for task in boat:
        row = write_task(xlsx, boat[task], boat_name, task, TASKS[task], row)
    return row

def write_totals(xlsx, row):
    """write out totals on sheet"""
    bold = xlsx.font(bold=True)
    row -= 1
    for task in DEPTS:
        text = f"=SUM({dept_ref(task, 2)}:{dept_ref(task, row-1)})"
        xlsx.cell(row=row, column=TASK_COLUMN[task]).value =  text
        xlsx.cell(row=row, column=TASK_COLUMN[task]).font =  bold
        xlsx.cell(row=row, column=TASK_COLUMN[task]).number_format = "#,##0.00"
    xlsx.cell(row=row, column=4).value = f"=SUM(D2:D{row-1})"
    xlsx.cell(row=row, column=4).font = bold
    xlsx.cell(row=row, column=4).number_format = "#,##0.00"
    # xlsx.set_active_cell(f"A{row}")


def write_boats(xlsx, labor):
    """write out all boats"""
    row = 2
    for boat in labor:
        if boat == 'total':
            total = labor[boat]
            continue
        row = write_boat(xlsx, labor[boat], boat, row)
        row += 1
    write_totals(xlsx, row)

def write_sheet(file_path, labor):
    """write spreadsheet"""
    xlsx = ExcelOpenDocument()
    xlsx.new(file_path)
    write_headers(xlsx)
    write_boats(xlsx, labor)
    xlsx.save()

def gui(original):
    """build/show gui and handle event loop"""
    layout = [
        [sg.Text('Spreadsheet to process')],
        [sg.Text(original.name)],
        [sg.Push(), sg.Button('Exit')],
    ]
    window = sg.Window('Format TimeForce Labor Report', layout, finalize=True)

    timeout = thread = None
    window.write_event_value('-READSHEET-', True)
    # --------------------- EVENT LOOP ---------------------
    while True:
        event, values = window.read(timeout=timeout)
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        elif event == '-READSHEET-':
            print("This is the dog's bollocks mon")
    window.close()

def main():
    """main function show gui"""
    original = Path("May 2022.xlsx")
    file_path = "test.xlsx"
    labor = read_sheet(original)
    write_sheet(file_path, labor)
    # gui(original)



if __name__ == "__main__":
    main()  # pylint: disable=no-value-for-parameter