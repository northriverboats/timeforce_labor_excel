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

def read_sheet(original):
    """read spreadsheet"""
    xlsx = ExcelOpenDocument()
    xlsx.open(original)
    labor = {'total': Decimal(0.0)}
    # for employee, _, job, task, _, _, hours in xlsx.rows(min_row=2, max_row=32, min_col=3, max_col=9):
    for employee, _, job, task, _, _, hours in xlsx.rows(min_row=2, min_col=3, max_col=9):
        if task.value not in TASKS:
            continue
        if job.value not in labor:
            labor[job.value] = {'total': Decimal(0.0)}
        if task.value not in labor[job.value]:
            labor[job.value][task.value] = {'total': Decimal(0.0)}
        labor[job.value][task.value][employee.value] = hours.value
        labor[job.value][task.value]['total'] += Decimal(hours.value)
        labor[job.value]['total'] += Decimal(hours.value)
        labor['total'] += Decimal(hours.value)
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

def write_task(xlsx, employees, boat, task_name, task, totals):
    """write out task of one boat to spreadsheet"""
    hours = Decimal(0.0)
    total = Decimal(0.0)
    for employee in employees:
        hour = employees[employee]
        if employee == 'total':
            total = employees[employee]
            continue
        hours += Decimal(hour)
        totals['Total'] += Decimal(hour)
        totals['Subtotal'] += Decimal(hour)
        totals[task] += Decimal(hour)
        xlsx.cell(row=totals['Row'], column=1).value = employee
        xlsx.cell(row=totals['Row'], column=2).value = boat
        xlsx.cell(row=totals['Row'], column=3).value = task_name
        xlsx.cell(row=totals['Row'], column=4).value = hour
        xlsx.cell(row=totals['Row'], column=4).number_format = "#,##0.00"
        totals['Row'] += 1
        if hours == total:
            print(f"        {employee:24.24} {hour:8.2f}  {hours:8.2f}")
        else:
            print(f"        {employee:24.24} {hour:8.2f}")
    xlsx.cell(row=totals['Row']-1, column=TASK_COLUMN[task]).value = totals['Subtotal']
    xlsx.cell(row=totals['Row']-1, column=TASK_COLUMN[task]).number_format = "#,##0.00"


def write_boat(xlsx, boat, boat_name, totals):
    """write out one boat to spreadsheet"""
    totals['Subtotal'] = Decimal(0.0)
    for task in boat:
        if task == 'total':
            continue
        print(f"    {task}")
        write_task(xlsx, boat[task], boat_name, task, TASKS[task], totals)
    print(f"                                                    {totals['Subtotal']:8.2f}")

def write_totals(xlsx, totals):
    """write out totals on sheet"""
    for task in ["Fabrication", "Paint", "Canvas", "Outfitting", "Floorboard"]:
        text = "=SUM(" + chr(TASK_COLUMN[task] +64) + "2:" + chr(TASK_COLUMN[task] +64) + str(totals['Row']-2) + ")"
        xlsx.cell(row=totals['Row']-1, column=TASK_COLUMN[task]).value =  text # totals['Subtotal']
        xlsx.cell(row=totals['Row']-1, column=TASK_COLUMN[task]).number_format = "#,##0.00"
    print(f"     {totals['Total']:8.2f}    {totals['Fabrication']:8.2f}    {totals['Paint']:8.2f}    {totals['Canvas']:8.2f}    {totals['Floorboard']:8.2f}    {totals['Outfitting']:8.2f}")

def write_boats(xlsx, labor):
    """write out all boats"""
    grand_total = Decimal(0.0)
    totals = {
        'Total': Decimal(0.0),
        'Fabrication': Decimal(0.0),
        'Canvas': Decimal(0.0),
        'Paint': Decimal(0.0),
        'Floorboard': Decimal(0.0),
        'Outfitting': Decimal(0.0),
        'Subtotal': Decimal(0.0),
        'Row': Decimal(2.0),
    }
    for boat in labor:
        if boat == 'total':
            total = labor[boat]
            continue
        print(f"{boat}")
        write_boat(xlsx, labor[boat], boat, totals)
        totals['Row'] += 1
    write_totals(xlsx, totals)

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