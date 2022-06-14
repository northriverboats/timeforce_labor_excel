"""
    Oh yes, the classic "Hello World". The problem is that you
    can do it so many ways using PySimpleGUI
"""

from pathlib import Path
import PySimpleGUI as sg
import sys
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
    "2 Canvas and Upholstery": "Canvas",
    "4 Paint": "Paint",
    "5 Outfitting - Floorboard": "Floorboard"
    "5 Outfitting": "Outfitting",
}

def read_sheet(original):
    """read spreadsheet"""
    pass

def write_sheet(file_path):
    """write spreadsheet"""
    xlsx = ExcelOpenDocument()
    xlsx.new(file_path)
    bold = xlsx.font(bold=True)
    for column, title in enumerate(TITLES, start=1):
        xlsx.set_width(chr(column + 64), title[1])
        cell = xlsx.cell(column=column, row=1)
        cell.font = bold
        cell.value = title[0]
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
    write_sheet(file_path)
    # gui(original)



if __name__ == "__main__":
    main()  # pylint: disable=no-value-for-parameter