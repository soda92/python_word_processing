from openpyxl import load_workbook
from docx import Document
from pathlib import Path
import os


def get_output_path(filename):
    curr_dir = Path(__file__).resolve().parent
    out_dir = Path.joinpath(curr_dir, 'out')
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)
    out_file_path = Path.joinpath(out_dir, filename)
    return out_file_path


def get_in_path(filename):
    curr_dir = Path(__file__).resolve().parent
    in_file_path = Path.joinpath(curr_dir, filename)
    return in_file_path


def read_data(filename):
    mapping = {}
    wb = load_workbook(filename=get_in_path(filename))
    ws = wb.active

    for row in list(range(1, 20)):
        for col in list(range(1, 20)):
            val = ws.cell(row=row, column=col).value
            if val != None:
                cell_right = ws.cell(row=row, column=col+1)
                crval = cell_right.value
                if crval != None:
                    mapping[val] = crval
    wb.close()
    return mapping


def modify_xlsx(mapping, filename):
    wb = load_workbook(filename=get_in_path(filename))
    ws = wb.active

    for row in list(range(1, 20)):
        for col in list(range(1, 20)):
            val = ws.cell(row=row, column=col).value
            cell_right = ws.cell(row=row, column=col+1)
            crval = cell_right.value
            if crval == None:
                if val in mapping:
                    print(f"写入数据：{val}->{mapping[val]}")
                    ws.cell(row=row, column=col+1, value=mapping[val])
    wb.save(get_output_path(filename))


def modify_docx(mapping, filename):
    doc = Document(docx=get_in_path(filename))
    table = doc.tables[0]

    for i in range(len(table.rows)):
        for j in range(len(table.rows[i].cells)):
            text = table.cell(i, j).text
            if text in mapping:
                print(f"写入数据：{text}->{mapping[text]}")
                table.cell(i, j+1).text = str(mapping[text])
    doc.save(get_output_path(filename))


data = read_data("订单数据.xlsx")
modify_xlsx(data, "订单信息.xlsx")
modify_docx(data, "订单信息.docx")
