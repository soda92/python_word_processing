from docx import Document

doc = Document('订单数据.xlsx')
table = doc.tables[0]

table.cell(0,0)