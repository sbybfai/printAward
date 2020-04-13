import xlrd, os
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document

TEMPLATE_DOC = r"模板.docx"
TEMP_DOC = r"temp.docx"
PARAM_FILE = r'参数.xlsx'


def readParamSheet():
    xls = xlrd.open_workbook(PARAM_FILE)
    sheet1 = xls.sheet_by_name('Sheet1')
    return sheet1

def getTitles(sheet1):
    cols = sheet1.ncols  # 表格列数
    return sheet1.row_values(0, 0, cols)

#将结果合并到一个文件中
def composeDoc(composer, i, rows):
    new_doc = Document(TEMP_DOC)
    if (i < rows - 1):
        new_doc.add_page_break()
    if composer is None:
        composer = Composer(new_doc)
    else:
        composer.append(new_doc)
    return composer

def generalFile():
    sheet1 = readParamSheet()
    titles = getTitles(sheet1)
    print("=======>参数列表：" + str(titles))

    rows = sheet1.nrows  # 表格行数
    cols = sheet1.ncols  # 表格列数
    composer = None
    print("======>根据模板生成结果")
    # 读取数据并生成文件
    for i in range(1, rows):  # 跳过表头一行
        data = {}  # 构造填充模板需要的数据
        for j in range(0, cols):
            val = sheet1.cell_value(i, j)  # 第i行代表第i行数据
            title = titles[j]
            data[title] = val
        doc = DocxTemplate(TEMPLATE_DOC)  # 打开一个模板
        doc.render(data)  # 填充数据data到模板
        doc.save("temp.docx")
        composer = composeDoc(composer, i, rows)

    print(">>>>>>>>>正在保存结果")
    composer.save("结果.docx")
    os.remove("temp.docx")


generalFile()
