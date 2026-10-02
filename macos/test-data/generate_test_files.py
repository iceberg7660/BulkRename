#!/usr/bin/env python
"""生成 BulkRename-mac 测试文件：名单.xlsx + 若干 Word 文件。

用法：
    /opt/miniconda3/envs/iceberg/bin/python generate_test_files.py

说明：
- 名单.xlsx 生成在本脚本同级目录（项目根目录），「使用预设名单」按钮可直接使用
- Word 文件生成在 test-data/待重命名/ 下；重复运行会先清空该目录里的 .docx 再重新生成，
  可用于重命名测试后还原测试数据
- Word 文件为最小 OOXML 结构（无第三方依赖），Word / Pages / WPS 均可正常打开
"""
import os
import zipfile

from openpyxl import Workbook

STUDENTS = [
    (2021001, '张三'),
    (2021002, '李四'),
    (2021003, '王五'),
    (2021004, '赵六'),
    (2021005, '钱七'),
    (2021006, '孙八'),
]

# (文件名, 文档正文)；最后一个文件不在名单中，用于验证程序不会改动无关文件
DOC_FILES = [
    ('2021001张三_第一次作业.docx', '张三（2021001）的第一次作业'),
    ('2021002李四_第一次作业.docx', '李四（2021002）的第一次作业'),
    ('2021003王五_第一次作业.docx', '王五（2021003）的第一次作业'),
    ('2021004赵六_第一次作业.docx', '赵六（2021004）的第一次作业'),
    ('2021005钱七_第一次作业.docx', '钱七（2021005）的第一次作业'),
    ('孙八的作业.docx', '孙八（2021006）的作业'),
    ('郑九_第一次作业.docx', '郑九的第一次作业（此人不在名单中）'),
]

CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)

RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '</Relationships>'
)


def xml_escape(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def make_docx(path, text):
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>'
        f'<w:p><w:r><w:t>{xml_escape(text)}</w:t></w:r></w:p>'
        '<w:sectPr/>'
        '</w:body>'
        '</w:document>'
    )
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES_XML)
        zf.writestr('_rels/.rels', RELS_XML)
        zf.writestr('word/document.xml', document_xml)


def make_roster(path):
    wb = Workbook()
    ws = wb.active
    ws.title = '名单'
    ws.append(['学号', '姓名'])
    for row in STUDENTS:
        ws.append(list(row))
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 10
    wb.save(path)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(here)
    doc_folder = os.path.join(here, '待重命名')

    os.makedirs(doc_folder, exist_ok=True)
    # 还原测试环境：清掉上次生成的（或已被重命名测试改名的）docx
    for name in os.listdir(doc_folder):
        if name.endswith('.docx'):
            os.remove(os.path.join(doc_folder, name))

    roster_path = os.path.join(project_root, '名单.xlsx')
    make_roster(roster_path)
    print(f'已生成名单：{roster_path}（{len(STUDENTS)} 名学生）')

    for filename, text in DOC_FILES:
        make_docx(os.path.join(doc_folder, filename), text)
    print(f'已生成 Word 文件 {len(DOC_FILES)} 个：{doc_folder}')


if __name__ == '__main__':
    main()
