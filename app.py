from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet
import io
import zipfile
import re

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# 尝试注册中文字体
try:
    pdfmetrics.registerFont(TTFont('SimSun', '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'))
    FONT_NAME = 'SimSun'
except:
    try:
        pdfmetrics.registerFont(TTFont('SimSun', 'C:\\Windows\\Fonts\\simsun.ttc'))
        FONT_NAME = 'SimSun'
    except:
        FONT_NAME = 'Helvetica'

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/api/process', methods=['POST'])
def process_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': '未上传文件'}), 400
        
        file = request.files['file']
        project_id = request.form.get('project_id', '').strip()
        
        if not project_id:
            return jsonify({'error': '请输入项目编号'}), 400
        
        if file.filename == '':
            return jsonify({'error': '文件名为空'}), 400
        
        # 读取 Excel 文件
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        
        # 提取元数据
        drug_name = extract_drug_name(ws)
        date_range = extract_date_range(ws)
        
        if not drug_name:
            drug_name = "未知药品"
        if not date_range:
            date_range = "未知日期"
        
        # 清理文件名
        drug_name = re.sub(r'[\\/:*?"<>|]', '_', drug_name)
        date_range = re.sub(r'[\\/:*?"<>|]', '_', date_range)
        
        # 查找项目编号列
        project_col, data_start_row = find_project_column(ws)
        
        if project_col is None:
            return jsonify({'error': '未找到项目编号列'}), 400
        
        # 分类数据
        matching_rows = []
        non_matching_rows = []
        
        for row_idx in range(data_start_row + 1, ws.max_row + 1):
            cell_value = str(ws.cell(row_idx, project_col).value or '').strip()
            if cell_value == project_id:
                matching_rows.append(row_idx)
            elif cell_value:
                non_matching_rows.append(row_idx)
        
        # 生成 PDF 文件
        pdf_files = []
        
        if matching_rows:
            pdf_buffer = create_pdf(ws, data_start_row, matching_rows)
            pdf_files.append((f'本项目外院SUSAR_{drug_name}_{date_range}.pdf', pdf_buffer))
        
        if non_matching_rows:
            pdf_buffer = create_pdf(ws, data_start_row, non_matching_rows)
            pdf_files.append((f'非本项目外院SUSAR_{drug_name}_{date_range}.pdf', pdf_buffer))
        
        if not pdf_files:
            return jsonify({'error': '没有找到任何数据'}), 400
        
        # 如果只有一个文件，直接返回
        if len(pdf_files) == 1:
            filename, buffer = pdf_files[0]
            return send_file(
                buffer,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=filename
            )
        
        # 多个文件打包成 ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for filename, buffer in pdf_files:
                zip_file.writestr(filename, buffer.getvalue())
        
        zip_buffer.seek(0)
        
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'SUSAR_{drug_name}_{date_range}.zip'
        )
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({'error': f'处理失败: {str(e)}'}), 500

def extract_drug_name(ws):
    """提取药品名称"""
    for row in range(1, min(11, ws.max_row + 1)):
        for col in range(1, min(ws.max_column + 1, 20)):
            cell_value = str(ws.cell(row, col).value or '')
            if 'Investigational Drug' in cell_value or '试验药物' in cell_value:
                match = re.search(r'(?:Investigational Drug|试验药物)[:\s：]+(.+)', cell_value)
                if match:
                    return match.group(1).strip()
                if col + 1 <= ws.max_column:
                    next_val = str(ws.cell(row, col + 1).value or '').strip()
                    if next_val:
                        return next_val
    return None

def extract_date_range(ws):
    """提取日期区间"""
    for row in range(1, min(11, ws.max_row + 1)):
        for col in range(1, min(ws.max_column + 1, 20)):
            cell_value = str(ws.cell(row, col).value or '')
            if any(keyword in cell_value for keyword in ['传输数据区间', 'Data Transfer Period', '数据区间']):
                match = re.search(r'[:\s：]+(.+)', cell_value)
                if match:
                    return match.group(1).strip()
                if col + 1 <= ws.max_column:
                    next_val = str(ws.cell(row, col + 1).value or '').strip()
                    if next_val:
                        return next_val
    return None

def find_project_column(ws):
    """查找项目编号列"""
    for row in range(1, min(11, ws.max_row + 1)):
        for col in range(1, min(ws.max_column + 1, 30)):
            cell_value = str(ws.cell(row, col).value or '').lower()
            if any(keyword in cell_value for keyword in ['study', '项目', 'study id', 'protocol', '编号']):
                return col, row
    return None, None

def create_pdf(ws, data_start_row, keep_rows):
    """创建 PDF"""
    buffer = io.BytesIO()
    
    # 创建 PDF 文档（横向 A4）
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=8*mm,
        rightMargin=8*mm,
        topMargin=8*mm,
        bottomMargin=8*mm
    )
    
    # 准备表格数据
    table_data = []
    
    # 添加头部行
    for row in range(1, data_start_row + 1):
        row_data = []
        for col in range(1, min(ws.max_column + 1, 30)):  # 限制最多30列
            cell = ws.cell(row, col)
            value = str(cell.value or '')
            # 限制单元格长度
            if len(value) > 50:
                value = value[:47] + '...'
            row_data.append(value)
        table_data.append(row_data)
    
    # 添加数据行
    for row_idx in keep_rows:
        row_data = []
        for col in range(1, min(ws.max_column + 1, 30)):
            cell = ws.cell(row_idx, col)
            value = str(cell.value or '')
            if len(value) > 50:
                value = value[:47] + '...'
            row_data.append(value)
        table_data.append(row_data)
    
    # 创建表格
    table = Table(table_data)
    
    # 设置表格样式
    style = TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, -1), 6),
        ('FONTSIZE', (0, 0), (-1, data_start_row - 1), 8),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, data_start_row - 1), colors.lightgrey),
    ])
    
    table.setStyle(style)
    
    # 构建 PDF
    elements = [table]
    doc.build(elements)
    
    buffer.seek(0)
    return buffer

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)