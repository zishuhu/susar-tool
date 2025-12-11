from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.units import mm
import io
import zipfile
import re
import os
from urllib.parse import quote

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# 配置
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# 注册中文字体
try:
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    FONT_NAME = 'STSong-Light'
except:
    FONT_NAME = 'Helvetica'

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/api/health')
def health():
    return jsonify({'status': 'ok'})

@app.route('/api/process', methods=['POST'])
def process_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': '未上传文件'}), 400
        
        file = request.files['file']
        project_id = request.form.get('project_id', '').strip()
        
        if not project_id:
            return jsonify({'error': '请输入项目编号'}), 400
        
        if not file.filename:
            return jsonify({'error': '文件名为空'}), 400
        
        # 读取 Excel
        wb = openpyxl.load_workbook(file, data_only=True, read_only=True)
        ws = wb.active
        
        # 提取信息
        drug_name = extract_drug_name(ws) or "未知药品"
        date_range = extract_date_range(ws) or "未知日期"
        
        drug_name = re.sub(r'[\\/:*?"<>|]', '_', drug_name)
        date_range = re.sub(r'[\\/:*?"<>|]', '_', date_range)
        
        # 查找列
        project_col, data_start_row = find_project_column(ws)
        if not project_col:
            return jsonify({'error': '未找到项目编号列'}), 400
        
        # 分类
        matching = []
        non_matching = []
        
        for row_idx in range(data_start_row + 1, ws.max_row + 1):
            val = str(ws.cell(row_idx, project_col).value or '').strip()
            if val == project_id:
                matching.append(row_idx)
            elif val:
                non_matching.append(row_idx)
        
        # 生成 PDF
        pdfs = []
        
        if matching:
            buf = create_pdf(ws, data_start_row, matching)
            pdfs.append((f'本项目外院SUSAR_{drug_name}_{date_range}.pdf', buf))
        
        if non_matching:
            buf = create_pdf(ws, data_start_row, non_matching)
            pdfs.append((f'非本项目外院SUSAR_{drug_name}_{date_range}.pdf', buf))
        
        if not pdfs:
            return jsonify({'error': '没有数据'}), 400
        
        # 返回文件
        if len(pdfs) == 1:
            fname, buf = pdfs[0]
            return send_file(buf, mimetype='application/pdf', 
                           as_attachment=True, download_name=fname)
        
        # 打包 ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            for fname, buf in pdfs:
                zf.writestr(fname, buf.getvalue())
        zip_buf.seek(0)
        
        return send_file(zip_buf, mimetype='application/zip',
                        as_attachment=True, 
                        download_name=f'SUSAR_{drug_name}_{date_range}.zip')
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def extract_drug_name(ws):
    for r in range(1, 11):
        for c in range(1, 20):
            v = str(ws.cell(r, c).value or '')
            if 'Investigational Drug' in v or '试验药物' in v:
                m = re.search(r'[:\s：]+(.+)', v)
                if m:
                    return m.group(1).strip()
                if c + 1 <= ws.max_column:
                    nv = str(ws.cell(r, c+1).value or '').strip()
                    if nv:
                        return nv
    return None

def extract_date_range(ws):
    for r in range(1, 11):
        for c in range(1, 20):
            v = str(ws.cell(r, c).value or '')
            if '传输数据区间' in v or 'Data Transfer Period' in v:
                m = re.search(r'[:\s：]+(.+)', v)
                if m:
                    return m.group(1).strip()
                if c + 1 <= ws.max_column:
                    nv = str(ws.cell(r, c+1).value or '').strip()
                    if nv:
                        return nv
    return None

def find_project_column(ws):
    for r in range(1, 11):
        for c in range(1, 30):
            v = str(ws.cell(r, c).value or '').lower()
            if 'study' in v or '项目' in v or '编号' in v:
                return c, r
    return None, None

def create_pdf(ws, start_row, rows):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                          leftMargin=5*mm, rightMargin=5*mm,
                          topMargin=5*mm, bottomMargin=5*mm)
    
    data = []
    for r in range(1, start_row + 1):
        data.append([str(ws.cell(r, c).value or '')[:40] for c in range(1, 30)])
    
    for r in rows:
        data.append([str(ws.cell(r, c).value or '')[:40] for c in range(1, 30)])
    
    t = Table(data)
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), FONT_NAME),
        ('FONTSIZE', (0,0), (-1,-1), 5.5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,start_row-1), colors.lightgrey),
    ]))
    
    doc.build([t])
    buf.seek(0)
    return buf

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
