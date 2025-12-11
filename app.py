from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
from openpyxl.utils import get_column_letter
import io
import zipfile
import re
import os
from datetime import datetime

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/api/health')
def health():
    return jsonify({'status': 'ok', 'time': datetime.now().isoformat()})

@app.route('/api/process', methods=['POST'])
def process_file():
    try:
        print("开始处理请求...")
        
        if 'file' not in request.files:
            return jsonify({'error': '未上传文件'}), 400
        
        file = request.files['file']
        project_id = request.form.get('project_id', '').strip()
        
        if not project_id:
            return jsonify({'error': '请输入项目编号'}), 400
        
        print(f"读取文件: {file.filename}")
        
        # 读取 Excel
        wb = openpyxl.load_workbook(file, data_only=True, read_only=True)
        ws = wb.active
        
        print(f"Excel 行数: {ws.max_row}, 列数: {ws.max_column}")
        
        # 提取信息
        drug_name = extract_drug_name(ws) or "未知药品"
        date_range = extract_date_range(ws) or "未知日期"
        
        drug_name = re.sub(r'[\\/:*?"<>|]', '_', drug_name)[:30]
        date_range = re.sub(r'[\\/:*?"<>|]', '_', date_range)[:30]
        
        print(f"药品: {drug_name}, 日期: {date_range}")
        
        # 查找列
        project_col, data_start_row = find_project_column(ws)
        if not project_col:
            return jsonify({'error': '未找到项目编号列'}), 400
        
        print(f"项目列: {project_col}, 起始行: {data_start_row}")
        
        # 分类
        matching = []
        non_matching = []
        
        for row_idx in range(data_start_row + 1, min(ws.max_row + 1, 1000)):  # 限制最多1000行
            val = str(ws.cell(row_idx, project_col).value or '').strip()
            if val == project_id:
                matching.append(row_idx)
            elif val:
                non_matching.append(row_idx)
        
        print(f"本项目: {len(matching)}, 非本项目: {len(non_matching)}")
        
        # 生成 Excel 文件（不生成 PDF，避免内存问题）
        excel_files = []
        
        if matching:
            print("生成本项目 Excel...")
            buf = create_excel(wb, ws, data_start_row, matching)
            excel_files.append((f'本项目外院SUSAR_{drug_name}_{date_range}.xlsx', buf))
        
        if non_matching:
            print("生成非本项目 Excel...")
            buf = create_excel(wb, ws, data_start_row, non_matching)
            excel_files.append((f'非本项目外院SUSAR_{drug_name}_{date_range}.xlsx', buf))
        
        if not excel_files:
            return jsonify({'error': '没有数据'}), 400
        
        # 返回单个文件
        if len(excel_files) == 1:
            fname, buf = excel_files[0]
            print(f"返回文件: {fname}")
            return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           as_attachment=True, download_name=fname)
        
        # 打包 ZIP
        print("打包 ZIP...")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fname, buf in excel_files:
                zf.writestr(fname, buf.getvalue())
        zip_buf.seek(0)
        
        zip_name = f'SUSAR_{drug_name}_{date_range}.zip'
        print(f"返回 ZIP: {zip_name}")
        
        return send_file(zip_buf, mimetype='application/zip',
                        as_attachment=True, download_name=zip_name)
        
    except Exception as e:
        print(f"错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

def create_excel(original_wb, original_ws, data_start_row, keep_rows):
    """创建筛选后的 Excel（保留格式）"""
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    new_ws.title = original_ws.title
    
    # 复制列宽
    for col_idx in range(1, min(original_ws.max_column + 1, 30)):
        col_letter = get_column_letter(col_idx)
        if original_ws.column_dimensions[col_letter].width:
            new_ws.column_dimensions[col_letter].width = original_ws.column_dimensions[col_letter].width
    
    # 复制头部
    for row_idx in range(1, data_start_row + 1):
        for col_idx in range(1, min(original_ws.max_column + 1, 30)):
            src_cell = original_ws.cell(row_idx, col_idx)
            dst_cell = new_ws.cell(row_idx, col_idx)
            
            dst_cell.value = src_cell.value
            if src_cell.font:
                dst_cell.font = src_cell.font.copy()
            if src_cell.border:
                dst_cell.border = src_cell.border.copy()
            if src_cell.fill:
                dst_cell.fill = src_cell.fill.copy()
            if src_cell.alignment:
                dst_cell.alignment = src_cell.alignment.copy()
    
    # 复制数据行
    new_row_idx = data_start_row + 1
    for orig_row_idx in keep_rows:
        for col_idx in range(1, min(original_ws.max_column + 1, 30)):
            src_cell = original_ws.cell(orig_row_idx, col_idx)
            dst_cell = new_ws.cell(new_row_idx, col_idx)
            
            dst_cell.value = src_cell.value
            if src_cell.font:
                dst_cell.font = src_cell.font.copy()
            if src_cell.border:
                dst_cell.border = src_cell.border.copy()
            if src_cell.fill:
                dst_cell.fill = src_cell.fill.copy()
            if src_cell.alignment:
                dst_cell.alignment = src_cell.alignment.copy()
        
        new_row_idx += 1
    
    # 保存到内存
    buf = io.BytesIO()
    new_wb.save(buf)
    buf.seek(0)
    return buf

def extract_drug_name(ws):
    for r in range(1, min(11, ws.max_row + 1)):
        for c in range(1, min(20, ws.max_column + 1)):
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
    for r in range(1, min(11, ws.max_row + 1)):
        for c in range(1, min(20, ws.max_column + 1)):
            v = str(ws.cell(r, c).value or '')
            if '传输数据区间' in v or 'Data Transfer Period' in v or '数据区间' in v:
                m = re.search(r'[:\s：]+(.+)', v)
                if m:
                    return m.group(1).strip()
                if c + 1 <= ws.max_column:
                    nv = str(ws.cell(r, c+1).value or '').strip()
                    if nv:
                        return nv
    return None

def find_project_column(ws):
    for r in range(1, min(11, ws.max_row + 1)):
        for c in range(1, min(30, ws.max_column + 1)):
            v = str(ws.cell(r, c).value or '').lower()
            if 'study' in v or '项目' in v or '编号' in v or 'protocol' in v:
                return c, r
    return None, None

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
