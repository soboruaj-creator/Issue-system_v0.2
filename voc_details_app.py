from flask import Flask, render_template, request, jsonify, send_file, redirect
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['SECRET_KEY'] = 'voc-details-secret-key'
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

# 업로드 폴더 생성
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/voc/model/<model_name>')
def show_model_vocs(model_name):
    """모델별 VOC 목록 페이지"""
    conn = sqlite3.connect('voc_data.db')
    
    # 모델별 VOC 조회
    query = """
        SELECT case_code, model_name, cause, solution, created_date, title, problem
        FROM internal_voc
        WHERE model_name = ?
        ORDER BY created_date DESC
    """
    df = pd.read_sql_query(query, conn, params=(model_name,))
    
    conn.close()
    
    return render_template('voc_model_list.html', 
                          vocs=df.to_dict('records'), 
                          model_name=model_name)

@app.route('/voc/monthly/<month>')
def show_monthly_vocs(month):
    """월별 VOC 목록 페이지"""
    conn = sqlite3.connect('voc_data.db')
    
    # 월별 모델별 VOC 건수 조회
    query = """
        SELECT model_name, COUNT(*) as count
        FROM internal_voc
        WHERE strftime('%Y-%m', created_date) = ?
        GROUP BY model_name
        ORDER BY count DESC
    """
    df = pd.read_sql_query(query, conn, params=(month,))
    
    # 3달 연속 상위 5개 확인
    top5_models = df.head(5)['model_name'].tolist()
    
    # 이전 2달 데이터 조회
    previous_months = []
    current_date = datetime.strptime(month, '%Y-%m')
    for i in range(1, 3):
        prev_date = current_date - timedelta(days=30*i)
        prev_month = prev_date.strftime('%Y-%m')
        previous_months.append(prev_month)
    
    # 이전 2달의 상위 5개 모델 조회
    consecutive_top5 = []
    for prev_month in previous_months:
        query = """
            SELECT model_name, COUNT(*) as count
            FROM internal_voc
            WHERE strftime('%Y-%m', created_date) = ?
            GROUP BY model_name
            ORDER BY count DESC
            LIMIT 5
        """
        df_prev = pd.read_sql_query(query, conn, params=(prev_month,))
        prev_top5 = df_prev['model_name'].tolist()
        consecutive_top5.append(prev_top5)
    
    # 3달 연속 상위 5개 모델 확인
    consecutive_models = set(top5_models)
    for prev_top5 in consecutive_top5:
        consecutive_models = consecutive_models.intersection(set(prev_top5))
    
    # 전달대비 증가폭 계산
    prev_month = previous_months[0]
    query_prev = """
        SELECT model_name, COUNT(*) as count
        FROM internal_voc
        WHERE strftime('%Y-%m', created_date) = ?
        GROUP BY model_name
    """
    df_prev = pd.read_sql_query(query_prev, conn, params=(prev_month,))
    
    # 증가율 계산
    df['growth_rate'] = df.apply(lambda row: calculate_growth_rate(row['model_name'], row['count'], df_prev), axis=1)
    
    # 모델별 비율 계산
    total_count = df['count'].sum()
    df['percentage'] = df.apply(lambda row: round((row['count'] / total_count) * 100, 1) if total_count > 0 else 0.0, axis=1)
    
    # 월별 VOC 조회 (모델별로 그룹화)
    query = """
        SELECT case_code, model_name, cause, solution, created_date, title, problem
        FROM internal_voc
        WHERE strftime('%Y-%m', created_date) = ?
        ORDER BY model_name, created_date DESC
    """
    df_vocs = pd.read_sql_query(query, conn, params=(month,))
    
    conn.close()
    
    return render_template('voc_monthly_list.html',
                          vocs=df_vocs.to_dict('records'),
                          model_stats=df.to_dict('records'),
                          month=month,
                          consecutive_models=list(consecutive_models))

def calculate_growth_rate(model_name, current_count, df_prev):
    """전달대비 증가율 계산"""
    prev_row = df_prev[df_prev['model_name'] == model_name]
    if len(prev_row) > 0:
        prev_count = prev_row.iloc[0]['count']
        if prev_count > 0:
            growth_rate = ((current_count - prev_count) / prev_count) * 100
            return round(growth_rate, 1)
    return None

@app.route('/api/voc/model/<model_name>/export')
def export_model_vocs(model_name):
    """모델별 VOC 엑셀 다운로드"""
    try:
        conn = sqlite3.connect('voc_data.db')
        
        query = """
            SELECT case_code, model_name, cause, solution, created_date, title, problem
            FROM internal_voc
            WHERE model_name = ?
            ORDER BY created_date DESC
        """
        df = pd.read_sql_query(query, conn, params=(model_name,))
        
        conn.close()
        
        # 엑셀 파일 생성
        filename = f"voc_{model_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        df.to_excel(filepath, index=False)
        
        # 파일 존재 여부 확인
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': '파일 생성 실패'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/voc/monthly/<month>/export')
def export_monthly_vocs(month):
    """월별 VOC 엑셀 다운로드"""
    try:
        conn = sqlite3.connect('voc_data.db')
        
        query = """
            SELECT case_code, model_name, cause, solution, created_date, title, problem
            FROM internal_voc
            WHERE strftime('%Y-%m', created_date) = ?
            ORDER BY model_name, created_date DESC
        """
        df = pd.read_sql_query(query, conn, params=(month,))
        
        conn.close()
        
        # 엑셀 파일 생성
        filename = f"voc_monthly_{month}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        df.to_excel(filepath, index=False)
        
        # 파일 존재 여부 확인
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': '파일 생성 실패'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
