from flask import Flask, render_template, request, jsonify, redirect, url_for, session, send_file, abort
import pandas as pd
import os
import pdfkit
import yagmail
from datetime import datetime, timedelta, timezone
from hashids import Hashids

hashids = Hashids(salt="saedam_secret_salt", min_length=8)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'saedam_secure_key_2026')

# --- [저장 경로 설정: /mnt/data 디스크 적용] ---
if os.path.exists('/mnt/data'):
    MOUNT_PATH = '/mnt/data'
else:
    MOUNT_PATH = os.getcwd()

EXCEL_FILE = os.path.join(MOUNT_PATH, 'admin_list.xlsx')
CONTRACTS_DIR = os.path.join(MOUNT_PATH, 'contracts')
TERMS_DIR = os.path.join(os.getcwd(), 'terms') 

if not os.path.exists(CONTRACTS_DIR):
    os.makedirs(CONTRACTS_DIR)

# [설정] 리눅스 표준 wkhtmltopdf 경로
WKHTMLTOPDF_PATH = '/usr/bin/wkhtmltopdf'
PDF_CONFIG = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

SENDER_EMAIL = os.environ.get('MAIL_USERNAME')
SENDER_PASSWORD = os.environ.get('MAIL_PASSWORD')
ADMIN_PASSWORD = 'school97$$'
KST = timezone(timedelta(hours=9))

def init_excel():
    """엑셀 초기화 (모든 읽기 작업에 dtype=str 적용)"""
    columns = [
        '계약구분', '수탁학교명', '부서명', '성명', '주민번호', '수수료', '보조금', '경력수당', '직책수당', '기타', '근무시간', '계약기간', 'email', '연락처', '거주지', '계약완료일시', '연도', '파일명', 'IP'
    ]
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=columns)
        df.to_excel(EXCEL_FILE, index=False)
    else:
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        df = df.reindex(columns=columns)
        df.to_excel(EXCEL_FILE, index=False)

init_excel()

# --- [강사 서비스 로직] ---
@app.route('/')
def home():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        name = request.form['name']
        ssn = request.form['ssn'].replace("-", "")
        ssn_last4 = request.form['ssn_last4']
        try:
            df = pd.read_excel(EXCEL_FILE, dtype=str)
            user_rows = df[(df['성명'] == name) & (df['주민번호'].astype(str).str.replace("-", "") == ssn)]
            if not user_rows.empty and ssn[-4:] == ssn_last4:
                session['user_name'] = name
                session['user_ssn'] = request.form['ssn']
                return redirect(url_for('contract_list'))
            return "<script>alert('정보가 일치하지 않습니다.'); history.back();</script>"
        except Exception as e:
            return f"에러: {str(e)}"
    return render_template('login.html')

@app.route('/list')
def contract_list():
    if 'user_name' not in session: return redirect(url_for('login'))
    df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
    my_contracts_df = df[(df['성명'] == session['user_name']) & (df['주민번호'].astype(str) == session['user_ssn']) & (df['계약완료일시'] == "")]
    contracts = []
    for idx, row in my_contracts_df.iterrows():
        item = row.to_dict()
        item['safe_id'] = hashids.encode(idx) 
        contracts.append(item)
    return render_template('list.html', contracts=contracts, name=session['user_name'])

@app.route('/contract/<string:safe_id>')
def contract(safe_id):
    if 'user_name' not in session or 'user_ssn' not in session: return redirect(url_for('login'))
    decoded = hashids.decode(safe_id)
    if not decoded: return abort(404)
    orig_idx = decoded[0]
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        if orig_idx >= len(df): return abort(404)
        target_row = df.iloc[orig_idx]
        if target_row['성명'] != session.get('user_name'):
            return "<script>alert('해당 계약서에 대한 접근 권한이 없습니다.'); location.href='/list';</script>"
        user_data = target_row.to_dict()
        user_data['orig_idx'] = orig_idx
        contract_type = user_data.get('계약구분', '방과후강사')
        
        def load_and_replace(filename):
            path = os.path.join(TERMS_DIR, filename)
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    c = f.read()
                    for col in df.columns:
                        val = str(user_data.get(col, '')) if pd.notna(user_data.get(col)) else ""
                        c = c.replace(f"{{{{ data.{col} }}}}", val)
                        display_style = "display:none" if not val or val == '0' or str(val).strip() == '' else "display:table-row"
                        c = c.replace(f"{{{{ style.{col} }}}}", display_style)
                    return c.replace('\n', '<br>').replace('<br><table', '<table').replace('</table><br>', '</table>')
            return ""
            
        user_data['terms_content1'] = load_and_replace(f"{contract_type}.txt")
        user_data['terms_content2'] = load_and_replace(f"{contract_type}2.txt")
        return render_template('contract.html', data=user_data)
    except Exception as e: return f"에러 발생: {str(e)}", 500

@app.route('/save_contract', methods=['POST'])
def save_contract():
    data = request.json
    idx = int(data['orig_idx'])
    now_dt = datetime.now(KST)
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        if (str(df.at[idx, '성명']) != session.get('user_name')):
             return jsonify({"status": "error", "message": "잘못된 접근입니다."}), 403
        
        contract_type = df.at[idx, '계약구분']
        
        config = {
            '방과후강사': ("새담청소년교육문화원 위탁교육계약서", "수탁학교 :", "담당부서 :"),
            '맞춤형강사': ("새담청소년교육문화원 위탁교육계약서", "수탁학교 :", "담당부서 :"),
            '코디근로자': ("새담청소년교육문화원 센터장 계약서", "수탁학교 :", "직책 :"),
            '코디사업자': ("새담청소년교육문화원 센터장 계약서", "수탁학교 :", "직책 :"),
            '원어민근로자': ("방과후 영어 원어민 강사 위탁 계약서", "School Name :", "Part :"),
            '원어민사업자': ("방과후 영어 원어민 강사 위탁 계약서", "School Name :", "Part :"),
            '안전코디': ("새담청소년교육문화원 위수탁계약서", "수탁학교 :", "직책 :"),
            '직원근로자': ("새담청소년교육문화원 근로계약서", "기관명 :", "직책 :"),
            '직원사업자': ("새담청소년교육문화원 위탁업무계약서", "기관명 :", "위탁업무 :")
        }
        doc_title, school_label, dept_label = config.get(contract_type, (f"새담청소년교육문화원 계약서 ({contract_type})", "수탁학교 :", "담당부서 :"))
        final_school_name = "새담청소년교육문화원" if contract_type in ['직원근로자', '직원사업자'] else data.get('school', '')
        
        stamp_path = os.path.abspath(os.path.join(os.getcwd(), 'static', 'stamp7.png'))
        stamp_uri = f"file://{stamp_path}"

        # [디자인 복구] 서명란 레이아웃 원복
        signature_section = f"""
        <div class="signature-area" style="margin-top: 40px; position: relative; min-height: 150px;">
            <p style="text-align: center; margin-bottom: 50px;">{now_dt.strftime('%Y년 %m월 %d일')}</p>
            <br>
            <div style="float: left; width: 50%; position: relative;">
                <p><b>[위탁자]</b></p>
                <p style="font-size: 20px; line-height: 1.6; position: relative; width: 280px;">
               (사)새담청소년교육문화원
             <span style="display: block; text-align: right; padding-right: 45px;">이사장</span>
            <img src="{stamp_uri}" style="position: absolute; right: -60; bottom: -10px; width: 90px;">
                </p>
            </div>
            <div style="float: right; width: 45%; text-align: left;">
                <p><b>[수탁자]</b></p>
                <p style="line-height: 40px;">
                    성명: {data['name']} <br>
                    서명: <img src="{data['signature']}" style="width: 200px; border-bottom: 1px solid #000; vertical-align: middle; margin-left: 10px;">
                </p>
            </div>
            <div style="clear: both;"></div>
        </div>
        """

        def get_cleaned_content(filename):
            path = os.path.join(TERMS_DIR, filename)
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    c = f.read()
                    for col in df.columns:
                        val = str(final_school_name) if col == '수탁학교명' else (str(df.at[idx, col]) if pd.notna(df.at[idx, col]) else "")
                        c = c.replace(f"{{{{ data.{col} }}}}", val)
                        display_style = "display:none" if not val or val == '0' or str(val).strip() == '' else "display:table-row"
                        c = c.replace(f"{{{{ style.{col} }}}}", display_style)
                    return c.replace('\n', '<br>').replace('<br><table', '<table').replace('</table><br>', '</table>')
            return ""

        content1, content2 = get_cleaned_content(f"{contract_type}.txt"), get_cleaned_content(f"{contract_type}2.txt")
        
        # [디자인 복구] CSS 및 여백 원복
        html_content = f"""
        <html><head><meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;700;900&display=swap" rel="stylesheet">
        <style>
            @page {{ size: A4; margin: 25mm 20mm; }} 
            body {{ margin: 0; padding: 0; font-family: 'Noto Sans KR', sans-serif; background-color: #fff; color: #000; }} 
            .document-wrapper {{ position: relative; z-index: 1; }} 
            .title {{ text-align: center; font-size: 28px; font-weight: bold; margin-bottom: 35px; text-decoration: underline; }} 
            .info-table {{ width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 15px; border: none; }} 
            .info-table th, .info-table td {{ border: none; padding: 8px 5px; text-align: left; }} 
            .info-table th {{ font-weight: bold; width: 15%; color: #333; }} 
            .info-table td {{ width: 35%; border-bottom: 1px solid #eee; }} 
            .terms-area {{ text-align: justify; line-height: 1.6; font-size: 14.5px; margin-top: 10px; word-break: keep-all; }} 
            .signature-area {{ margin-top: 50px; position: relative; font-size: 16px; }} 
        </style></head>
        <body><div class="document-wrapper"><div class="title"><h1 style="text-align:center; line-height:1.4; margin-bottom:30px;"><span style="display:block; font-family:'Noto Sans KR', sans-serif; font-weight:900; font-size:26px; letter-spacing:-0.03em; color:#222;">{doc_title}</span></h1></div><br><table class="info-table"><tr><th>{school_label}</th><td>{final_school_name}</td><th>{dept_label}</th><td>{data.get('dept', '')}</td></tr><tr><th>성명 :</th><td>{data.get('name', '')}</td><th>주민번호 :</th><td>{data.get('ssn', session.get('user_ssn', ''))}</td></tr><tr><th>연락처 :</th><td>{data.get('phone', '')}</td><th>이메일 :</th><td>{data.get('email', '')}</td></tr><tr><th>거주지 :</th><td colspan="3">{data.get('address', '')}</td></tr></table><br><div class="terms-area">{content1}</div>{signature_section}{"<div style='page-break-before: always;'></div>" if content2.strip() else ""}{f"<div class='terms-area' style='margin-top:10mm;'>{content2}</div>{signature_section}" if content2.strip() else ""}</div></body></html>
        """
        
        safe_school, safe_name = str(final_school_name).replace(' ', ''), str(data['name']).replace(' ', '')
        filename = f"{contract_type}_{safe_school}_{safe_name}_{now_dt.strftime('%Y%m%d_%H%M%S')}.pdf"
        pdf_path = os.path.join(CONTRACTS_DIR, filename)
        
        # [디자인 복구] PDF 출력 여백 원복
        pdfkit.from_string(html_content, pdf_path, configuration=PDF_CONFIG, options={'page-size': 'A4', 'encoding': "UTF-8", 'javascript-delay': '1000', 'enable-local-file-access': None, 'margin-top': '25', 'margin-bottom': '25', 'margin-left': '20', 'margin-right': '20'})
        
        # 실제 IP 주소 가져오기 (프록시 대응)
        user_ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        if user_ip and ',' in user_ip:
            user_ip = user_ip.split(',')[0].strip()

        df.at[idx, '연락처'], df.at[idx, 'email'], df.at[idx, '거주지'], df.at[idx, '계약완료일시'], df.at[idx, '파일명'], df.at[idx, 'IP'] = str(data.get('phone', '')), str(data.get('email', '')), str(data.get('address', '')), now_dt.strftime('%Y-%m-%d %H:%M:%S'), filename, user_ip
        df.to_excel(EXCEL_FILE, index=False)
        
        try:
            target_user_email = str(data.get('email', '')).strip()
            if target_user_email and "@" in target_user_email:
                yag = yagmail.SMTP(SENDER_EMAIL, SENDER_PASSWORD)
                yag.send(to=[target_user_email, SENDER_EMAIL], subject=f"[계약완료] {final_school_name} {data['name']}님 {doc_title}", contents=[f"{contract_type} 계약 작성이 완료되었습니다. 첨부된 파일을 확인하세요."], attachments=pdf_path)
        except: pass
        return jsonify({"status": "success", "message": "계약이 완료되었습니다."})
    except Exception as e: return jsonify({"status": "error", "message": f"오류 발생: {str(e)}"}), 500

# --- [관리자 기능 로직] ---
@app.route('/c_admin', methods=['GET', 'POST'])
def admin_page():
    if request.method == 'POST':
        if request.form.get('admin_pw') == ADMIN_PASSWORD:
            session['admin_logged_in'] = True
            return redirect(url_for('admin_page'))
        return "<script>alert('비밀번호가 틀렸습니다.'); history.back();</script>"
    
    if not session.get('admin_logged_in'):
        return '<div style="text-align:center; margin-top:100px; font-family:sans-serif;"><h2>🔐 관리자 인증</h2><form method="POST"><input type="password" name="admin_pw" placeholder="관리자 비밀번호" style="padding:10px; width:250px;"><button type="submit" style="padding:10px 20px;">접속</button></form></div>'

    page = request.args.get('page', 1, type=int)
    per_page = 20
    s_year, s_cat, s_school, s_dept, s_name = request.args.get('year', ''), request.args.get('category', ''), request.args.get('school', ''), request.args.get('dept', ''), request.args.get('name', '')

    try:
        full_df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
        df = full_df.copy().sort_index(ascending=False)
        if s_year: df = df[df['연도'].astype(str).str.contains(s_year)]
        if s_cat: df = df[df['계약구분'] == s_cat]
        if s_school: df = df[df['수탁학교명'] == s_school]
        if s_dept: df = df[df['부서명'] == s_dept]
        if s_name: df = df[df['성명'].str.contains(s_name)]

        years = sorted([str(y) for y in full_df['연도'].unique() if y != ""], reverse=True)
        schools = sorted([s for s in full_df['수탁학교명'].unique() if s != ""])
        depts = sorted([d for d in full_df['부서명'].unique() if d != ""])

        total_pages = (len(df) // per_page) + (1 if len(df) % per_page > 0 else 0)
        total_count = len(df) # 전체 데이터 개수
        items = df.iloc[(page-1)*per_page : page*per_page].to_dict('records')
        for i, item in enumerate(items):
            item['orig_idx'] = df.index[(page-1)*per_page + i]
        return render_template('c_admin_.html', items=items, total_pages=total_pages, total_count=total_count, current_page=page, years=years, schools=schools, depts=depts)
    except Exception as e: return f"에러: {str(e)}"

@app.route('/c_admin/upload_excel', methods=['POST'])
def upload_excel():
    if 'excel_file' not in request.files: return jsonify({'status': 'error', 'message': '파일 없음'}), 400
    file = request.files['excel_file']
    try:
        new_df = pd.read_excel(file, dtype=str)
        existing_df = pd.read_excel(EXCEL_FILE, dtype=str) if os.path.exists(EXCEL_FILE) else pd.DataFrame()
        now_dt_kst = datetime.now(KST)
        if '연도' not in new_df.columns: new_df['연도'] = str(now_dt_kst.year)
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        combined_df.to_excel(EXCEL_FILE, index=False)
        return jsonify({'status': 'success', 'message': f'{len(new_df)}명 추가됨'})
    except Exception as e: return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/c_admin/add', methods=['POST'])
def admin_add():
    try:
        new_data = request.json
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        new_row = {
            '계약구분': new_data.get('계약구분', '방과후강사'), 
            '수탁학교명': new_data.get('수탁학교명'),
            '부서명': new_data.get('부서명'), 
            '성명': new_data.get('성명'), 
            '주민번호': new_data.get('주민번호'),
            '수수료': str(new_data.get('수수료', '0')),
            '보조금': str(new_data.get('보조금', '0')),
            '경력수당': str(new_data.get('경력수당', '0')),
            '직책수당': str(new_data.get('직책수당', '0')),
            '기타': str(new_data.get('기타', '0')),
            '근무시간': new_data.get('근무시간', ''),
            '계약기간': new_data.get('계약기간', ''),
            '연도': str(datetime.now(KST).year), 
            '계약완료일시': "", 
            '파일명': "",
            'IP': ""
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        return jsonify({'status': 'success'})
    except Exception as e: return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/c_admin/delete', methods=['POST'])
def delete_contracts():
    indices = request.json.get('indices', [])
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        for idx in [int(i) for i in indices]:
            if idx in df.index:
                filename = df.at[idx, '파일명']
                if filename and not pd.isna(filename):
                    p = os.path.join(CONTRACTS_DIR, str(filename))
                    if os.path.exists(p): os.remove(p)
        df = df.drop([int(i) for i in indices])
        df.to_excel(EXCEL_FILE, index=False)
        return jsonify({"status": "success"})
    except Exception as e: return jsonify({"status": "error", "message": str(e)})

@app.route('/download_pdf/<int:idx>')
def download_pdf(idx):
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        pdf_path = os.path.join(CONTRACTS_DIR, str(df.at[idx, '파일명']))
        return send_file(pdf_path, mimetype='application/pdf')
    except: return "파일 없음", 404

@app.route('/c_admin/logout')
def admin_logout():
    session.pop('admin_logged_in', None)
    return redirect(url_for('admin_page'))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)