from flask import Flask, render_template, request, jsonify, redirect, url_for, session, send_file, abort, Response
import pandas as pd
import os
import zipfile
import io
import tempfile
import traceback
from urllib.parse import quote
from html import escape
import pdfkit
import yagmail
import PyPDF2  # PDF 페이지 분리를 위한 라이브러리 추가
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

def format_value(val):
    """소수점(0.85)은 퍼센트(85%)로, 큰 숫자는 콤마(,) 형식으로 변환"""
    if not val or pd.isna(val) or str(val).strip() == "":
        return ""
    val = str(val).strip()
    try:
        num = float(val)
        if 0 < num < 1:
            return f"{int(num * 100)}%"
        if num >= 100:
            return "{:,}".format(int(num))
    except ValueError:
        pass
    return val

def init_excel():
    """엑셀 초기화 (모든 읽기 작업에 dtype=str 적용)"""
    columns = [
        '계약구분', '수탁학교명', '부서명', '성명', '주민번호', '수수료', '비고1', '보조금', '비고2', '경력수당', '비고3', '직책수당', '비고4', '기타', '근무시간', '계약기간', 'email', '연락처', '거주지', '계약완료일시', '연도', '파일명', 'IP'
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

# --- [수정된 zip 다운로드 로직] ---

@app.route('/admin/download_selected')
def download_selected_contracts():
    id_param = request.args.get('ids', '')
    limit = request.args.get('limit', type=int) # limit 파라미터 추가
    
    if not id_param:
        return "<script>alert('선택된 항목이 없습니다.'); history.back();</script>", 400
        
    try:
        target_indices = [int(i) for i in id_param.split(',')]
        df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
        memory_file = io.BytesIO()
        
        with zipfile.ZipFile(memory_file, 'w') as zf:
            file_count = 0
            for idx in target_indices:
                if idx in df.index:
                    filename = df.at[idx, '파일명']
                    if filename:
                        file_path = os.path.join(CONTRACTS_DIR, filename)
                        if os.path.exists(file_path):
                            
                            # limit(페이지 제한)이 설정된 경우 PDF 분리 작업 수행
                            if limit:
                                try:
                                    reader = PyPDF2.PdfReader(file_path)
                                    writer = PyPDF2.PdfWriter()
                                    
                                    # 요청한 페이지 수와 실제 PDF 페이지 수 중 작은 값을 선택하여 에러 방지
                                    pages_to_extract = min(limit, len(reader.pages))
                                    
                                    for page_num in range(pages_to_extract):
                                        writer.add_page(reader.pages[page_num])
                                        
                                    pdf_bytes = io.BytesIO()
                                    writer.write(pdf_bytes)
                                    pdf_bytes.seek(0)
                                    
                                    # 잘라낸 PDF 데이터를 zip 파일에 쓰기
                                    zf.writestr(filename, pdf_bytes.read())
                                    file_count += 1
                                    
                                except Exception as pdf_e:
                                    # PDF 처리에 실패할 경우 안전하게 원본 파일을 그대로 포함
                                    print(f"PDF 분리 실패 ({filename}): {pdf_e}")
                                    zf.write(file_path, arcname=filename)
                                    file_count += 1
                            else:
                                # 페이지 제한이 없으면 기존처럼 원본 그대로 포함
                                zf.write(file_path, arcname=filename)
                                file_count += 1
            
            if file_count == 0:
                return "<script>alert('선택한 항목 중 작성 완료된 PDF 파일이 없습니다.'); history.back();</script>"
        
        memory_file.seek(0)
        current_time = datetime.now(KST).strftime('%Y%m%d_%H%M%S')
        # 파일명으로 페이지 제한 버전인지 전체 버전인지 구분
        prefix = f"1-{limit}page_" if limit else "" 
        
        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'selected_{prefix}contracts_{current_time}.zip'
        )
    except Exception as e:
        return f"다운로드 중 오류 발생: {str(e)}", 500

# --- [전체 계약서 + 계약리스트 일괄 다운로드] ---

def _download_error_page(title, detail=""):
    """다운로드 실패 시 관리자에게 원인을 보여주는 안내 페이지"""
    safe_detail = escape(str(detail))
    return f"""
    <div style="font-family:'Pretendard','Malgun Gothic',sans-serif; max-width:760px; margin:80px auto; padding:30px;
                border:1px solid #f0c2c2; border-radius:12px; background:#fff8f8;">
        <h2 style="color:#c92a2a; margin-top:0;">⚠ 전체 백업 다운로드 실패</h2>
        <p style="font-size:1.05rem; color:#333;">{escape(str(title))}</p>
        <pre style="white-space:pre-wrap; word-break:break-all; background:#fff; border:1px solid #eee;
                    border-radius:8px; padding:15px; font-size:0.85rem; color:#555; max-height:340px; overflow:auto;">{safe_detail}</pre>
        <p style="font-size:0.85rem; color:#888;">위 내용을 관리자에게 전달하시면 원인 확인에 도움이 됩니다.</p>
        <button onclick="location.href='/c_admin'"
                style="padding:12px 22px; background:#002c63; color:#fff; border:none; border-radius:8px;
                       font-weight:700; cursor:pointer;">관리자 페이지로 돌아가기</button>
    </div>
    """, 500


def _iter_and_cleanup(path, chunk_size=65536):
    """임시 zip 파일을 스트리밍으로 전송한 뒤 삭제"""
    try:
        with open(path, 'rb') as f:
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                yield chunk
    finally:
        try:
            os.remove(path)
        except Exception:
            pass


def _safe_arcname(raw):
    """엑셀에 저장된 파일명에서 경로 구분자를 제거해 안전한 파일명만 추출"""
    name = str(raw).strip().replace(chr(92), '/')
    name = name.split('/')[-1].strip()
    return name


@app.route('/admin/download_all')
def download_all_contracts():
    """서버에 저장된 전체 계약서(PDF) + 계약리스트(엑셀)를 하나의 ZIP으로 다운로드

    ?debug=1 : 파일을 내려받지 않고 처리 결과 요약만 화면에 표시(장애 진단용)
    """
    if not session.get('admin_logged_in'):
        return "<script>alert('관리자 로그인이 필요합니다.'); location.href='/c_admin';</script>", 403

    debug_mode = request.args.get('debug') == '1'
    now_dt = datetime.now(KST)
    stamp = now_dt.strftime('%Y%m%d_%H%M%S')
    errors = []

    # 1) 계약리스트 읽기 (실패하면 계약서만이라도 받을 수 있도록 빈 표로 진행)
    df = None
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
    except Exception as e:
        app.logger.exception('[download_all] 계약리스트 읽기 실패')
        errors.append(f'계약리스트 읽기 실패 : {type(e).__name__}: {e}')
        df = pd.DataFrame()

    tmp_path = None
    try:
        # 임시 zip 저장 위치 : 기본 임시폴더 → 데이터 디스크 순으로 시도(디스크 부족 대비)
        tmp_path, last_err = None, None
        for tmp_dir in [None, MOUNT_PATH, os.getcwd()]:
            try:
                fd, tmp_path = tempfile.mkstemp(prefix='saedam_backup_', suffix='.zip', dir=tmp_dir)
                os.close(fd)
                break
            except Exception as e:
                last_err = e
                tmp_path = None
        if not tmp_path:
            raise IOError(f'임시 파일을 만들 수 없습니다 : {last_err}')

        included, missing, extra = {}, [], []

        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED, allowZip64=True) as zf:
            # (1) 계약리스트 엑셀 : 실패 시 원본파일 → CSV 순으로 대체
            list_name = f'계약리스트_{stamp}.xlsx'
            try:
                list_bytes = io.BytesIO()
                df.to_excel(list_bytes, index=False)
                zf.writestr(list_name, list_bytes.getvalue())
            except Exception as e:
                app.logger.exception('[download_all] 리스트 엑셀 생성 실패')
                errors.append(f'엑셀 생성 실패, 원본 파일로 대체 : {type(e).__name__}: {e}')
                try:
                    zf.write(EXCEL_FILE, arcname=list_name)
                except Exception as e2:
                    errors.append(f'원본 엑셀 첨부 실패, CSV로 대체 : {type(e2).__name__}: {e2}')
                    try:
                        zf.writestr(f'계약리스트_{stamp}.csv',
                                    df.to_csv(index=False).encode('utf-8-sig'))
                    except Exception as e3:
                        errors.append(f'CSV 대체도 실패 : {type(e3).__name__}: {e3}')

            # (2) 계약리스트에 등록된 계약서 PDF
            #     PDF는 이미 압축된 형식이므로 무압축(STORED)으로 담아 처리 시간을 줄임
            if '파일명' in df.columns:
                for idx, row in df.iterrows():
                    try:
                        raw = str(row.get('파일명', '')).strip()
                        if not raw or raw.lower() in ('nan', 'none', 'nat'):
                            continue
                        filename = _safe_arcname(raw)
                        if not filename:
                            continue
                        if filename in included:
                            continue
                        file_path = os.path.join(CONTRACTS_DIR, filename)
                        if os.path.isfile(file_path):
                            zf.write(file_path, arcname=f'계약서/{filename}',
                                     compress_type=zipfile.ZIP_STORED)
                            included[filename] = True
                        else:
                            missing.append(f"{idx} / {row.get('성명', '')} / {row.get('수탁학교명', '')} / {filename}")
                    except Exception as e:
                        app.logger.exception('[download_all] 계약서 압축 실패')
                        errors.append(f'{idx}행 계약서 압축 실패 : {type(e).__name__}: {e}')
            elif len(df) > 0:
                errors.append("계약리스트에 '파일명' 열이 없어 등록 계약서를 찾지 못했습니다.")

            # (3) 리스트에는 없지만 서버에 남아있는 파일 (하위 폴더 포함)
            try:
                if os.path.isdir(CONTRACTS_DIR):
                    for root, _dirs, files in os.walk(CONTRACTS_DIR):
                        for filename in sorted(files):
                            try:
                                file_path = os.path.join(root, filename)
                                rel_path = os.path.relpath(file_path, CONTRACTS_DIR).replace(os.sep, '/')
                                if rel_path in included or filename in included:
                                    continue
                                zf.write(file_path, arcname=f'계약서_목록외/{rel_path}',
                                         compress_type=zipfile.ZIP_STORED)
                                extra.append(rel_path)
                            except Exception as e:
                                app.logger.exception('[download_all] 목록 외 파일 압축 실패')
                                errors.append(f'목록 외 파일 압축 실패({filename}) : {type(e).__name__}: {e}')
                else:
                    errors.append(f'계약서 보관 폴더를 찾을 수 없습니다 : {CONTRACTS_DIR}')
            except Exception as e:
                app.logger.exception('[download_all] 계약서 폴더 탐색 실패')
                errors.append(f'계약서 폴더 탐색 실패 : {type(e).__name__}: {e}')

            # (4) 요약 정보
            try:
                done_count = 0
                if '파일명' in df.columns:
                    done_count = int((df['파일명'].astype(str).str.strip() != "").sum())
                summary = [
                    f"다운로드 일시 : {now_dt.strftime('%Y-%m-%d %H:%M:%S')} (KST)",
                    f"보관 경로 : {CONTRACTS_DIR}",
                    f"계약리스트 총 건수 : {len(df)}건",
                    f"계약 완료(파일명 등록) 건수 : {done_count}건",
                    f"포함된 계약서 PDF : {len(included)}개",
                    f"목록 외 파일(리스트 미등록) : {len(extra)}개",
                    f"파일 없음(리스트에는 있으나 서버에 없음) : {len(missing)}건",
                    f"처리 중 오류 : {len(errors)}건",
                    "",
                    "[포함된 폴더 구성]",
                    f"  {list_name} : 전체 계약 리스트",
                    "  계약서/ : 계약리스트에 등록된 계약서 PDF",
                    "  계약서_목록외/ : 리스트에 없으나 서버에 보관 중인 파일",
                ]
                if missing:
                    summary += ["", "[파일 없음 목록] 번호 / 성명 / 수탁학교명 / 파일명"] + missing
                if extra:
                    summary += ["", "[목록 외 파일]"] + extra
                if errors:
                    summary += ["", "[처리 중 오류]"] + errors
                zf.writestr('다운로드_요약.txt', chr(10).join(summary).encode('utf-8'))
            except Exception as e:
                app.logger.exception('[download_all] 요약 생성 실패')

        file_size = os.path.getsize(tmp_path)
        if file_size <= 0:
            raise IOError('생성된 ZIP 파일이 비어 있습니다.')

    except Exception as e:
        app.logger.exception('[download_all] ZIP 생성 실패')
        if tmp_path:
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        return _download_error_page(
            '압축 파일을 만드는 중 오류가 발생했습니다.',
            f'{type(e).__name__}: {e}{chr(10)}{chr(10)}{traceback.format_exc()}'
        )

    # 진단 모드 : 실제 다운로드 없이 처리 결과만 확인
    if debug_mode:
        try:
            with zipfile.ZipFile(tmp_path) as zf:
                names = zf.namelist()
                report = [
                    f'ZIP 크기 : {file_size:,} bytes',
                    f'항목 수 : {len(names)}개',
                    f'보관 경로 : {CONTRACTS_DIR}',
                    f'리스트 파일 : {EXCEL_FILE}',
                    '',
                    zf.read('다운로드_요약.txt').decode('utf-8', 'replace'),
                    '',
                    '[ZIP 내부 목록]',
                ] + names
            return Response(chr(10).join(report), content_type='text/plain; charset=utf-8')
        finally:
            try:
                os.remove(tmp_path)
            except Exception:
                pass

    download_name = f'전체계약백업_{stamp}.zip'
    quoted_name = quote(download_name)
    return Response(
        _iter_and_cleanup(tmp_path),
        mimetype='application/zip',
        headers={
            'Content-Disposition': f"attachment; filename=\"contracts_backup_{stamp}.zip\"; filename*=UTF-8''{quoted_name}",
            'Content-Length': str(file_size),
            'Cache-Control': 'no-store',
        }
    )

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
                        raw_val = user_data.get(col, '')
                        if col in ['수수료', '보조금', '경력수당', '직책수당', '기타']:
                            val = format_value(raw_val)
                        else:
                            val = str(raw_val) if pd.notna(raw_val) else ""
                        
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

        signature_section = f"""
        <div class="signature-area" style="margin-top: 40px; position: relative; min-height: 150px;">
            <p style="text-align: center; margin-bottom: 50px;">{now_dt.strftime('%Y년 %m월 %d일')}</p>
            <br>
            <div style="float: left; width: 50%; position: relative;">
                <p><b>[위탁자]</b></p>
                <p style="font-size: 20px; line-height: 1.6; position: relative; width: 280px;">
               (사)새담청소년교육문화원
             <span style="display: block; text-align: right; padding-right: 64px;">이사장</span>
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
                        raw_val = str(final_school_name) if col == '수탁학교명' else (str(df.at[idx, col]) if pd.notna(df.at[idx, col]) else "")
                        if col in ['수수료', '보조금', '경력수당', '직책수당', '기타']:
                            val = format_value(raw_val)
                        else:
                            val = raw_val
                        c = c.replace(f"{{{{ data.{col} }}}}", val)
                        display_style = "display:none" if not val or val == '0' or str(val).strip() == '' else "display:table-row"
                        c = c.replace(f"{{{{ style.{col} }}}}", display_style)
                    return c.replace('\n', '<br>').replace('<br><table', '<table').replace('</table><br>', '</table>')
            return ""

        content1, content2 = get_cleaned_content(f"{contract_type}.txt"), get_cleaned_content(f"{contract_type}2.txt")
        
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
        display_contract_type = "센터장" if contract_type in ['코디사업자', '코디근로자'] else contract_type
        filename = f"{display_contract_type}_{safe_school}_{safe_name}_{now_dt.strftime('%Y%m%d_%H%M%S')}.pdf"
        pdf_path = os.path.join(CONTRACTS_DIR, filename)
        
        pdfkit.from_string(html_content, pdf_path, configuration=PDF_CONFIG, options={'page-size': 'A4', 'encoding': "UTF-8", 'javascript-delay': '1000', 'enable-local-file-access': None, 'margin-top': '25', 'margin-bottom': '25', 'margin-left': '20', 'margin-right': '20'})
        
        user_ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        if user_ip and ',' in user_ip: user_ip = user_ip.split(',')[0].strip()

        df.at[idx, '연도'] = str(now_dt.year)
        df.at[idx, '연락처'] = str(data.get('phone', ''))
        df.at[idx, 'email'] = str(data.get('email', ''))
        df.at[idx, '거주지'] = str(data.get('address', ''))
        df.at[idx, '계약완료일시'] = now_dt.strftime('%Y-%m-%d %H:%M:%S')
        df.at[idx, '파일명'] = filename
        df.at[idx, 'IP'] = user_ip
        
        df.to_excel(EXCEL_FILE, index=False)
        
        try:
            target_user_email = str(data.get('email', '')).strip()
            if target_user_email and "@" in target_user_email:
                yag = yagmail.SMTP(SENDER_EMAIL, SENDER_PASSWORD)
                yag.send(to=[target_user_email, SENDER_EMAIL], subject=f"[계약완료] {data['name']}님 {doc_title} ({final_school_name})", contents=[f" '{doc_title}' 계약이 완료되었습니다. \n\n첨부된 파일을 확인하세요."], attachments=pdf_path)
        except: pass
        return jsonify({"status": "success", "message": "계약이 정상적으로 완료되었으며 이메일로 발송되었습니다."})
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
        return '''
        <div style="text-align:center; margin-top:100px; font-family:'Pretendard', sans-serif;">
            <div style="display:inline-block; padding:40px; border:1px solid #ddd; border-radius:15px; box-shadow:0 4px 15px rgba(0,0,0,0.1); background:#fff;">
                <h2 style="color:#002c63; margin-bottom:25px; font-weight:800;">🔐 관리자 인증</h2>
                <form method="POST" id="adminLoginForm">
                    <div style="margin-bottom:15px;">
                        <input type="password" name="admin_pw" id="admin_pw" placeholder="관리자 비밀번호" 
                               style="padding:12px; width:280px; border:1px solid #ccc; border-radius:8px; font-size:1rem;">
                    </div>
                    <div style="margin-bottom:20px; text-align:left; padding-left:5px;">
                        <label style="font-size:0.9rem; color:#666; cursor:pointer; display:flex; align-items:center; gap:8px;">
                            <input type="checkbox" id="remember_pw" style="width:16px; height:16px; cursor:pointer;"> 비밀번호 저장
                        </label>
                    </div>
                    <button type="submit" style="padding:12px; width:100%; background:#002c63; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer; font-size:1rem;">접속하기</button>
                </form>
                <div style="text-align: center; margin-bottom: 10px;">
                    <img src="http://www.saedam.org/img/logo01.gif" width="100" alt="Logo">
                </div>
            </div>
        </div>
        <script>
            const pwInput = document.getElementById('admin_pw');
            const rememberChk = document.getElementById('remember_pw');
            const loginForm = document.getElementById('adminLoginForm');
            window.onload = function() {
                const savedPw = localStorage.getItem('saedam_admin_pw');
                if (savedPw) {
                    pwInput.value = savedPw;
                    rememberChk.checked = true;
                }
            };
            loginForm.onsubmit = function() {
                if (rememberChk.checked) {
                    localStorage.setItem('saedam_admin_pw', pwInput.value);
                } else {
                    localStorage.removeItem('saedam_admin_pw');
                }
            };
        </script>
        '''

    page = request.args.get('page', 1, type=int)
    per_page = 20
    s_year, s_cat, s_school, s_dept, s_name = request.args.get('year', ''), request.args.get('category', ''), request.args.get('school', ''), request.args.get('dept', ''), request.args.get('name', '')

    try:
        full_df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
        
        total_count = len(full_df)
        completed_count = len(full_df[full_df['계약완료일시'].str.strip() != ""])
        pending_count = total_count - completed_count
        completion_rate = round((completed_count / total_count * 100), 1) if total_count > 0 else 0

        # --- [필터링 로직 수정 및 들여쓰기 교정 구간] ---
        df = full_df.copy().sort_index(ascending=False)
        
        if s_year: 
            df = df[df['연도'].astype(str).str.contains(s_year)]
        
        if s_cat:
            if s_cat == '미작성':
                # 계약완료일시가 비어있는 행만 추출
                df = df[df['계약완료일시'].astype(str).str.strip() == ""]
            else:
                # 그 외 카테고리는 기존 방식대로 추출
                df = df[df['계약구분'] == s_cat]
        
        if s_school: 
            df = df[df['수탁학교명'] == s_school]
        if s_dept: 
            df = df[df['부서명'] == s_dept]
        if s_name: 
            df = df[df['성명'].str.contains(s_name)]
        # -----------------------------------------------

        years = sorted([str(y) for y in full_df['연도'].unique() if y != ""], reverse=True)
        schools = sorted([s for s in full_df['수탁학교명'].unique() if s != ""])
        depts = sorted([d for d in full_df['부서명'].unique() if d != ""])

        total_pages = (len(df) // per_page) + (1 if len(df) % per_page > 0 else 0)
        filtered_count = len(df)
        items = df.iloc[(page-1)*per_page : page*per_page].to_dict('records')
        
        page_indices = df.index[(page-1)*per_page : page*per_page]
        for i, item in enumerate(items):
            item['orig_idx'] = page_indices[i]

        display_size = 20 
        move_size = 10     
        start_page = max(1, ((page - 1) // display_size) * display_size + 1)
        end_page = min(total_pages, start_page + display_size - 1)
        prev_block = max(1, page - move_size)
        next_block = min(total_pages, page + move_size)

        return render_template('c_admin_.html', 
                               items=items, 
                               total_pages=total_pages, 
                               current_page=page,
                               start_page=start_page,
                               end_page=end_page,
                               prev_block=prev_block,
                               next_block=next_block,
                               total_count=total_count,
                               completed_count=completed_count,
                               pending_count=pending_count,
                               completion_rate=completion_rate,
                               filtered_count=filtered_count,
                               years=years, 
                               schools=schools, 
                               depts=depts)
    except Exception as e: 
        return f"에러: {str(e)}"

@app.route('/c_admin/upload_excel', methods=['POST'])
def upload_excel():
    if 'excel_file' not in request.files: return jsonify({'status': 'error', 'message': '파일 없음'}), 400
    file = request.files['excel_file']
    try:
        new_df = pd.read_excel(file, dtype=str)
        target_cols = ['수수료', '보조금', '경력수당', '직책수당', '기타']
        for col in target_cols:
            if col in new_df.columns:
                new_df[col] = new_df[col].apply(format_value)
        
        if '연도' not in new_df.columns:
            new_df['연도'] = ""
        else:
            new_df['연도'] = new_df['연도'].fillna("")

        existing_df = pd.read_excel(EXCEL_FILE, dtype=str) if os.path.exists(EXCEL_FILE) else pd.DataFrame()
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        combined_df.to_excel(EXCEL_FILE, index=False)
        return jsonify({'status': 'success', 'message': f'{len(new_df)}명의 계약정보가 추가 되었습니다.'})
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
            '수수료': format_value(new_data.get('수수료', '0')),
            '비고1': new_data.get('비고1', ''), 
            '보조금': format_value(new_data.get('보조금', '0')),
            '비고2': new_data.get('비고2', ''), 
            '경력수당': format_value(new_data.get('경력수당', '0')),
            '비고3': new_data.get('비고3', ''), 
            '직책수당': format_value(new_data.get('직책수당', '0')),
            '비고4': new_data.get('비고4', ''), 
            '기타': format_value(new_data.get('기타', '0')),
            '근무시간': new_data.get('근무시간', ''),
            '계약기간': new_data.get('계약기간', ''),
            '연도': "", 
            '계약완료일시': "", '파일명': "", 'IP': ""
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