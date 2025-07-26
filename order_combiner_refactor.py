#####################################################################################################################
# exe파일 생성 (폴더 - 터미널에서 열기 - 아래의 문장 붙여넣기)
# cd C:\pythonProjects\orderCombiner
# pyinstaller --noconsole --onefile --add-data "json_files/jason-455001-adc816559e51.json;json_files" order_combiner.py
#####################################################################################################################

# 표준 라이브러리 및 외부 모듈 import
import os  # 운영체제 경로 및 파일 관련 기능 사용
import pandas as pd  # 데이터프레임 및 데이터 처리
import tkinter as tk  # GUI 생성을 위한 tkinter
from tkinter import filedialog, scrolledtext  # 파일 다이얼로그, 스크롤 텍스트 위젯
from googleapiclient.discovery import build  # 구글 API 클라이언트 빌드 함수
from google.oauth2.service_account import Credentials  # 구글 서비스 계정 인증
import threading  # 멀티스레드 처리를 위한 모듈
import time  # 시간 지연 및 측정
import datetime  # 날짜/시간 처리
import re  # 정규표현식 (범위 파싱용)
from get_coupang_data import get_coupang_data  # 쿠팡 데이터 처리 함수
from get_smart_data import get_smart_data  # 스마트스토어 데이터 처리 함수
from get_esm_data import get_esm_data  # ESM 데이터 처리 함수
from get_street_data import get_st11_data  # 11번가 데이터 처리 함수
from tkinterdnd2 import TkinterDnD, DND_FILES  # 드래그앤드롭 기능을 위한 모듈

#####################################################################################################################
# 상수 정의 (컬럼 이름, 색상 등 유지보수 쉽게 모음)
#####################################################################################################################

# 컬럼 이름 상수 (이전 대화 주석 기반, 필요 시 수정)
COLUMNS = {
  '주문일': '주문일',
  '샵': '샵',
  '주문번호': '주문번호',
  '고객명': '고객명',
  '통관번호': '통관번호',
  '휴대폰': '휴대폰',
  'POST': 'POST',
  '주소': '주소',
  '메세지': '메세지',
  '상품코드': '상품코드',
  '상품명': '상품명',
  '영문명': '영문명',
  '수량': '수량',
  '구매일': '구매일',
  '구매정보': '구매정보',
  '배송현황': '배송현황',
  '국내운송장': '국내운송장',
  '유니패스': '유니패스',
  '조회': '조회',
  '순수익': '순수익',
  '해외비용': '해외비용',
  '구매비': '구매비',
  '배송비': '배송비',
  '총비용': '총비용',
  '순손익': '순손익',
  '통관검증': '통관검증'
}

# 색상 상수 (샵별 배경색, 날짜별 색 등)
COLORS = {
  'odd_day': '#fff2cc',  # 홀수 날짜
  'even_day': '#faff9f',  # 짝수 날짜
  'gray': {"red": 0.8, "green": 0.8, "blue": 0.8},  # 배송현황 회색
  'shops': {
    "쿠팡": "#ff6699",
    "스마트스토어": "#a9d08e",
    "G마켓": "#00b050",
    "옥션": "#ffc000",
    "11번가": "#548dd4"
  }
}

# Google Sheets 설정 상수
SPREADSHEET_ID = '1WO__vUDFBnQNXSkKDsm3R-UeuyDMN1C6Dlo4anrkDEY'  # order_list 스프레드시트 ID
FAS_SPREADSHEET_ID = '18_Uj7-18Iosw8I2BbTVSWhUodpbQBbVIdGIOrlVwqG8'  # FAS 스프레드시트 ID
RANGE_NAME = 'order'  # 사용할 시트 이름
FAS_RANGE_NAME = 'FAS'  # FAS 시트 이름

#####################################################################################################################
# 함수 정의
#####################################################################################################################

def clean_data(data):
  """데이터 리스트의 각 항목에서 NaN 값을 빈 문자열로 변환하고, 문자열로 변환 후 공백 제거"""
  cleaned_data = []
  for item in data:
    cleaned_item = {}
    for key, value in item.items():
      if pd.isna(value):
        cleaned_item[key] = ''
      else:
        cleaned_item[key] = str(value).strip()
    cleaned_data.append(cleaned_item)
  return cleaned_data

def get_fas_data(service):
  """Google Sheets의 FAS 시트에서 데이터를 읽어와 딕셔너리 리스트로 반환"""
  result = service.spreadsheets().values().get(
    spreadsheetId=FAS_SPREADSHEET_ID,
    range=FAS_RANGE_NAME
  ).execute()
  values = result.get('values', [])
  if not values:
    log_message("FAS 시트에서 데이터를 찾을 수 없습니다.")
    return []
  keys = values[0]
  data_rows = values[1:]
  return [dict(zip(keys, row)) for row in data_rows]

def enrich_data_with_fas(data, service):
  """FAS 데이터와 매칭하여 입력 데이터에 영문명, 링크, 상품코드 등 추가 및 가공 (수식은 여기서 넣지 않음)"""
  fas_data = get_fas_data(service)
  fas_lookup = {item.get('korName', ''): item for item in fas_data}
  enriched_data = []

  for item in data:
    enriched_item = item.copy()
    product_name = item.get(COLUMNS['상품명'], '')
    if product_name in fas_lookup:
      fas_item = fas_lookup[product_name]
      engName = fas_item.get('engName', '')
      link = fas_item.get('link', '')
      enriched_item[COLUMNS['영문명']] = f'=HYPERLINK("{link}", "{engName}")'  # 간단한 하이퍼링크만 여기서
      packQty = fas_item.get('packQty', 1)
      enriched_item[COLUMNS['수량']] = int(packQty) * int(item.get(COLUMNS['수량'], 1))
      if int(item.get(COLUMNS['수량'], 0)) >= 2:
        enriched_item[COLUMNS['구매정보']] = '⠀'
      if item.get(COLUMNS['메세지'], '') == '문 앞':
        enriched_item[COLUMNS['메세지']] = '문 앞에 놓아주세요.'
      enriched_item[COLUMNS['상품코드']] = fas_item.get('sbCode', '')
      enriched_item[COLUMNS['배송현황']] = '미구입'
      enriched_item[COLUMNS['구매비']] = 0
      enriched_item[COLUMNS['배송비']] = 0
    else:
      log_message(f"FAS 데이터에서 '{product_name}'에 해당하는 항목을 찾을 수 없습니다.")
    enriched_data.append(enriched_item)
  return enriched_data

def get_existing_order_ids(service):
  """구글 시트에서 이미 존재하는 주문번호 집합 반환"""
  result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=RANGE_NAME
  ).execute()
  values = result.get('values', [])
  if not values:
    return set()
  headers = values[0]
  try:
    order_id_index = headers.index(COLUMNS['주문번호'])
  except ValueError:
    raise Exception(f"시트에 '{COLUMNS['주문번호']}' 컬럼이 존재하지 않습니다.")
  existing_ids = {row[order_id_index].strip() for row in values[1:] if len(row) > order_id_index and row[order_id_index].strip()}
  return existing_ids

def get_sheet_id(service, sheet_name):
  """시트 이름으로 sheetId 반환"""
  spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
  sheets = spreadsheet.get('sheets', [])
  for sheet in sheets:
    if sheet['properties']['title'] == sheet_name:
      return sheet['properties']['sheetId']
  raise Exception(f"Sheet '{sheet_name}' not found.")

def update_formulas_and_styles(service, start_row, end_row, headers, sheet_id, platform_name):
  """새 행에 수식과 스타일 동적으로 적용 (컬럼 이름 기반)"""
  requests = []

  # 컬럼 인덱스 캐시 (유지보수 쉽게)
  col_indices = {col: headers.index(col) for col in COLUMNS.values() if col in headers}

  for row in range(start_row, end_row + 1):
    # 조회 수식 (국내운송장 기반 동적)
    if all(c in col_indices for c in [COLUMNS['조회'], COLUMNS['국내운송장']]):
      transport_col = chr(65 + col_indices[COLUMNS['국내운송장']])
      requests.append({
        "updateCells": {
          "range": {
            "sheetId": sheet_id,
            "startRowIndex": row - 1,
            "endRowIndex": row,
            "startColumnIndex": col_indices[COLUMNS['조회']],
            "endColumnIndex": col_indices[COLUMNS['조회']] + 1
          },
          "rows": [{
            "values": [{
              "userEnteredValue": {
                "formulaValue": f'=HYPERLINK(IF(LEFT({transport_col}{row},1)="6", "https://search.daum.net/search?nil_suggest=btn&w=tot&DA=SBC&q=우체국택배조회"&{transport_col}{row}, "https://search.daum.net/search?w=tot&DA=YZR&t__nil_searchbox=btn&q=cj택배조회"&{transport_col}{row}), "조회")'
              }
            }]
          }],
          "fields": "userEnteredValue"
        }
      })

    # 총비용 수식
    if all(c in col_indices for c in [COLUMNS['총비용'], COLUMNS['구매비'], COLUMNS['배송비']]):
      buy_col = chr(65 + col_indices[COLUMNS['구매비']])
      ship_col = chr(65 + col_indices[COLUMNS['배송비']])
      requests.append({
        "updateCells": {
          "range": {
            "sheetId": sheet_id,
            "startRowIndex": row - 1,
            "endRowIndex": row,
            "startColumnIndex": col_indices[COLUMNS['총비용']],
            "endColumnIndex": col_indices[COLUMNS['총비용']] + 1
          },
          "rows": [{
            "values": [{
              "userEnteredValue": {
                "formulaValue": f'={buy_col}{row}+{ship_col}{row}'
              }
            }]
          }],
          "fields": "userEnteredValue"
        }
      })

    # 순손익 수식
    if all(c in col_indices for c in [COLUMNS['순손익'], COLUMNS['순수익'], COLUMNS['총비용']]):
      profit_col = chr(65 + col_indices[COLUMNS['순수익']])
      total_col = chr(65 + col_indices[COLUMNS['총비용']])
      requests.append({
        "updateCells": {
          "range": {
            "sheetId": sheet_id,
            "startRowIndex": row - 1,
            "endRowIndex": row,
            "startColumnIndex": col_indices[COLUMNS['순손익']],
            "endColumnIndex": col_indices[COLUMNS['순손익']] + 1
          },
          "rows": [{
            "values": [{
              "userEnteredValue": {
                "formulaValue": f'={profit_col}{row}-{total_col}{row}'
              }
            }]
          }],
          "fields": "userEnteredValue"
        }
      })

    # 배송현황 회색 스타일
    if COLUMNS['배송현황'] in col_indices:
      requests.append({
        "updateCells": {
          "range": {
            "sheetId": sheet_id,
            "startRowIndex": row - 1,
            "endRowIndex": row,
            "startColumnIndex": col_indices[COLUMNS['배송현황']],
            "endColumnIndex": col_indices[COLUMNS['배송현황']] + 1
          },
          "rows": [{"values": [{"userEnteredFormat": {"backgroundColor": COLORS['gray']}}]}],
          "fields": "userEnteredFormat.backgroundColor"
        }
      })

  if requests:
    try:
      service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests}
      ).execute()
      log_message(f"[{platform_name}] 새 행 수식 및 스타일 적용 완료!")
    except Exception as e:
      log_message(f"[{platform_name}] 수식/스타일 적용 오류: {str(e)}")

def append_new_orders(service, dict_list, platform_name):
  """중복되지 않은 신규 주문 데이터만 구글 시트에 추가하고, 수식/스타일 적용"""
  existing_ids = get_existing_order_ids(service)
  new_data = []
  for item in dict_list:
    order_id = str(item.get(COLUMNS['주문번호'], '')).strip()
    if not order_id:
      log_message(f"[{platform_name}] 주문번호가 없는 데이터: {item}")
      continue
    if order_id in existing_ids:
      log_message(f"[{platform_name}] 중복된 주문번호: {order_id}")
      continue
    new_data.append(item)

  if not new_data:
    log_message(f"[{platform_name}] 추가할 데이터가 없습니다.")
    return

  new_data = clean_data(new_data)
  headers = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=f"{RANGE_NAME}!1:1"
  ).execute().get('values', [[]])[0]

  rows = [[item.get(header, '') for header in headers] for item in new_data]
  body = {'values': rows}
  result = service.spreadsheets().values().append(
    spreadsheetId=SPREADSHEET_ID,
    range=RANGE_NAME,
    valueInputOption='USER_ENTERED',
    insertDataOption='INSERT_ROWS',
    body=body
  ).execute()
  log_message(f"[{platform_name}] {len(rows)}개 데이터 추가 완료!")

  # 새 행 범위 계산
  updated_range = result.get('updates', {}).get('updatedRange', '')
  if updated_range:
    try:
      range_part = updated_range.split('!')[1]
      start_cell, end_cell = range_part.split(':')
      start_row = int(re.search(r'\d+$', start_cell).group())
      end_row = int(re.search(r'\d+$', end_cell).group())
      sheet_id = get_sheet_id(service, RANGE_NAME)
      update_formulas_and_styles(service, start_row, end_row, headers, sheet_id, platform_name)
    except Exception as e:
      log_message(f"[{platform_name}] 새 행 범위 계산 오류: {str(e)}")

def process_coupang(file_path):
  """쿠팡 파일 처리 함수"""
  log_message(f"쿠팡 파일 처리 중: {file_path}")
  return get_coupang_data(file_path)

def process_smartstore(file_path):
  """스마트스토어 파일 처리 함수"""
  log_message(f"스마트스토어 파일 처리 중: {file_path}")
  return get_smart_data(file_path)

def process_esm(file_path):
  """ESM 파일 처리 함수"""
  log_message(f"ESM 파일 처리 중: {file_path}")
  return get_esm_data(file_path)

def process_st11(file_path):
  """11번가 파일 처리 함수"""
  log_message(f"11번가 파일 처리 중: {file_path}")
  return get_st11_data(file_path)

def select_file(platform_name):
  """파일 선택 다이얼로그를 열고 선택된 파일 경로를 selected_files 딕셔너리에 저장"""
  file_path = filedialog.askopenfilename(
    title=f"{platform_name} 엑셀 파일 선택",
    filetypes=[("Excel files", "*.xlsx *.xls")]
  )
  if file_path:
    log_message(f"{platform_name} 파일 선택됨: {file_path}")
    selected_files[platform_name] = file_path

def log_message(message):
  """로그 메시지를 텍스트 영역에 출력"""
  text_area.insert(tk.END, message + "\n")
  text_area.see(tk.END)

def reset_program():
  """프로그램 초기화"""
  global selected_files
  selected_files.clear()
  text_area.delete('1.0', tk.END)
  log_message("프로그램이 초기화되었습니다.")

def run_program(service):
  """전체 데이터 처리 및 구글 시트 반영, 스타일링 적용을 별도 스레드에서 실행"""
  def task():
    log_message("프로그램 실행 시작...")
    for platform_name, file_path in selected_files.items():
      if platform_name == "쿠팡":
        data = process_coupang(file_path)
      elif platform_name == "스마트스토어":
        data = process_smartstore(file_path)
      elif platform_name == "ESM":
        data = process_esm(file_path)
      elif platform_name == "11번가":
        data = process_st11(file_path)
      else:
        continue

      enriched_data = enrich_data_with_fas(data, service)
      append_new_orders(service, enriched_data, platform_name)
    log_message("모든 작업 완료!")

  threading.Thread(target=task).start()

def open_file_through_drag_and_drop(event):
  """파일을 드래그앤드롭으로 열 때 자동으로 마켓별로 파일을 분류하여 selected_files에 저장"""
  file_path_string = event.data
  if file_path_string.startswith('{') and file_path_string.endswith('}'):
    file_path_string = file_path_string[1:-1]
  file_paths = root.tk.splitlist(file_path_string)
  for file_path in file_paths:
    if "DeliveryList" in file_path:
      selected_files["쿠팡"] = file_path
      log_message(f"쿠팡 파일 선택됨: {file_path}")
    elif "스마트스토어" in file_path:
      selected_files["스마트스토어"] = file_path
      log_message(f"스마트스토어 파일 선택됨: {file_path}")
    elif "발송관리" in file_path or "신규주문" in file_path:
      selected_files["ESM"] = file_path
      log_message(f"ESM 파일 선택됨: {file_path}")
    elif "logistics" in file_path:
      selected_files["11번가"] = file_path
      log_message(f"11번가 파일 선택됨: {file_path}")

def create_main_window():
  """메인 윈도우와 모든 위젯을 생성 및 배치하고, 드래그앤드롭 이벤트를 등록"""
  global root, text_area, selected_files
  root = TkinterDnD.Tk()
  root.title("order_combiner")

  selected_files = {}

  frame = tk.Frame(root)
  frame.pack(pady=10)

  btn_coupang = tk.Button(frame, text="쿠팡 파일 찾아보기", command=lambda: select_file("쿠팡"))
  btn_coupang.grid(row=0, column=0, padx=5)

  btn_smartstore = tk.Button(frame, text="스마트스토어 파일 찾아보기", command=lambda: select_file("스마트스토어"))
  btn_smartstore.grid(row=0, column=1, padx=5)

  btn_esm = tk.Button(frame, text="ESM 파일 찾아보기", command=lambda: select_file("ESM"))
  btn_esm.grid(row=0, column=2, padx=5)

  btn_st11 = tk.Button(frame, text="11번가 파일 찾아보기", command=lambda: select_file("11번가"))
  btn_st11.grid(row=0, column=3, padx=5)

  btn_frame = tk.Frame(root)
  btn_frame.pack(pady=10)

  btn_run = tk.Button(btn_frame, text="실행", command=lambda: run_program(service))
  btn_run.pack(side=tk.LEFT, padx=5)

  btn_reset = tk.Button(btn_frame, text="초기화", command=reset_program)
  btn_reset.pack(side=tk.LEFT, padx=5)

  text_area = scrolledtext.ScrolledText(root, width=60, height=20)
  text_area.pack(pady=10)

  root.drop_target_register(DND_FILES)
  root.dnd_bind('<<Drop>>', open_file_through_drag_and_drop)

  return root, text_area

########################################################################################################################

if __name__ == "__main__":
  # Google Sheets API 인증 설정
  SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
  current_dir = os.path.dirname(__file__)
  json_file_path = os.path.join(current_dir, 'json_files/jason-455001-adc816559e51.json')
  credentials = Credentials.from_service_account_file(json_file_path, scopes=SCOPES)
  service = build('sheets', 'v4', credentials=credentials)

  root, text_area = create_main_window()
  root.mainloop()
