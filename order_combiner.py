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
from get_coupang_data import get_coupang_data  # 쿠팡 데이터 처리 함수
from get_smart_data import get_smart_data  # 스마트스토어 데이터 처리 함수
from get_esm_data import get_esm_data  # ESM 데이터 처리 함수
from get_street_data import get_st11_data  # 11번가 데이터 처리 함수
from tkinterdnd2 import TkinterDnD, DND_FILES  # 드래그앤드롭 기능을 위한 모듈

#####################################################################################################################
# 함수 정의
#####################################################################################################################


"""
  데이터 리스트의 각 항목에서 NaN 값을 빈 문자열로 변환하고, 문자열로 변환 후 공백 제거
  Args:
    data (list): 비정제된 딕셔너리형 리스트
  Returns:
    cleaned_data (list): 정제된 딕셔너리형 리스트
"""
def clean_data(data):
  cleaned_data = []
  for item in data:  # 각 딕셔너리 순회
    cleaned_item = {}
    for key, value in item.items(): # 딕셔너리의 각 (키, 값) 쌍을 튜플로 묶어 반복할 수 있게 해주는 메서드
      if pd.isna(value):  # NaN이면 true (빈 문자열('')이나 numpy.inf(무한대)는 결측치로 간주하지 않음)
        cleaned_item[key] = ''  # value를 빈 문자열('')로 대체
      else:
        cleaned_item[key] = str(value).strip()  # 문자열로 변환 후 양쪽 공백 제거
    cleaned_data.append(cleaned_item)  # 정제된 항목을 리스트에 추가
  return cleaned_data  # 정제된 데이터 리스트 반환

"""
  Google Sheets의 FAS 시트에서 데이터를 읽어와 딕셔너리 리스트로 반환
  Returns:
    fas_dict_list (list): FAS 시트의 각 행을 딕셔너리로 변환한 리스트
"""
def get_fas_data():
  SPREADSHEET_ID = '18_Uj7-18Iosw8I2BbTVSWhUodpbQBbVIdGIOrlVwqG8'  # FAS 데이터가 있는 스프레드시트 ID
  RANGE_NAME = 'FAS'  # 읽어올 시트 이름
  result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=RANGE_NAME
  ).execute()  # Google Sheets API를 통해 데이터 읽기
  values = result.get('values', [])  # 읽어온 값 중 'values' 키의 값 추출
  if not values:  # 데이터가 없으면
    log_message("FAS 시트에서 데이터를 찾을 수 없습니다.")  # 로그 출력
    return []  # 빈 리스트 반환
  keys = values[0]  # 첫 번째 행을 키(컬럼명)로 사용
  data_rows = values[1:]  # 두 번째 행부터 실제 데이터 행
  fas_dict_list = [dict(zip(keys, row)) for row in data_rows]  # 각 행을 딕셔너리로 변환해 리스트 생성
  return fas_dict_list  # 변환된 리스트 반환

"""
  FAS 데이터와 매칭하여 입력 데이터에 영문명, 링크, 상품코드 등 추가 및 가공
  Args:
    data (list): 원본 데이터(딕셔너리 리스트)
  Returns:
    enriched_data (list): FAS 정보가 추가된 데이터 리스트
"""
def enrich_data_with_fas(data):
  time.sleep(1)  # API 호출 간격을 위해 1초 대기
  fas_data = get_fas_data()  # FAS 데이터 불러오기
  fas_lookup = {item['korName']: item for item in fas_data}  # 한글명 기준으로 FAS 데이터 딕셔너리 생성

  enriched_data = []  # 가공된 데이터를 저장할 리스트

  for idx, item in enumerate(data, start=1):  # start=1: 1-based 인덱스
    enriched_item = item.copy()  # 원본 항목 복사
    product_name = item.get('상품명')  # 상품명 추출
    if product_name in fas_lookup:  # FAS 데이터에 상품명이 있으면
      fas_item = fas_lookup[product_name]  # 해당 FAS 항목 추출
      engName = fas_item.get('engName', '')  # 영문명 추출
      link = fas_item.get('link', '')  # 링크 추출
      # HYPERLINK 수식 추가 (Q열 기준, ROW_PLACEHOLDER는 나중에 실제 행 번호로 교체)
      enriched_item['조회'] = f'=HYPERLINK(IF(LEFT(QROW_PLACEHOLDER,1)="6", ' \
                            f'"https://search.daum.net/search?nil_suggest=btn&w=tot&DA=SBC&q=우체국택배조회"&QROW_PLACEHOLDER, ' \
                            f'"https://search.daum.net/search?w=tot&DA=YZR&t__nil_searchbox=btn&q=cj택배조회"&QROW_PLACEHOLDER), "조회")'
      # 총비용, 순손익 수식 추가 (V, W, T, X열 기준)
      enriched_item['총비용'] = '=VROW_PLACEHOLDER+WROW_PLACEHOLDER'
      enriched_item['순손익'] = '=TROW_PLACEHOLDER-XROW_PLACEHOLDER'

      enriched_item['영문명'] = f'=HYPERLINK("{link}", "{engName}")'  # 하이퍼링크 형식으로 영문명 추가
      packQty = fas_item.get('packQty', '')  # 포장수량 추출
      enriched_item['수량'] = int(packQty) * int(item.get('수량'))  # 수량 계산 및 추가
      if int(item.get('수량')) >= 2:  # 수량이 2개 이상이면
        enriched_item['구매정보'] = '⠀'  # 구매정보에 특수문자 추가 (스타일링할때 true 느낌...)
      if item.get('메세지') == '문 앞':  # 메세지가 '문 앞'이면
        enriched_item['메세지'] = '문 앞에 놓아주세요.'  # 메세지 문구 수정
      enriched_item['상품코드'] = fas_item.get('sbCode', '')  # 상품코드 추가
      enriched_item['배송현황'] = '미구입'
      enriched_item['구매비'] = 0
      enriched_item['배송비'] = 0
      # if len(enriched_item['POST']) == 4:
      #   enriched_item['POST'] = '0' + enriched_item['POST']

    else:
      log_message(f"FAS 데이터에서 '{product_name}'에 해당하는 항목을 찾을 수 없습니다.")  # 매칭 실패 시 로그 출력
    enriched_data.append(enriched_item)  # 가공된 항목을 리스트에 추가
  return enriched_data  # 가공된 데이터 리스트 반환

"""
  구글 시트에서 데이터를 읽어와 딕셔너리 리스트로 반환
  Returns:
    dict_list (list): 시트의 각 행을 딕셔너리로 변환한 리스트
"""
def get_data_as_dict_list():
  service = build('sheets', 'v4', credentials=credentials)  # 구글 시트 서비스 객체 생성
  sheet = service.spreadsheets()  # 시트 객체 생성
  result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()  # 데이터 읽기
  values = result.get('values', [])  # 읽어온 값 추출
  if not values:  # 데이터가 없으면
    log_message("No data found.")  # 로그 출력
    return []  # 빈 리스트 반환
  keys = values[0]  # 첫 번째 행을 키(컬럼명)로 사용
  data_rows = values[1:]  # 두 번째 행부터 실제 데이터 행
  dict_list = [dict(zip(keys, row)) for row in data_rows]  # 각 행을 딕셔너리로 변환해 리스트 생성
  return dict_list  # 변환된 리스트 반환

"""
  구글 시트의 특정 셀 값을 업데이트
  Args:
    row (int): 행 번호
    col (str): 열 알파벳
    value (str): 입력할 값
  Returns:
    result: API 응답 결과
"""
def update_cell(row, col, value):
  service = build('sheets', 'v4', credentials=credentials)  # 구글 시트 서비스 객체 생성
  sheet = service.spreadsheets()  # 시트 객체 생성
  range_to_update = f'{RANGE_NAME}!{col}{row}'  # 업데이트할 셀 범위 지정
  body = {
    'values': [[value]]  # 업데이트할 값
  }
  result = sheet.values().update(
    spreadsheetId=SPREADSHEET_ID,
    range=range_to_update,
    valueInputOption='RAW',
    body=body
  ).execute()  # 셀 값 업데이트 실행
  log_message(f"Updated cell {col}{row} with value '{value}'")  # 로그 출력
  return result  # 결과 반환

"""
  구글 시트에서 이미 존재하는 주문번호 집합 반환
  Returns:
    existing_ids (set): 주문번호 집합
"""
def get_existing_order_ids():
  result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=RANGE_NAME
  ).execute()  # 시트 데이터 읽기
  values = result.get('values', [])  # 값 추출
  if not values:  # 데이터 없으면
    return set()  # 빈 집합 반환
  headers = values[0]  # 헤더 추출
  try:
    order_id_index = headers.index('주문번호')  # '주문번호' 컬럼 인덱스 찾기
  except ValueError:
    raise Exception("시트에 '주문번호' 컬럼이 존재하지 않습니다.")  # 없으면 예외 발생
  existing_ids = set()  # 주문번호 저장 집합
  for row in values[1:]:  # 데이터 행 반복
    if len(row) > order_id_index and row[order_id_index].strip():  # 주문번호가 있으면
      existing_ids.add(row[order_id_index].strip())  # 집합에 추가
  return existing_ids  # 집합 반환

"""
  중복되지 않은 신규 주문 데이터만 구글 시트에 추가
  Args:
    dict_list (list): 추가할 데이터 리스트
    platform_name (str): 플랫폼 이름(로그용)
"""
def append_new_orders(dict_list, platform_name):
  time.sleep(1)  # API 호출 간격 조절
  existing_ids = get_existing_order_ids()  # 기존 주문번호 집합 가져오기
  new_data = []  # 신규 데이터 저장 리스트
  for item in dict_list:  # 각 데이터 반복
    order_id = str(item.get('주문번호', '')).strip()  # 주문번호 추출
    if not order_id:  # 주문번호 없으면
      log_message(f"[{platform_name}] 주문번호가 없는 데이터: {item}")  # 로그 출력
      continue  # 건너뜀
    if order_id in existing_ids:  # 이미 존재하면
      log_message(f"[{platform_name}] 중복된 주문번호: {order_id}")  # 로그 출력
      continue  # 건너뜀
    new_data.append(item)  # 신규 데이터로 추가
  if not new_data:  # 추가할 데이터 없으면
    log_message(f"[{platform_name}] 추가할 데이터가 없습니다.")  # 로그 출력
    return  # 함수 종료

  new_data = clean_data(new_data)  # 데이터 정제
  headers = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=f"{RANGE_NAME}!1:1"
  ).execute().get('values', [[]])[0]  # 헤더 추출

  # '배송현황' 컬럼 인덱스 찾기
  delivery_status_index = headers.index('배송현황') if '배송현황' in headers else None

  rows = []  # 추가할 행 리스트
  for item in new_data:  # 신규 데이터 반복
    row = [item.get(header, '') for header in headers]  # 헤더 순서대로 값 추출
    rows.append(row)  # 행 추가
  body = {'values': rows}  # 추가할 값
  result = service.spreadsheets().values().append(
    spreadsheetId=SPREADSHEET_ID,
    range=RANGE_NAME,
    valueInputOption='USER_ENTERED',
    insertDataOption='INSERT_ROWS',
    body=body
  ).execute()  # 시트에 데이터 추가
  log_message(f"[{platform_name}] {len(rows)}개 데이터 추가 완료!")  # 로그 출력

  # [추가] 새로 추가된 행의 '배송현황' 컬럼 회색으로 설정
  if delivery_status_index is not None:
    import re

    # 1. 업데이트된 범위(updatedRange)에서 행 번호 추출
    updated_range = result.get('updates', {}).get('updatedRange', '')
    if updated_range:
      try:
        range_part = updated_range.split('!')[1]
        start_cell, end_cell = range_part.split(':')
        start_row = int(re.search(r'\d+$', start_cell).group())
        end_row = int(re.search(r'\d+$', end_cell).group())

        # 컬럼 문자 매핑
        col_map = {}
        for col_name in ['국내운송장', '구매비', '배송비', '순수익', '총비용', '순손익', '조회']:
          if col_name in headers:
            idx = headers.index(col_name)
            col_map[col_name] = chr(65 + idx)

        sheet_id = get_sheet_id(RANGE_NAME.split('!')[0] if '!' in RANGE_NAME else RANGE_NAME)
        requests = []
        for row in range(start_row, end_row + 1):
          # 조회
          if '조회' in headers and '국내운송장' in col_map:
            requests.append({
              "updateCells": {
                "range": {
                  "sheetId": sheet_id,
                  "startRowIndex": row - 1,
                  "endRowIndex": row,
                  "startColumnIndex": headers.index('조회'),
                  "endColumnIndex": headers.index('조회') + 1
                },
                "rows": [{
                  "values": [{
                    "userEnteredValue": {
                      "formulaValue": f'=HYPERLINK(IF(LEFT({col_map["국내운송장"]}{row},1)="6", "https://search.daum.net/search?nil_suggest=btn&w=tot&DA=SBC&q=우체국택배조회"&{col_map["국내운송장"]}{row}, "https://search.daum.net/search?w=tot&DA=YZR&t__nil_searchbox=btn&q=cj택배조회"&{col_map["국내운송장"]}{row}), "조회")'
                    }
                  }]
                }],
                "fields": "userEnteredValue"
              }
            })
          # 총비용
          if '총비용' in headers and '구매비' in col_map and '배송비' in col_map:
            requests.append({
              "updateCells": {
                "range": {
                  "sheetId": sheet_id,
                  "startRowIndex": row - 1,
                  "endRowIndex": row,
                  "startColumnIndex": headers.index('총비용'),
                  "endColumnIndex": headers.index('총비용') + 1
                },
                "rows": [{
                  "values": [{
                    "userEnteredValue": {
                      "formulaValue": f'={col_map["구매비"]}{row}+{col_map["배송비"]}{row}'
                    }
                  }]
                }],
                "fields": "userEnteredValue"
              }
            })
          # 순손익
          if '순손익' in headers and '순수익' in col_map and '총비용' in col_map:
            requests.append({
              "updateCells": {
                "range": {
                  "sheetId": sheet_id,
                  "startRowIndex": row - 1,
                  "endRowIndex": row,
                  "startColumnIndex": headers.index('순손익'),
                  "endColumnIndex": headers.index('순손익') + 1
                },
                "rows": [{
                  "values": [{
                    "userEnteredValue": {
                      "formulaValue": f'={col_map["순수익"]}{row}-{col_map["총비용"]}{row}'
                    }
                  }]
                }],
                "fields": "userEnteredValue"
              }
            })
        if requests:
          service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests}
          ).execute()
      except Exception as e:
        log_message(f"수식 업데이트 오류: {str(e)}")

      try:
        # 예: updated_range = "order!A11:C13" → start_row=11, end_row=13
        range_part = updated_range.split('!')[1]
        start_cell, end_cell = range_part.split(':')
        start_row = int(re.search(r'\d+$', start_cell).group())
        end_row = int(re.search(r'\d+$', end_cell).group())

        # 2. 시트 ID 가져오기
        sheet_id = get_sheet_id(RANGE_NAME.split('!')[0] if '!' in RANGE_NAME else RANGE_NAME)

        # 3. 스타일 요청 생성
        requests = []
        for row in range(start_row, end_row + 1):  # end_row 포함
          requests.append({
            "updateCells": {
              "range": {
                "sheetId": sheet_id,
                "startRowIndex": row - 1,  # 0-based
                "endRowIndex": row,
                "startColumnIndex": delivery_status_index,
                "endColumnIndex": delivery_status_index + 1
              },
              "rows": [{
                "values": [{
                  "userEnteredFormat": {
                    "backgroundColor": {
                      "red": 0.8,
                      "green": 0.8,
                      "blue": 0.8
                    }
                  }
                }]
              }],
              "fields": "userEnteredFormat.backgroundColor"
            }
          })

        # 4. 스타일 일괄 적용
        if requests:
          service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests}
          ).execute()
          log_message(f"[{platform_name}] 배송현황 컬럼 스타일 적용 완료!")

      except Exception as e:
        log_message(f"배송현황 스타일 적용 오류: {str(e)}")

def reset_program():
  # 선택된 파일 정보 초기화
  global selected_files
  selected_files.clear()
  # 로그 영역 초기화
  text_area.delete('1.0', tk.END)
  # (필요하다면 추가로 내부 변수, 임시 데이터 등도 여기서 초기화)
  log_message("프로그램이 초기화되었습니다.")


"""
  시트 이름으로 sheetId 반환
  Args:
    sheet_name (str): 시트 이름
  Returns:
    sheet_id (int): 시트 ID
"""
def get_sheet_id(sheet_name):
  spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()  # 스프레드시트 정보 가져오기
  sheets = spreadsheet.get('sheets', [])  # 모든 시트 정보 추출
  for sheet in sheets:  # 각 시트 반복
    if sheet['properties']['title'] == sheet_name:  # 이름 일치하면
      return sheet['properties']['sheetId']  # sheetId 반환
  raise Exception(f"Sheet '{sheet_name}' not found.")  # 없으면 예외

"""
  구글 시트의 데이터에 스타일(테두리, 색상, 서식 등) 적용
"""
def apply_styling():
  sheet_id = get_sheet_id(RANGE_NAME)  # 시트 ID 가져오기 ('order')
  sheet = service.spreadsheets()  # 시트 객체 생성
  result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()  # 데이터 읽기
  values = result.get('values', [])  # 값 추출
  if not values:  # 데이터 없으면
    log_message("No data found.")  # 로그 출력
    return  # 함수 종료
  headers = values[0]  # 헤더 추출
  # 각 컬럼 인덱스 추출
  order_date_index = headers.index('주문일')
  shop_index = headers.index('샵')
  order_number_index = headers.index('주문번호')
  post_index = headers.index('POST')
  순수익_index = headers.index('순수익')
  구매비_index = headers.index('구매비')
  배송비_index = headers.index('배송비')
  총비용_index = headers.index('총비용')
  순손익_index = headers.index('순손익')
  수량_index = headers.index('수량')
  구매정보_index = headers.index('구매정보')
  requests = []  # 스타일 적용 요청 리스트

  # 전체 테두리 스타일 적용
  requests.append({
    "updateBorders": {
      "range": {
        "sheetId": sheet_id,
        "startRowIndex": 0,
        "endRowIndex": len(values),
        "startColumnIndex": 0,
        "endColumnIndex": len(headers)
      },
      "top": {"style": "SOLID", "width": "1", "color": {"red": 0, "green": 0, "blue": 0}},
      "bottom": {"style": "SOLID", "width": "1", "color": {"red": 0, "green": 0, "blue": 0}},
      "left": {"style": "SOLID", "width": "1", "color": {"red": 0, "green": 0, "blue": 0}},
      "right": {"style": "SOLID", "width": "1", "color": {"red": 0, "green": 0, "blue": 0}},
      "innerHorizontal": {"style": "SOLID", "width": "1", "color": {"red": 0, "green": 0, "blue": 0}},
      "innerVertical": {"style": "SOLID", "width": "1", "color": {"red": 0, "green": 0, "blue": 0}}
    }
  })

  # POST 컬럼 텍스트 서식 적용
  if post_index is not None:
    for row_index in range(1, len(values)):
      requests.append({
        "updateCells": {
          "range": {
            "sheetId": sheet_id,
            "startRowIndex": row_index,
            "endRowIndex": row_index + 1,
            "startColumnIndex": post_index,
            "endColumnIndex": post_index + 1
          },
          "rows": [{"values": [{"userEnteredFormat": {"numberFormat": {"type": "TEXT"}}}]}],
          "fields": "userEnteredFormat.numberFormat"
        }
      })

  # 통화 서식 컬럼 적용
  currency_columns = {
    "순수익": 순수익_index,
    "구매비": 구매비_index,
    "배송비": 배송비_index,
    "총비용": 총비용_index,
    "순손익": 순손익_index,
  }
  for col_name, col_index in currency_columns.items():
    if col_index is not None:
      for row_index in range(1, len(values)):
        requests.append({
          "updateCells": {
            "range": {
              "sheetId": sheet_id,
              "startRowIndex": row_index,
              "endRowIndex": row_index + 1,
              "startColumnIndex": col_index,
              "endColumnIndex": col_index + 1
            },
            "rows": [{"values": [{"userEnteredFormat": {"numberFormat": {"type": "CURRENCY", "pattern": "₩#,##0"}}}]}],
            "fields": "userEnteredFormat.numberFormat"
          }
        })

  # 구매정보 특수문자 행의 수량 컬럼 볼드 처리
  if 구매정보_index is not None and 수량_index is not None:
    for row_index in range(1, len(values)):
      try:
        if values[row_index][구매정보_index].strip() == '⠀':
          requests.append({
            "updateCells": {
              "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_index,
                "endRowIndex": row_index + 1,
                "startColumnIndex": 수량_index,
                "endColumnIndex": 수량_index + 1
              },
              "rows": [{"values": [{"userEnteredFormat": {"textFormat": {"bold": True}}}]}],
              "fields": "userEnteredFormat.textFormat.bold"
            }
          })
      except IndexError as e:
        log_message(f"Error processing row {row_index + 1}: {e}")

  # 주문일, 샵별 배경색 적용
  for row_index, row in enumerate(values[1:], start=2):
    order_date_str = row[order_date_index]
    try:
      if len(order_date_str.split('.')) == 2:
        order_date_str = f"2025. {order_date_str.strip()}"
      order_date = datetime.datetime.strptime(order_date_str, '%Y. %m. %d')
      cell_color = "#fff2cc" if order_date.day % 2 != 0 else "#faff9f"
    except ValueError:
      log_message(f"Invalid date format: {order_date_str}")
      continue

    requests.append({
      "updateCells": {
        "range": {
          "sheetId": sheet_id,
          "startRowIndex": row_index - 1,
          "endRowIndex": row_index,
          "startColumnIndex": order_date_index,
          "endColumnIndex": order_date_index + 1
        },
        "rows": [{"values": [{"userEnteredFormat": {"backgroundColor": {
          "red": int(cell_color[1:3], 16) / 255,
          "green": int(cell_color[3:5], 16) / 255,
          "blue": int(cell_color[5:], 16) / 255
        }}}]}],
        "fields": "userEnteredFormat.backgroundColor"
      }
    })

    shop = row[shop_index]
    shop_colors = {
      "쿠팡": "#ff6699",
      "스마트스토어": "#a9d08e",
      "G마켓": "#00b050",
      "옥션": "#ffc000",
      "11번가": "#548dd4"
    }

    if shop in shop_colors:
      color = shop_colors[shop]
      for col_index in [shop_index, order_number_index]:
        requests.append({
          "updateCells": {
            "range": {
              "sheetId": sheet_id,
              "startRowIndex": row_index - 1,
              "endRowIndex": row_index,
              "startColumnIndex": col_index,
              "endColumnIndex": col_index + 1
            },
            "rows": [{"values": [{"userEnteredFormat": {"backgroundColor": {
              "red": int(color[1:3], 16) / 255,
              "green": int(color[3:5], 16) / 255,
              "blue": int(color[5:], 16) / 255
            }}}]}],
            "fields": "userEnteredFormat.backgroundColor"
          }
        })

  body = {'requests': requests}  # 스타일 적용 요청 바디
  sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()  # 스타일 일괄 적용
  log_message("Styling applied successfully.")  # 로그 출력

"""
  로그 메시지를 텍스트 영역에 출력
  Args:
    message (str): 출력할 메시지
"""
def log_message(message):
  text_area.insert(tk.END, message + "\n")  # 텍스트 영역에 메시지 추가
  text_area.see(tk.END)  # 스크롤을 맨 아래로 이동

"""
  쿠팡 파일 처리 함수
  Args:
    file_path (str): 파일 경로
  Returns:
    list: 처리된 데이터 리스트
"""
def process_coupang(file_path):
  log_message(f"쿠팡 파일 처리 중: {file_path}")  # 로그 출력
  return get_coupang_data(file_path)  # 쿠팡 데이터 처리 함수 호출

"""
  스마트스토어 파일 처리 함수
  Args:
    file_path (str): 파일 경로
  Returns:
    list: 처리된 데이터 리스트
"""
def process_smartstore(file_path):
  log_message(f"스마트스토어 파일 처리 중: {file_path}")  # 로그 출력
  return get_smart_data(file_path)  # 스마트스토어 데이터 처리 함수 호출

"""
  ESM 파일 처리 함수
  Args:
    file_path (str): 파일 경로
  Returns:
    list: 처리된 데이터 리스트
"""
def process_esm(file_path):
  log_message(f"ESM 파일 처리 중: {file_path}")  # 로그 출력
  return get_esm_data(file_path)  # ESM 데이터 처리 함수 호출

"""
  11번가 파일 처리 함수
  Args:
    file_path (str): 파일 경로
  Returns:
    list: 처리된 데이터 리스트
"""
def process_st11(file_path):
  log_message(f"11번가 파일 처리 중: {file_path}")  # 로그 출력
  return get_st11_data(file_path)  # 11번가 데이터 처리 함수 호출

"""
  파일 선택 다이얼로그를 열고 선택된 파일 경로를 selected_files 딕셔너리에 저장
  Args:
    platform_name (str): 마켓 이름
"""
def select_file(platform_name):

  file_path = filedialog.askopenfilename(
    title = f"{platform_name} 엑셀 파일 선택",
    filetypes = [("Excel files", "*.xlsx *.xls")]
  )  # 파일 선택 다이얼로그 표시
  if file_path:  # 파일이 선택되면
    log_message(f"{platform_name} 파일 선택됨: {file_path}")  # 로그 출력
    selected_files[platform_name] = file_path  # 파일 경로 저장

"""
  전체 데이터 처리 및 구글 시트 반영, 스타일링 적용을 별도 스레드에서 실행
"""
def run_program():
  def task():
    log_message("프로그램 실행 시작...")  # 로그 출력
    for platform_name, file_path in selected_files.items():  # 선택된 각 마켓별 파일 반복
      if platform_name == "쿠팡":
        # 1. 쿠팡 데이터 처리 함수 호출
        coupang_data = process_coupang(file_path)
        # 2. FAS에서 일치 상품 가져오고 각 데이터 입력
        enriched_coupang_data = enrich_data_with_fas(coupang_data)
        # 3.
        append_new_orders(enriched_coupang_data, "쿠팡")
      elif platform_name == "스마트스토어":
        smart_data = process_smartstore(file_path)
        enriched_smart_data = enrich_data_with_fas(smart_data)
        append_new_orders(enriched_smart_data, "스마트스토어")
      elif platform_name == "ESM":
        esm_data = process_esm(file_path)
        enriched_esm_data = enrich_data_with_fas(esm_data)
        append_new_orders(enriched_esm_data, "ESM")
      elif platform_name == "11번가":
        st11_data = process_st11(file_path)
        print(st11_data)
        enriched_st11_data = enrich_data_with_fas(st11_data)
        append_new_orders(enriched_st11_data, "11번가")
    log_message(f"{platform_name} 데이터 처리 완료.")  # 로그 출력
    log_message("스타일링 적용중")  # 로그 출력
    apply_styling()  # 스타일 적용
    log_message("모든 작업 완료!")  # 로그 출력

  threading.Thread(target=task).start()  # 별도 스레드에서 실행

"""
  파일을 드래그앤드롭으로 열 때 자동으로 마켓별로 파일을 분류하여 selected_files에 저장
  Args:
    event: 드롭 이벤트 객체
"""
def open_file_through_drag_and_drop(event):
  file_path_string = event.data  # 드롭된 파일 경로 문자열
  if file_path_string.startswith('{') and file_path_string.endswith('}'):  # 중괄호 제거
    file_path_string = file_path_string[1:-1]
  file_paths = root.tk.splitlist(file_path_string)  # 여러 파일 분리
  for file_path in file_paths:  # 각 파일 반복
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

# 1. 화면 생성
"""
  메인 윈도우와 모든 위젯을 생성 및 배치하고, 드래그앤드롭 이벤트를 등록합니다.
  Returns:
    root (Tk): 생성된 TkinterDnD 메인 윈도우 객체
    text_area (ScrolledText): 로그/결과 출력을 위한 텍스트 영역 위젯
"""
def create_main_window():

  root = TkinterDnD.Tk()  # TkinterDnD 기반 메인 윈도우 생성
  root.title("order_combiner")  # 윈도우 제목 설정

  global selected_files  # 전역 selected_files 딕셔너리 선언
  selected_files = {}  # 선택된 파일 경로 저장용 딕셔너리

  frame = tk.Frame(root)  # 버튼들을 담을 프레임 생성
  frame.pack(pady=10)  # 프레임 배치

  # 각 마켓별 파일 선택 버튼 생성 및 배치
  btn_coupang = tk.Button(frame, text="쿠팡 파일 찾아보기", command=lambda: select_file("쿠팡"))
  btn_coupang.grid(row=0, column=0, padx=5)

  btn_smartstore = tk.Button(frame, text="스마트스토어 파일 찾아보기", command=lambda: select_file("스마트스토어"))
  btn_smartstore.grid(row=0, column=1, padx=5)

  btn_esm = tk.Button(frame, text="ESM 파일 찾아보기", command=lambda: select_file("ESM"))
  btn_esm.grid(row=0, column=2, padx=5)

  btn_st11 = tk.Button(frame, text="11번가 파일 찾아보기", command=lambda: select_file("11번가"))
  btn_st11.grid(row=0, column=3, padx=5)

  # 실행 버튼과 초기화 버튼을 같은 행에 배치
  btn_frame = tk.Frame(root)
  btn_frame.pack(pady=10)

  btn_run = tk.Button(btn_frame, text="실행", command=run_program)
  btn_run.pack(side=tk.LEFT, padx=5)

  btn_reset = tk.Button(btn_frame, text="초기화", command=reset_program)
  btn_reset.pack(side=tk.LEFT, padx=5)

  text_area = scrolledtext.ScrolledText(root, width=60, height=20)  # 로그/결과 출력용 텍스트 영역 생성
  text_area.pack(pady=10)  # 텍스트 영역 배치

  root.drop_target_register(DND_FILES)  # 드래그앤드롭 기능 활성화
  root.dnd_bind('<<Drop>>', open_file_through_drag_and_drop)  # 드롭 이벤트 바인딩

  return root, text_area  # 생성된 윈도우와 텍스트 영역 반환

########################################################################################################################

if __name__ == "__main__":

  # Google Sheets API 인증 설정
  SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # 스프레드시트 접근 권한
  current_dir = os.path.dirname(__file__)  # 현재 파일 경로
  json_file_path = os.path.join(current_dir, 'json_files/jason-455001-adc816559e51.json')  # 인증키 경로
  SERVICE_ACCOUNT_FILE = json_file_path  # 인증키 파일명
  credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)  # 인증 객체 생성
  service = build('sheets', 'v4', credentials=credentials)  # 구글 시트 서비스 객체 생성

  print(service)
  SPREADSHEET_ID = '1WO__vUDFBnQNXSkKDsm3R-UeuyDMN1C6Dlo4anrkDEY'  # 사용할 스프레드시트 ID (order_list)
  RANGE_NAME = 'order'  # 사용할 시트 이름

  root, text_area = create_main_window()  # 메인 윈도우 및 텍스트 영역 생성
  root.mainloop()  # GUI 이벤트 루프 실행