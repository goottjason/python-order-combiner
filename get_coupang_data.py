import pandas as pd
import math
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def get_coupang_data(file_path):
  delivery_data = pd.read_excel(file_path, sheet_name='Delivery')

  key_mapping = {
    '수취인이름': '수취인',
    '개인통관번호(PCCC)': '통관번호',
    '통관용수취인전화번호': '휴대폰',
    '우편번호': 'POST',
    '수취인 주소': '주소',
    '배송메세지': '메세지',
    '등록상품명': '상품명',
    '구매수(수량)': '수량',
    '결제액': '순수익'
  }
  delivery_data = delivery_data.rename(columns=key_mapping)

  delivery_data['샵'] = '쿠팡'
  delivery_data['상품코드'] = ''
  delivery_data['영문명'] = ''
  delivery_data['구매일'] = ''
  delivery_data['구매정보'] = ''
  delivery_data['배송현황'] = '미구입'
  delivery_data['국내운송장'] = ''
  delivery_data['유니패스'] = ''
  delivery_data['조회'] = '조회'
  delivery_data['해외비용'] = '원화'
  delivery_data['구매비'] = 0
  delivery_data['배송비'] = 0
  delivery_data['총비용'] = 0
  delivery_data['순손익'] = 0
  delivery_data['통관검증'] = ''

  # 데이터를 딕셔너리 리스트로 변환
  list_of_dicts = delivery_data.to_dict(orient='records')

  # 데이터 변경 및 추가 작업
  updated_list = []
  for item in list_of_dicts:
    # 기존 데이터 복사 (기존 키와 값을 유지)
    updated_item = item.copy()

    # 1. '주문일': 'YYYY. MM. DD' 형식으로 변경
    if '주문일' in updated_item:
      updated_item['주문일'] = pd.to_datetime(updated_item['주문일']).strftime('%Y. %m. %d')

    # 2. '고객명': '수취인(구매자)' 형식으로 변경 (동일하면 '수취인')
    if '수취인' in updated_item and '구매자' in updated_item:
      if updated_item['수취인'] != updated_item['구매자']:
        updated_item['고객명'] = f"{updated_item['수취인']}({updated_item['구매자']})"
      else:
        updated_item['고객명'] = updated_item['수취인']

    # 3. '순수익': 마켓수수료율을 차감
    if '순수익' in updated_item:
      updated_item['순수익'] = math.ceil(updated_item['순수익'] * 0.89)  # 소수점 올림
    # 4. '통관검증': '고객명/통관번호/휴대폰' 형식으로 변경
    if '통관검증' in updated_item:
      updated_item['통관검증'] = f'{updated_item["고객명"]}/{updated_item["통관번호"]}/{updated_item["휴대폰"]}'
    updated_list.append(updated_item)

  return updated_list

if __name__ == "__main__":
  list_of_dicts = get_coupang_data()