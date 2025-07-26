import pandas as pd
import math
import os

def get_esm_data(file_path):

  delivery_data = pd.read_excel(file_path, sheet_name='Sheet 1')
  
  # 데이터를 딕셔너리 리스트로 변환
  list_of_dicts = delivery_data.to_dict(orient='records')
  # 데이터 변경 및 추가 작업
  updated_list = []

  for item in list_of_dicts:
    # 기존 데이터 복사 (기존 키와 값을 유지)
    updated_item = item.copy()
    if '주문번호' in updated_item:
      updated_item['주문번호'] = str(updated_item['주문번호']).split('.')[0]  # .0 제거
    #if any('주문일' in key for key in updated_item):
    for key in updated_item:
      if '주문일' in key:
        updated_item['주문일'] = pd.to_datetime(updated_item[key]).strftime('%Y. %m. %d')
        break

    # 2. 각 딕셔너리에 key값을 '샵', value값을 '쿠팡'으로 추가
    if '판매아이디' in updated_item:
      if updated_item['판매아이디'] == '지마켓(younzara)':
        updated_item['샵'] = 'G마켓'
      else:
        updated_item['샵'] = '옥션'
    updated_item['조회'] = '조회'
    updated_item['해외비용'] = '원화'
    # 3. '고객명' 추가 (조건에 따라)
    if '수령인명' in updated_item and '구매자명' in updated_item:
      if updated_item['수령인명'] != updated_item['구매자명']:
        updated_item['고객명'] = f"{updated_item['수령인명']}({updated_item['구매자명']})"
      else:
        updated_item['고객명'] = updated_item['수령인명']
    # 4. 기존 key값 변경 (새로운 이름으로 매핑)
    key_mapping = {
      '수령인 통관정보': '통관번호',
      '수령인 휴대폰': '휴대폰',
      '우편번호': 'POST',
      '배송시 요구사항': '메세지',
      '정산예정금액': '순수익'
    }
    for old_key, new_key in key_mapping.items():
      if old_key in updated_item:
        updated_item[new_key] = updated_item.pop(old_key)  # 기존 키를 새 키로 변경
    
    updated_item['통관검증'] = f"=\"{updated_item['고객명']}\"&\"/\"&\"{updated_item['통관번호']}\"&\"/\"&\"{updated_item['휴대폰']}\""
    # 업데이트된 항목 추가
    updated_list.append(updated_item)

  # 데이터를 딕셔너리 리스트로 변환 (첫 행을 키로 사용)
  return updated_list

if __name__ == "__main__":
  # 직접 실행 시에만 출력
  list_of_dicts = get_esm_data()
  # print(list_of_dicts)