import pandas as pd
import math
import os

def get_st11_data(file_path):
  delivery_data = pd.read_excel(file_path, sheet_name='Sheet', header=1)

  # 데이터를 딕셔너리 리스트로 변환
  list_of_dicts = delivery_data.to_dict(orient="records")

  # 데이터 변경 및 추가 작업
  updated_list = []
  for item in list_of_dicts:
    updated_item = item.copy()

    key_mapping = {
      '주문일시': '주문일',
      '세관신고정보': '통관번호',
      '휴대폰번호': '휴대폰',
      '우편번호': 'POST',
      '배송메시지': '메세지',
      '정산예정금액': '순수익',
      '판매자 상품코드': '상품코드'
    }
    for old_key, new_key in key_mapping.items():
      if old_key in updated_item:
        updated_item[new_key] = updated_item.pop(old_key)  # 기존 키를 새 키로 변경

    # 주문일 <- 주문일시
    if '주문일' in updated_item:
      updated_item['주문일'] = pd.to_datetime(updated_item['주문일']).strftime('%Y. %m. %d')
    # 샵
    updated_item['샵'] = '11번가'
    # 주문번호 (동일)
    # 고객명
    if '수취인' in updated_item and '구매자' in updated_item:
      if updated_item['수취인'] != updated_item['구매자']:
        updated_item['고객명'] = f"{updated_item['수취인']}({updated_item['구매자']})"
      else:
        updated_item['고객명'] = updated_item['수취인']
    # 통관번호 <- 세관신고정보
    # 휴대폰 <- 휴대폰번호
    # POST <- 우편번호
    # 주소 (동일)
    # 메세지 < - 배송메시지
    # 상품코드 <- 판매자 상품코드
    # 상품명 (동일)
    # 영문명 ■
    updated_item['영문명'] = ''
    # 수량 (동일)
    # 구매일
    updated_item['구매일'] = ''
    # 구매정보
    updated_item['구매정보'] = ''
    # 배송현황
    updated_item['배송현황'] = '미구입'
    # 국내운송장
    updated_item['국내운송장'] = ''
    # 유니패스
    updated_item['유니패스'] = ''
    # 조회
    updated_item['조회'] = '조회'
    # 순수익 <- 정산예정금액
    # 해외비용
    updated_item['해외비용'] = '원화'
    # 구매비
    updated_item['구매비'] = 0
    # 배송비
    updated_item['배송비'] = 0
    # 총비용 ■ FAS 조회 후 반영
    updated_item['총비용'] = 0
    # 순손익 ■ FAS 조회 후 반영
    updated_item['순손익'] = 0
    # 통관검증
    updated_item['통관검증'] = f'{updated_item["고객명"]}/{updated_item["통관번호"]}/{updated_item["휴대폰"]}'
    # 업데이트된 항목 추가
    updated_list.append(updated_item)

  return updated_list

if __name__ == "__main__":
  # 직접 실행 시에만 출력
  list_of_dicts = get_st11_data()
