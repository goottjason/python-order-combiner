import msoffcrypto
import pandas as pd
from io import BytesIO
import os


def get_smart_data(file_path):

	password = "1234"

	# 복호화된 파일을 메모리에 로드
	with open(file_path, "rb") as f:
		decrypted_file = BytesIO()
		office_file = msoffcrypto.OfficeFile(f)
		office_file.load_key(password=password)  # 비밀번호 설정
		office_file.decrypt(decrypted_file)

	# pandas로 복호화된 파일 읽기
	decrypted_file.seek(0)  # 메모리 스트림의 시작으로 이동
	delivery_data = pd.read_excel(decrypted_file, header=None)
	# 둘째 행을 키로 설정하여 DataFrame 생성
	delivery_data.columns = delivery_data.iloc[1]  # 둘째 행을 열 이름으로 설정
	delivery_data = delivery_data[2:]  # 첫 두 행 제외 (데이터만 남김)

	list_of_dicts = delivery_data.to_dict(orient="records")

	# 데이터 변경 및 추가 작업
	updated_list = []
	for item in list_of_dicts:
		updated_item = item.copy()
		# 1. '결제일' key값의 value값을 'YYYY. MM. DD' 형식으로 변경
		if '결제일' in updated_item:
			updated_item['결제일'] = pd.to_datetime(updated_item['결제일']).strftime('%Y. %m. %d')
		# 2. 각 딕셔너리에 key값을 '샵', value값을 '쿠팡'으로 추가
		updated_item['샵'] = '스마트스토어'
		updated_item['조회'] = '조회'
		updated_item['해외비용'] = '원화'
		# 3. '고객명' 추가 (조건에 따라)
		if '수취인명' in updated_item and '구매자명' in updated_item:
			if updated_item['수취인명'] != updated_item['구매자명']:
				updated_item['고객명'] = f"{updated_item['수취인명']}({updated_item['구매자명']})"
			else:
				updated_item['고객명'] = updated_item['수취인명']
		# 4. 기존 key값 변경 (새로운 이름으로 매핑)
		key_mapping = {
			'결제일': '주문일',
			'개인통관고유부호': '통관번호',
			'수취인연락처1': '휴대폰',
			'우편번호': 'POST',
			'통합배송지': '주소',
			'배송메세지': '메세지',
			'정산예정금액': '순수익'
		}
		for old_key, new_key in key_mapping.items():
			if old_key in updated_item:
				updated_item[new_key] = updated_item.pop(old_key)  # 기존 키를 새 키로 변경
		updated_item['통관검증'] = f"=\"{updated_item['고객명']}\"&\"/\"&\"{updated_item['통관번호']}\"&\"/\"&\"{updated_item['휴대폰']}\""
		# 업데이트된 항목 추가
		updated_list.append(updated_item)
	
	return updated_list

if __name__ == "__main__":
	# 직접 실행 시에만 출력
	list_of_dicts = get_smart_data()
	# print(list_of_dicts)