Attribute VB_Name = "selectMallMoudle"
Option Explicit

Sub selectMall(ByRef 입점사 As String, ByRef 조건 As Variant)

Dim 이공홈조건 As Variant, 크공홈조건 As Variant, 스스조건 As Variant, 무신사조건 As Variant, 아몬즈조건 As Variant, 겟트조건 As Variant, 이십구조건 As Variant, 하이버조건 As Variant, oco조건 As Variant, 루앱조건 As Variant, gvg조건 As Variant, w컨셉조건 As Variant, 하고조건 As Variant

크공홈조건 = Array("주문번호", "자체 상품코드", "옵션정보", "수취인명", "수취인 연락처", "주문자 연락처", "주소", "배송메세지", "수량", "상품별 금액", "배송비 합계", "브랜드")
이공홈조건 = Array("주문번호", "상품명", "옵션정보", "수취인명", "수취인 연락처", "주문자 연락처", "주소", "배송메세지", "수량", "상품별 금액", "배송비 합계", "브랜드")
스스조건 = Array("상품주문번호", "옵션관리코드", "수량", "수취인명", "수취인연락처1", "수취인연락처2", "통합배송지", "배송메세지", "수량", "상품별 총 주문금액", "배송비 합계")
무신사조건 = Array("주문일련번호", "상품명", "옵션", "수령자", "핸드폰", "전화번호", "주소", "특이사항", "주문수량", "판매가", "입금일시", "업체")
이십구조건 = Array("주문번호", "업체상품명", "옵션명", "수령인", "수령자 연락처", "주문자 연락처", "수령자 주소", "배송요청사항", "수량", "판매가 합계", "출고연기사유", "브랜드")
w컨셉조건 = Array("주문번호", "상품명", "옵션1", "수취인", "수취인연락처1", "수취인연락처2", "배송지", "배송메모", "수량", "판매가", "주문일자")
하고조건 = Array("주문번호", "상품명", "옵션", "수취인", "수취인 전화번호", "수취인 휴대폰 번호", "배송지주소", "배송메세지", "수량", "판매가", "배송 지연일시")
아몬즈조건 = Array("주문번호", "상품명", "옵션정보", "수취인명", "구매자 연락처", "수취인 연락처", "배송지", "배송메시지", "수량", "상품 가격(정가)", "결제 일시")
루앱조건 = Array("주문번호", "상품 영문명", "상품옵션", "수취인 이름", "수취인 전화번호", "주문자 전화번호", "주소", "배송 메모", "수량", "현 판매단가", "주문일자")


With ActiveWorkbook
Select Case True
Case .Name Like "*무신사*.xls": 조건 = 무신사조건: 입점사 = "무신사"
Case .Name Like "*스스*.xlsx": 조건 = 스스조건: 입점사 = "스스"
Case .Name Like "*크공홈*.xls*": 조건 = 크공홈조건: 입점사 = "공홈"
Case .Name Like "*이공홈*.xls*": 조건 = 이공홈조건: 입점사 = "공홈" : Call handle29cm
Case .Name Like "*29cm*.xls*": 조건 = 이십구조건: 입점사 = "29cm"
Case .Name Like "*컨셉*.xlsx": 조건 = w컨셉조건: 입점사 = "w컨셉"
Case .Name Like "*하고*.xls*": 조건 = 하고조건: 입점사 = "하고"
Case .Name Like "*아몬즈*.xls*": 조건 = 아몬즈조건: 입점사 = "아몬즈"
Case .Name Like "*루앱*.csv*": 조건 = 루앱조건: 입점사 = "루앱" : Call handleLuaeb
Case Else: 입점사 = "X"
End Select
End With



End Sub