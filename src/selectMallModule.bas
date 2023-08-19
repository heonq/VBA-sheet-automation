Attribute VB_Name = "selectMallMoudle"
Option Explicit

Sub selectMall(ByRef mall As String, ByRef mallOption As Variant)

Dim eiMallOption As Variant, crtsMallOption As Variant, smartStoreOption As Variant, musinsaOption As Variant, amondzOption As Variant, luaebOption As Variant, wconceptOption As Variant, hagoOption As Variant

crtsMallOption = Array("주문번호", "자체 상품코드", "옵션정보", "수취인명", "수취인 연락처", "주문자 연락처", "주소", "배송메세지", "수량", "상품별 금액", "배송비 합계", "브랜드")
eiMallOption = Array("주문번호", "상품명", "옵션정보", "수취인명", "수취인 연락처", "주문자 연락처", "주소", "배송메세지", "수량", "상품별 금액", "배송비 합계", "브랜드")
smartStoreOption = Array("상품주문번호", "옵션관리코드", "수량", "수취인명", "수취인연락처1", "수취인연락처2", "통합배송지", "배송메세지", "수량", "상품별 총 주문금액", "배송비 합계")
musinsaOption = Array("주문일련번호", "상품명", "옵션", "수령자", "핸드폰", "전화번호", "주소", "특이사항", "주문수량", "판매가", "입금일시", "업체")
twentynineOption = Array("주문번호", "업체상품명", "옵션명", "수령인", "수령자 연락처", "주문자 연락처", "수령자 주소", "배송요청사항", "수량", "판매가 합계", "출고연기사유", "브랜드")
wconceptOption = Array("주문번호", "상품명", "옵션1", "수취인", "수취인연락처1", "수취인연락처2", "배송지", "배송메모", "수량", "판매가", "주문일자")
hagoOption = Array("주문번호", "상품명", "옵션", "수취인", "수취인 전화번호", "수취인 휴대폰 번호", "배송지주소", "배송메세지", "수량", "판매가", "배송 지연일시")
amondzOption = Array("주문번호", "상품명", "옵션정보", "수취인명", "구매자 연락처", "수취인 연락처", "배송지", "배송메시지", "수량", "상품 가격(정가)", "결제 일시")
luaebOption = Array("주문번호", "상품 영문명", "상품옵션", "수취인 이름", "수취인 전화번호", "주문자 전화번호", "주소", "배송 메모", "수량", "현 판매단가", "주문일자")


With ActiveWorkbook
Select Case True
Case .Name Like "*무신사*.xls": mallOption = musinsaOption: mall = "무신사"
Case .Name Like "*스스*.xlsx": mallOption = smartStoreOption: mall = "스스"
Case .Name Like "*크공홈*.xls*": mallOption = crtsMallOption: mall = "공홈"
Case .Name Like "*이공홈*.xls*": mallOption = eiMallOption: mall = "공홈"
Case .Name Like "*29cm*.xls*": mallOption = twentynineOption: mall = "29cm"
Case .Name Like "*컨셉*.xlsx": mallOption = wconceptOption: mall = "w컨셉"
Case .Name Like "*하고*.xls*": mallOption = hagoOption: mall = "하고"
Case .Name Like "*아몬즈*.xls*": mallOption = amondzOption: mall = "아몬즈"
Case .Name Like "*루앱*.csv*": mallOption = luaebOption: mall = "루앱"
Case Else: mall = "X"
End Select
End With



End Sub