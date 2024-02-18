<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="5">가격비교 매칭불가 상품 사유</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>항목</td>
	<td>내용</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="4">가격 어뷰징</td>
	<td align="left">동일 상품군의 평균 가격에 비해 가격이 너무 낮거나, 고의적인 가격 어뷰징으로 확인될 경우</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left">색상 및 디자인별 추가금 있는 상품</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left">지식쇼핑 상품정보 내 배송비 표기가 되어있으나, 쇼핑몰 페이지 배송비가 미표기된 경우</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left">지식쇼핑 가격비교 페이지 상위 노출을 목적으로 배송비를 높이고 상품 가격을 낮추는 행위</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="2">배송비 오류</td>
	<td align="left">지식쇼핑 상품정보 내 배송비와 쇼핑몰 페이지 배송비가 상이할 경우 (*조건부 무료배송 오류 상품 포함</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left">구매수량 별 차등 배송비를 적용하는 경우 (*단품 구매시 배송비 오류 상품일 경우)</td>
</tr>

<tr align="center" bgcolor="#FFFFFF">
	<td>중고/반품/렌탈</td>
	<td align="left">중고/반품/전시/스크래치 상품, (*리퍼비시 상품명 맨 앞에 [중고] 키워드를 표기하고 판매하는 경우)</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>해외상품</td>
	<td align="left">[해외],해외쇼핑,구매대행,OO직배송,글로벌셀러,글로벌쇼핑, 글로벌팩토리 등 해외배송 상품</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>옵션 추가금</td>
	<td align="left">옵션 선택 상품이면서 해당 상품 선택 시 추가금이 발생하는 경우</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>옵션 품절</td>
	<td align="left">옵션 선택 상품이면서 해당 상품 선택 시 품절인 경우</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>상품명 어뷰징</td>
	<td align="left">고의적으로 상품명의 일부만 수정하여 동일상품을 대량 등록하는 경우</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>기타</td>
	<td align="left">쇼핑몰 페이지 접근 시 19금 인증 페이지 노출되는 경우</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

