<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventOtherCls_wishlist.asp"-->

<!-- 엑셀파일로 저장 헤더 부분 -->
<%
Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=위시리스트이벤트.xls"

dim oeventuserlist, arrList, intLoop
set oeventuserlist = new CWishList
arrList = oeventuserlist.fnGetWishListExcel
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body>
<% IF isArray(arrList) THEN %>
	<table width="100%" border="1" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr bgcolor=#DDDDFF>
			<td align="center">아이디</td>
			<td align="center">회원등급</td>
			<td align="center">폴더번호</td>
			<td align="center">상품코드</td>
			<td align="center">정상가</td>
			<td align="center">상품명</td>
			<td align="center">브랜드명</td>
			<td align="center">카테고리</td>
		</tr>
		<% For intLoop =0 To UBound(arrList,2) %>
		<tr bgcolor=#FFFFFF>
			<td align="center"><%=arrList(0,intLoop)%></td>
			<td align="center"><%=UserGrade(arrList(1,intLoop))%></td>
			<td align="center"><%=arrList(2,intLoop)%></td>
			<td align="center"><%=arrList(3,intLoop)%></td>
			<td align="center"><%=FormatNumber(arrList(4,intLoop),0)%></td>
			<td align="center"><%=arrList(5,intLoop)%></td>
			<td align="center"><%=arrList(6,intLoop)%></td>
			<td align="center"><%=CategoryName(arrList(7,intLoop))%></td>
		</tr>
		<% next %>
	</table>
<% Else %>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    <td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
    </tr>
	</table>
<% End If %>

<!-- #include virtual="/lib/db/dbclose.asp" -->

