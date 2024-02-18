<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : 재고자산(월별) FIX 엑셀다운로드
' History : 이상구 생성
'			2023.10.11 한용민 수정(csv파일 -> 엑셀파일 생성으로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls_2.asp"-->
<%
dim yyyymm, placeGubun, PriceGbn, ver, oCMonthlyMaeipLedge, arrList, i
    yyyymm = RequestCheckVar(request("yyyymm"),7)
    placeGubun = RequestCheckVar(request("placeGubun"),1)
    PriceGbn = RequestCheckVar(request("PriceGbn"),1)
    ver = RequestCheckVar(request("ver"),10)

if (ver = "") then
	ver = "V2"
end if

set oCMonthlyMaeipLedge = new CMonthlyMaeipLedge
oCMonthlyMaeipLedge.FCurrPage = 1
oCMonthlyMaeipLedge.FPageSize = 1000000
oCMonthlyMaeipLedge.frectver = ver
oCMonthlyMaeipLedge.frectyyyymm = yyyymm
oCMonthlyMaeipLedge.frectplaceGubun = placeGubun
oCMonthlyMaeipLedge.frectPriceGbn = PriceGbn
oCMonthlyMaeipLedge.GetMaeipLedgeListNotPaging

if oCMonthlyMaeipLedge.FTotalCount>0 then
    arrLIst=oCMonthlyMaeipLedge.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENMonthlyMaeipLedgeList" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="50">
		검색결과 : <b><%= oCMonthlyMaeipLedge.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>YYYY-MM</td>
    <td>재고위치</td>
    <td>부서</td>
    <td>과세구분</td>
    <td>브랜드</td>
    <td>상품구분</td>
    <td>상품코드</td>
    <td>옵션코드</td>
    <td>바코드</td>
    <td>단가(평균)</td>
    <td>기초수량</td>
    <td>기초금액</td>
    <td>입고수량</td>
    <td>입고금액</td>
    <td>이동수량</td>
    <td>이동금액</td>
    <td>판매수량</td>
    <td>판매금액</td>
    <td>오프출고수량</td>
    <td>오프출고금액</td>
    <td>기타출고수량</td>
    <td>기타출고금액</td>
    <td>CS출고수량</td>
    <td>CS출고금액</td>
    <td>오차수량</td>
    <td>오차금액</td>
    <td>기말수량</td>
    <td>기말금액</td>

	<% if placeGubun <> "S" then %>
		<td>최종입고월</td>
	<% end if %>
	<% if placeGubun <> "L" and placeGubun <> "T" and placeGubun <> "O" and placeGubun <> "N" and placeGubun <> "F" and placeGubun <> "A" and placeGubun <> "R" then %>
		<td>최종입고월(매입구분별)</td>
	<% end if %>

    <td>사업자번호</td>
    <td>재고구분</td>
    <td>전시카테고리</td>
    <td>관리카테1</td>
    <td>관리카테2</td>
    <td>구매유형</td>
    <td>센터매입구분</td>
    <td>상품매입구분</td>
    <td>소비자가</td>
    <td>현재판매가</td>
    <td>현재판매여부</td>

	<% if (ver = "DW") then %>
        <td>보너스쿠폰적용가</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td>상품단종여부</td>
		<td>옵션단종여부</td>
	<% end if %>
</tr>
<% if isarray(arrLIst) then %>
<%
for i=0 to ubound(arrLIst,2)
%>
<tr bgcolor="#FFFFFF" align="center">
    <td class="txt"><%= arrList(1,i) %></td><% 'YYYY-MM %>
    <td class="txt"><%= trim(arrList(2,i)) %></td><% '재고위치 %>
    <td><%= arrList(3,i) %></td><% '부서 %>
    <td><%= arrList(4,i) %></td><% '과세구분 %>
    <td class="txt"><%= arrList(26,i) %></td><% '브랜드 %>
    <td><%= arrList(5,i) %></td><% '상품구분 %>
    <td><%= arrList(6,i) %></td><% '상품코드 %>
    <td class="txt"><%= arrList(7,i) %></td><% '옵션코드 %>

    <% if (ver = "DW") then %>
        <td class="txt"><%= arrList(44,i) %></td><% '바코드 %>
    <% else %>
        <td class="txt"><%= arrList(42,i) %></td><% '바코드 %>
    <% end if %>

    <td><%= arrList(28,i) %></td><% '단가(평균) %>
    <td><%= arrList(8,i) %></td><% '기초수량(SYS재고) %>
    <td><%= arrList(9,i) %></td><% '기초금액(SYS재고) %>
    <td><%= arrList(10,i) %></td><% '입고수량 %>
    <td><%= arrList(11,i) %></td><% '입고금액 %>
    <td><%= arrList(12,i) %></td><% '이동수량 %>
    <td><%= arrList(13,i) %></td><% '이동금액 %>
    <td><%= arrList(14,i) %></td><% '판매수량 %>
    <td><%= arrList(15,i) %></td><% '판매금액 %>
    <td><%= arrList(16,i) %></td><% '오프출고수량 %>
    <td><%= arrList(17,i) %></td><% '오프출고금액 %>
    <td><%= arrList(20,i) %></td><% '기타출고수량(구:로스출고) %>
    <td><%= arrList(21,i) %></td><% '기타출고금액(구:로스출고) %>
    <td><%= arrList(22,i) %></td><% 'CS출고수량 %>
    <td><%= arrList(23,i) %></td><% 'CS출고금액 %>
    <td><%= (arrList(8,i) + arrList(10,i)+ arrList(12,i)+ arrList(14,i)+arrList(16,i)+ arrList(18,i)+arrList(20,i) +arrList(22,i)- arrList(24,i))*-1 %></td><% '오차수량 %>
    <td><%= (arrList(9,i) + arrList(11,i)+ arrList(13,i)+ arrList(15,i)+arrList(17,i)+ arrList(19,i)+arrList(21,i) +arrList(23,i)- arrList(25,i))*-1 %></td><% '오차금액 %>
    <td><%= arrList(24,i) %></td><% '기말수량(시스템재고) %>
    <td><%= arrList(25,i) %></td><% '기말금액(시스템재고) %>

    <% if placeGubun <> "S" then %>
        <td class="txt"><%= arrList(29,i) %></td><% '최종입고월 %>
    <% end if %>

    <% if placeGubun <> "L" and placeGubun <> "T" and placeGubun <> "O" and placeGubun <> "N" and placeGubun <> "F" and placeGubun <> "A" and placeGubun <> "R" then %>
        <td class="txt"><%= arrList(30,i) %></td><% '최종입고월(매입구분별) %>
    <% end if %>

    <td class="txt"><%= arrList(32,i) %></td><% '사업자번호 %>
    <td><%= arrList(27,i) %></td><% '재고구분 %>
    <td><%= arrList(31,i) %></td><% '전시카테고리 %>
    <td><%= arrList(34,i) %></td><% '관리카테1 %>
    <td><%= arrList(35,i) %></td><% '관리카테2 %>
    <td><%= arrList(36,i) %></td><% '구매유형 %>
    <td><%= arrList(37,i) %></td><% '센터매입구분 %>
    <td><%= arrList(38,i) %></td><% '상품매입구분 %>
    <td><%= arrList(39,i) %></td><% '소비자가 %>
    <td><%= arrList(40,i) %></td><% '현재판매가 %>
    <td><%= arrList(41,i) %></td><% '현재판매여부 %>

    <% if (ver = "DW") then %>
        <td><%= arrList(42,i) %></td><% '취급액(보너스쿠폰적용가) %>
        <td><%= arrList(43,i) %></td><% '상품명 %>
        <td><%= arrList(47,i) %></td><% '옵션명 %>
        <td><%= arrList(45,i) %></td><% '상품단종여부 %>
        <td><%= arrList(46,i) %></td><% '옵션단종여부 %>
    <% end if %>
</tr>
<%
if i mod 300 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="50" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>

<%
set oCMonthlyMaeipLedge = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->