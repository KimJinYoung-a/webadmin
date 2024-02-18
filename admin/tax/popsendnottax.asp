<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이세로 미전송 계산서 엑셀출력
' History : 2012.09.20 한용민 생성
'###########################################################

Response.Expires=-1440
'Response.Buffer=true	
Response.ContentType = "application/vnd.ms-excel" 	
Response.AddHeader "Content-disposition","attachment;filename=TEN" & Left(CStr(now()),10) & "_미전송계산서.xls"
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/tax/EseroTaxCls.asp"-->

<%
dim i

Dim otax
Set otax = new CEsero
	otax.getsendnottax()
%>

<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>	
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="gray">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%=otax.FresultCount%></b>
	</td>
</tr>
<tr align="center" bgcolor="ffffff">
	<td>계산서번호</td>
	<td>발행일</td>
	<td>매출처<Br>사업자번호</td>
	<td>종사업자</td>
	<td>매출처</td>
	<td>매출처<Br>대표자</td>
	<td>매출처<Br>담당EMAIL</td>
	<td>매입처</td>
	<td>매입처<Br>회사명</td>
	<td>매입처<Br>대표자</td>
	<td>매입처<Br>담당자EMAIL</td>
	<td>합계</td>
	<td>공급가</td>
	<td>부가세</td>
	<td>매입구분</td>
	<td>수기여부</td>
	<td>과세구분</td>
	<td>비고</td>
	<td>품목</td>
	<td>사업부문</td>
</tr>

<%
if otax.FResultCount>0 then
	
For i = 0 To otax.FResultCount - 1

%>
<tr align="center" bgcolor="#FFFFFF">
	<td class='txt'><%= otax.FItemList(i).ftaxKey %></td>
	<td><%= otax.FItemList(i).fappDate %></td>
	<td><%= otax.FItemList(i).fsellCorpNo %></td>
	<td><%= otax.FItemList(i).fsellJongNo %></td>
	<td><%= otax.FItemList(i).fsellCorpName %></td>
	<td><%= otax.FItemList(i).fsellCeoName %></td>
	<td><%= otax.FItemList(i).fsellEmail %></td>
	<td><%= otax.FItemList(i).fbuyCorpNo %></td>
	<td><%= otax.FItemList(i).fBuyCorpName %></td>
	<td><%= otax.FItemList(i).fBuyCeoName %></td>
	<td><%= otax.FItemList(i).fbuyEmail %></td>
	<td align="right"><%= FormatNumber(otax.FItemList(i).ftotSum,0) %></td>
	<td align="right"><%= FormatNumber(otax.FItemList(i).fsuplySum,0) %></td>
	<td align="right"><%= FormatNumber(otax.FItemList(i).ftaxSum,0) %></td>
	<td><%= otax.FItemList(i).ftaxSellType %></td>
	<td><%= otax.FItemList(i).ftaxModiType %></td>
	<td><%= otax.FItemList(i).ftaxType %></td>
	<td><%= otax.FItemList(i).fBigo %></td>
	<td><%= otax.FItemList(i).fDtlName %></td>
	<td><%= otax.FItemList(i).fbizseccd %></td>
</tr>

<% Next %>

<% ELSE %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">등록된 내용이 없습니다.</td>
</tr>
<%END IF%>
</table>

</body>
</html>	

<%
Set otax = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" --> 