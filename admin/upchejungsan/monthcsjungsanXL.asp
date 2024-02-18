<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
Dim yyyy1, mm1, jgubun
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)
jgubun = "CC"

dim oCSetcjungsan
set oCSetcjungsan = new CUpcheJungsanTax
	oCSetcjungsan.FPageSize = 5000
	oCSetcjungsan.FCurrPage = 1
	oCSetcjungsan.FRectYYYYMM = yyyy1 & "-" & mm1
	oCSetcjungsan.FRectJGubun = jgubun
	oCSetcjungsan.getMonthCsjungsanList
dim i
%>

<!-- 엑셀파일로 저장 헤더 부분 -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
Response.ContentType = "application/unknown"
Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=CS기타정산_"&yyyy1&"-"&mm1&".xls"
%>
<style type="text/css">
/* 엑셀 다운로드로 저장시 숫자로 표시될 경우 방지 */
.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="100">브랜드ID</td>
    <td width="100">입출코드</td>
    <td width="100">제휴몰주문번호</td>
    <td width="100">판매채널</td>
    <td width="100">구매자</td>
    <td width="100">수령인</td>
    <td width="100">상품코드</td>
    <td width="100">옵션코드</td>
    <td width="100">상품명</td>
    <td width="180">옵션명</td>
    <td width="100">수량</td>
    <td width="100">구매총액</td>
    <td width="100">기본판매수수료</td>
    <td width="100">쿠폰할인액(텐바이텐부담)</td>
    <td width="200">고객실주문액(협력사매출액)</td>
    <td width="100">수수료</td>
  	<td width="200">결제대행수수료</td>
  	<td width="100">정산액</td>
  	<td width="200">정산합계(수량*정산액)</td>
  	<td width="100">주결제수단</td>
</tr>
<% For i=0 to oCSetcjungsan.FresultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FDesignerid%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FMastercode%></td>
    <td class="txt"><% if oCSetcjungsan.FItemList(i).FSitename<>"10x10" then %><%=oCSetcjungsan.FItemList(i).Fauthcode%><% end if %></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FSitename%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FBuyname%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FReqname%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemid%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemoption%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemname%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemoptionname%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemno%></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FSellcash, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FCouponPlusCommi, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FCoupoonDiscount, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FReducedprice, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FCommission, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FPgcommission, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FSuplycash, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FSumsuplycash,0) %></td>
    <td class="txt"><%= oCSetcjungsan.FItemList(i).FPaymethod %></td>
</tr>
<% Next %>
</table>
</body>
</html>
<% Set oCSetcjungsan = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
