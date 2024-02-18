<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/analysiscls.asp"-->

<%
'response.write "관리자 문의 요망"
'dbget.close()	:	response.End

dim yyyy1,mm1
yyyy1 = request("yyyy1")
mm1 = request("mm1")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim oanal
set oanal = new CAnalysis
oanal.FRectYYYYMM = yyyy1 + "-" + mm1
oanal.getOnLineMonthGainSum

dim i, shopmmttl, shopsuppttl
shopmmttl = 0
shopsuppttl = 0

for i=0 to oanal.FResultCount-1
    shopmmttl = shopmmttl + oanal.FItemList(i).FTotsum
    shopsuppttl = shopsuppttl + oanal.FItemList(i).FSuplysum
next
%>

<table width="900" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="">
	<tr>
		<td class="a" >
		검색대상년월:<% DrawYMBox yyyy1,mm1 %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<h3>수정중</h3>

<span class=a>** Admin 매출내역</span>
<table width="900" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td>웹매출액</td>
	<td>기타출고매출액</td>
	<td>배송비</td>
	<td>배송건수</td>
	<td>총매출액</td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td><%= FormatNumber(oanal.FOneItem.FWebTotalSel,0) %></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
</table>
<br>
<span class=a>** onLine 정산내역</span>
<table width="900" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td>업체배송</td>
	<td>매입</td>
	<td>위탁</td>
	<td>기타출고</td>
	<td>총매입액</td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td><%= FormatNumber(oanal.FOneItem.FUbTotal,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FMeTotal,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FWiTotal,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FEtTotal,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.getTotalMeaip,0) %></td>
</tr>
</table>
<br><br>
<table width="900" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td width=180>구분</td>
	<td width=180>내역</td>
	<td>매출</td>
	<td>매입(결제액)</td>
	<td>비고<br>(오프출고반영 매입액)</td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td rowspan=5 align=center>온라인</td>
	<td align=left>업체배송</td>
	<td rowspan=3 bgcolor="#337799"><%= FormatNumber(oanal.FOneItem.FWebTotalSel,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FUbTotal,0) %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td align=left>위탁</td>
	<td><%= FormatNumber(oanal.FOneItem.FwiTotal,0) %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td align=left>매입</td>
	<td><%= FormatNumber(oanal.FOneItem.FMeTotal,0) %></td>
	<td><%= FormatNumber(shopmmttl,0) %></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td align=left>기타</td>
	<td></td>
	<td><%= FormatNumber(oanal.FOneItem.FEtTotal,0) %></td>
	<td></td>
</tr>
<tr bgcolor="#DDDDDD" align=right>
	<td align=left>소계</td>
	<td><%= FormatNumber(oanal.FOneItem.FWebTotalSel+shopsuppttl,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FMeTotal + oanal.FOneItem.FwiTotal + oanal.FOneItem.FUbTotal + oanal.FOneItem.FEtTotal,0) %></td>
	<td></td>
</tr>

<% for i=0 to oanal.FResultCount-1 %>
<tr bgcolor="#FFFFFF" align=right>
    <% if i=0 then %>
    <td rowspan=<%= oanal.FResultCount + 1 %> align=center>오프라인<br>출고<br>(매입상품)</td>
    <% end if %>
	<td align=left>(<%= oanal.FItemList(i).FShopid %>)</td>
	<td><%= FormatNumber(oanal.FItemList(i).FSuplysum,0) %></td>
	<td></td>
	<td><%= FormatNumber(oanal.FItemList(i).FTotsum,0) %></td>
</tr>
<% next %>

<tr bgcolor="#DDDDDD" align=right>
	<td align=left>소계</td>
	<td></td>
	<td></td>
	<td><%= FormatNumber(shopmmttl,0) %></td>
</tr>

<tr bgcolor="#FFFFFF" align=right>
	<td rowspan=6 align=center>기타출고</td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>

<tr bgcolor="#FFFFFF" align=right>
	<td align=center>총계</td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
</table>
<br>
<table width="800" border="0" cellpadding="2" cellspacing="1" class="a" >
<tr>
	<td width=100 bgcolor="#337799"></td>
	<td>표시 되어 있는 매출이 어드민상에 잡혀있는 매출입니다.</td>
</tr>
<tr>
	<td colspan="2">매출은 결제완료 기준 / 매입은 배송완료 기준입니다.</td>
</tr>
</table>
<%
set oanal = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
