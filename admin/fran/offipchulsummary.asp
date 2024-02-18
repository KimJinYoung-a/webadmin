<%@ language=vbscript %>
<% option explicit %>
<%
response.write " 수정중 - 사용안함 관리자 문의요망 "
response.End

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshopstoragecls.asp"-->

<%
dim i, shopid, startdt, enddt, makerid
dim yyyy1,mm1

shopid = request("shopid")
makerid = request("makerid")
yyyy1 = request("yyyy1")
mm1 = request("mm1")

if yyyy1="" then
	yyyy1 = Cstr(now())
	mm1 = Mid(yyyy1,6,2)
	yyyy1 = left(yyyy1,4)
	startdt = yyyy1 + "-" + mm1 + "-01"
	enddt = CStr(DateSerial(yyyy1,mm1+1,1))
else
	startdt = yyyy1 + "-" + mm1 + "-01"
	enddt = CStr(DateSerial(yyyy1,mm1+1,1))
end if

dim ooffipchul
set ooffipchul = new COffShopStorage
ooffipchul.FRectShopid = shopid
ooffipchul.FRectStartDate = startdt
ooffipchul.FRectEndDate = enddt
ooffipchul.FRectMakerid = makerid
ooffipchul.getStorageNSellList
%>
<table width="800" border="0" cellpadding="5" cellspacing="0" class=a>
<tr>
	<td>
		* 현재 익월재고 내역이 없습니다.- 조만간 적용하겠습니다.<br>
		* 현재 입고내역 기준으로 작성되었습니다. - 조만간 재고기준으로 적용하겠습니다.(입고가 없으면 내역이 안나옵니다.)<br>
	</td>
</tr>
</table>
<br>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			년월 : <% DrawYMBox yyyy1,mm1 %>
			오프샾: <% drawSelectBoxOffShop "shopid",shopid %>
			<br>
			브랜드: <% drawSelectBoxPartnerDesigner "makerid", makerid %>

		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="800" border="0" cellpadding="5" cellspacing="0" class=a>
<tr>
	<td align=right>
		총건수 : <%=  ooffipchul.FResultCount %>
	</td>
</tr>
</table>
<br>
<table width=800 cellpadding=0 class=a cellspacing="1" bgcolor="#3d3d3d">
<tr bgcolor="#DDDDFF" align=center>
	<td width=100>브랜드</td>
	<td width=100>상품번호</td>
	<td width=160>상품명</td>
	<td width=80>옵션명</td>
	<td width=50>익월재고</td>
	<td width=50>입고</td>
	<td width=50>반품</td>
	<td width=50>판매</td>
	<td width=50>R</td>
</tr>
<% for i=0 to ooffipchul.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= ooffipchul.FItemList(i).FMakerid %></td>
	<td><%= ooffipchul.FItemList(i).GetBargode %></td>
	<td><%= ooffipchul.FItemList(i).Fitemname %></td>
	<td><%= ooffipchul.FItemList(i).Fitemoptionname %></td>
	<td><%= ooffipchul.FItemList(i).FLastrealno %></td>
	<td><%= ooffipchul.FItemList(i).Fipno %></td>
	<td><%= ooffipchul.FItemList(i).Freno %></td>
	<td><%= ooffipchul.FItemList(i).Fsellno %></td>
	<td><%= ooffipchul.FItemList(i).GetMayno %></td>
</tr>
<% next %>
</table>
<%
set ooffipchul = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->