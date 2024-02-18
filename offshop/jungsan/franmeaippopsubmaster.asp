<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 정산리스트
' History : 2009.04.07 서동석 생성
'			2012.09.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<%
dim idx ,i ,totalsellcash, totalbuycash, totalsuplycash, totalorgsellcash
	idx = request("idx")

dim ofranchulgomaster
set ofranchulgomaster = new CFranjungsan
	ofranchulgomaster.FRectidx = idx
	ofranchulgomaster.getOneFranJungsan

dim ofranchulgojungsan
set ofranchulgojungsan = new CFranjungsan
	ofranchulgojungsan.FPageSize=200
	ofranchulgojungsan.FRectIDx = idx
	ofranchulgojungsan.getFranMaeipSubmasterList

%>

<script language='javascript'>

function popSubdetailEdit(iid,itopid){
	var popwin = window.open('franmeaippopsubdetail.asp?idx=' + iid + '&topidx=' + itopid,'franmeaippopsubdetail','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

</script>

<table width="760" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_number %></td>
	<td bgcolor="#FFFFFF" width=280><%= ofranchulgomaster.FOneItem.Fidx %></td>
	<td bgcolor="#DDDDFF" width=120><%= CTX_SHOP %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Fshopid %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_divide %></td>
	<td bgcolor="#FFFFFF" >
		<font color="<%= ofranchulgomaster.FOneItem.GetDivCodeColor %>"><%= ofranchulgomaster.FOneItem.GetDivCodeName %></font>
	</td>
	<td bgcolor="#DDDDFF" width=120><%= CTX_title %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Ftitle %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_total_consumer_price %></td>
	<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgomaster.FOneItem.Ftotalsellcash,0) %></td>
	<td bgcolor="#DDDDFF" width=120> </td>
	<td bgcolor="#FFFFFF" > </td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_total_Supply_price %></td>
	<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgomaster.FOneItem.Ftotalsuplycash,0) %>
	<font color="#AAAAAA">(샾으로 공급한 상품가격)</font></td>
	<td bgcolor="#DDDDFF" width=120><%= CTX_Bill %>&nbsp;<%= CTX_Issue %></td>
	<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgomaster.FOneItem.Ftotalsum,0) %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_tax_Bill_date %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Ftaxdate %></td>
	<td bgcolor="#DDDDFF" width=120><%= CTX_Deposit_Date %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Fipkumdate %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_Status %></td>
	<td bgcolor="#FFFFFF" colspan=3>
		<font color="<%= ofranchulgomaster.FOneItem.GetStateColor %>"><%= ofranchulgomaster.FOneItem.GetStateName %></font>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_etc %></td>
	<td bgcolor="#FFFFFF" colspan=3>
		<%= nl2Br(ofranchulgomaster.FOneItem.Fetcstr) %>
	</td>
</tr>
<!--
<tr>
	<td bgcolor="#DDDDFF" width=100>최초등록자</td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Fregusername %>(<%= ofranchulgomaster.FOneItem.Freguserid %>)</td>
	<td bgcolor="#DDDDFF" width=100>최종처리자</td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Ffinishusername %>(<%= ofranchulgomaster.FOneItem.Ffinishuserid %>)</td>
</tr>
-->
</table>

<br>

<% if ofranchulgomaster.FOneItem.FDivcode="MC" then %>
	<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
	<tr bgcolor="#DDDDFF" align="center">
		<td><%= CTX_SHOP %></td>
		<td><%= CTX_release_code %></td>
		<td><%= CTX_Order_code %></td>
		<td><%= CTX_Shipment_Date %></td>
		<td><%= CTX_total_selling_price %></td>
		<td><%= CTX_total_Supply_price %></td>
		<td><%= CTX_etc %></td>
	</tr>
	<% for i=0 to  ofranchulgojungsan.FResultCount -1 %>
	<%
	totalsellcash	=	totalsellcash + ofranchulgojungsan.FItemList(i).Ftotalsellcash
	totalbuycash	=	totalbuycash + ofranchulgojungsan.FItemList(i).Ftotalbuycash
	totalsuplycash	=	totalsuplycash + ofranchulgojungsan.FItemList(i).Ftotalsuplycash
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td ><a href="javascript:popSubdetailEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>','<%= idx %>');"><%= ofranchulgojungsan.FItemList(i).Fshopid %></a></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fcode01 %></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fcode02 %></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fexecdate %></td>
		<td align="right"><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotalsellcash,0) %></td>
		<td align="right"><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotalsuplycash,0) %></td>
		<td align="center"><a href="javascript:popSubdetailEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>','<%= idx %>');"><%= CTX_DETAILVIEW %></a></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td align="right"><%= FormatNumber(totalsellcash,0) %></td>
		<td align="right"><%= FormatNumber(totalsuplycash,0) %></td>
		<td></td>
	</tr>
	</table>
<% elseif ofranchulgomaster.FOneItem.FDivcode="WS" then %>
	<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
	<tr bgcolor="#DDDDFF" align=center>
		<td><%= CTX_SHOP %></td>
		<td><%= CTX_date %></td>
		<td><%= CTX_Brand %></td>
		<td><%= CTX_total_selling_price %></td>
		<td><%= CTX_total_consumer_price %></td>
		<td><%= CTX_total_Supply_price %></td>
		<td><%= CTX_etc %></td>
	</tr>
	<% for i=0 to  ofranchulgojungsan.FResultCount -1 %>
	<%
	totalsellcash	=	totalsellcash + ofranchulgojungsan.FItemList(i).Ftotalsellcash
	totalsuplycash	=	totalsuplycash + ofranchulgojungsan.FItemList(i).Ftotalsuplycash
	totalorgsellcash =  totalorgsellcash + ofranchulgojungsan.FItemList(i).Ftotalorgsellcash
	%>
	<tr bgcolor="#FFFFFF">
		<td ><%= ofranchulgojungsan.FItemList(i).Fshopid %></td>
		<td ><a href="javascript:popSubdetailEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>','<%= idx %>');"><%= ofranchulgojungsan.FItemList(i).Fcode01 %></a></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fcode02 %></td>
		<td align=right><%= Formatnumber(ofranchulgojungsan.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><%= Formatnumber(ofranchulgojungsan.FItemList(i).Ftotalorgsellcash,0) %></td>
		<td align=right><%= Formatnumber(ofranchulgojungsan.FItemList(i).Ftotalsuplycash,0) %></td>
		<td align=center><a href="javascript:popSubdetailEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>','<%= idx %>');"><%= CTX_DETAILVIEW %></a></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td></td>
		<td></td>
		<td></td>
		<td align="right"><%= FormatNumber(totalsellcash,0) %></td>
		<td align="right"><%= FormatNumber(totalorgsellcash,0) %></td>
		<td align="right"><%= FormatNumber(totalsuplycash,0) %></td>
		<td></td>
	</tr>
	</table>
<% end if %>

<%
set ofranchulgomaster = Nothing
set ofranchulgojungsan = Nothing
%>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->