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
dim idx, topidx ,i ,totalsellcash,totalbuycash,totalsuplycash,totalorgsellcash
	idx = request("idx")
	topidx = request("topidx")

dim ofranchulgomaster
set ofranchulgomaster = new CFranjungsan
	ofranchulgomaster.FRectidx = topidx
	ofranchulgomaster.getOneFranJungsan

dim ofranchulgodetail
set ofranchulgodetail = new CFranjungsan
	ofranchulgodetail.FRectidx = idx
	ofranchulgodetail.getOneFranMaeipSubmaster

dim ofranchulgojungsan
set ofranchulgojungsan = new CFranjungsan
	ofranchulgojungsan.FPageSize=1000
	ofranchulgojungsan.FRectIDx = idx
	ofranchulgojungsan.getFranMaeipSubdetailList

%>

<script language='javascript'>

function SaveArr(frm){
	var ischecked = false;
	frm.suplycasharr.value = "";
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox")) {
			ischecked = (ischecked || e.checked);
			if (e.checked){
				if (frm.elements[i+1].type="text"){
					frm.suplycasharr.value = frm.suplycasharr.value + frm.elements[i+1].value + ",";
				}
			}
		}
	}

	if (!ischecked) {
		alert('<%= CTX_Please_select %> (ITEM)');
		return;
	}

	if (confirm('<%= CTX_Do_you_want_to_save %>?')){
		frm.submit();
	}
}

</script>

<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_number %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgodetail.FOneItem.Fidx %></td>
	<td bgcolor="#DDDDFF" width=120><%= CTX_SHOP %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgodetail.FOneItem.Fshopid %></td>
</tr>
<tr>
<% if ofranchulgomaster.FOneItem.FDivcode="WS" then %>
	<td bgcolor="#DDDDFF" width=120><%= CTX_date %></td>
<% else %>
	<td bgcolor="#DDDDFF" width=120><%= CTX_release_code %></td>
<% end if %>
	<td bgcolor="#FFFFFF" ><%= ofranchulgodetail.FOneItem.Fcode01 %></td>
<% if ofranchulgomaster.FOneItem.FDivcode="WS" then %>
	<td bgcolor="#DDDDFF" width=120><%= CTX_Brand %>&nbsp;ID</td>
<% else %>
	<td bgcolor="#DDDDFF" width=120><%= CTX_Order_code %></td>
<% end if %>
	<td bgcolor="#FFFFFF" ><%= ofranchulgodetail.FOneItem.Fcode02 %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_total_consumer_price %></td>
	<td bgcolor="#FFFFFF" >
		<%= FormatNumber(ofranchulgodetail.FOneItem.Ftotalorgsellcash,0) %>
	</td>
	<td bgcolor="#DDDDFF" width=120><%= CTX_total_selling_price %></td>
	<td bgcolor="#FFFFFF" >
		<%= FormatNumber(ofranchulgodetail.FOneItem.Ftotalsellcash,0) %>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_total_Supply_price %></td>
	<td bgcolor="#FFFFFF" colspan=3>
		<%= FormatNumber(ofranchulgodetail.FOneItem.Ftotalsuplycash,0) %>
	</td>
</tr>
</table>
<br>
<table width="760" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
<tr bgcolor="#DDDDFF" align=center>
	<td><%= CTX_divide %></td>
	<td><%= CTX_Order_code %><Br><%= CTX_input_output %>&nbsp;<%= CTX_code %></td>
	<td><%= CTX_Barcode %></td>
	<td><%= CTX_Description %></td>
	<td><%= CTX_Description_Option %></td>
	<td><%= CTX_quantity %></td>
	<td><%= CTX_consumer_price %></td>
	<td><%= CTX_selling_price %></td>
	<td><%= CTX_Supply_price %></td>
	<td><%= CTX_total_Supply_price %></td>
</tr>
<%
for i=0 to ofranchulgojungsan.FResultCount - 1

totalsuplycash = totalsuplycash + ofranchulgojungsan.FItemList(i).Fsuplycash * ofranchulgojungsan.FItemList(i).Fitemno
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><%= ofranchulgojungsan.FItemList(i).Flinkbaljucode %></td>
	<td><%= ofranchulgojungsan.FItemList(i).Flinkmastercode %></td>
	<td><%= ofranchulgojungsan.FItemList(i).GetBarCode %></td>
	<td><%= ofranchulgojungsan.FItemList(i).Fitemname %></td>
	<td><%= ofranchulgojungsan.FItemList(i).Fitemoptionname %></td>
	<td align="center">
	
	<% if ofranchulgojungsan.FItemList(i).Fitemno<0 then %>
		<font color="red"><%= ofranchulgojungsan.FItemList(i).Fitemno %></font>
	<% else %>
		<%= ofranchulgojungsan.FItemList(i).Fitemno %>
	<% end if %>
	</td>
	<td align="right"><%= FormatNumber(ofranchulgojungsan.FItemList(i).Forgsellcash,0) %></td>
	<td align="right"><%= FormatNumber(ofranchulgojungsan.FItemList(i).Fsellcash,0)  %></td>
	<td align="right"><%= FormatNumber(ofranchulgojungsan.FItemList(i).Fsuplycash,0)  %></td>
	
	<% if ofranchulgojungsan.FItemList(i).Fitemno<0 then %>
		<td align=right>
			<font color="red"><%= FormatNumber(ofranchulgojungsan.FItemList(i).Fsuplycash*ofranchulgojungsan.FItemList(i).Fitemno,0)  %></font>
		</td>
	<% else %>
		<td align="right"><%= FormatNumber(ofranchulgojungsan.FItemList(i).Fsuplycash*ofranchulgojungsan.FItemList(i).Fitemno,0)  %></td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="right">
	<td colspan=9></td>
	<td align="right"><%= FormatNumber(totalsuplycash,0) %></td>
</tr>
</table>

<%
set ofranchulgomaster = nothing
set ofranchulgodetail = nothing
set ofranchulgojungsan = nothing
%>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->