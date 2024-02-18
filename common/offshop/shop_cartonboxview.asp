<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 샵별패킹리스트(박스별)
' History : 2012.02.02 이상구 생성
'			2012.09.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<%
dim idx ,page, shopid, showmichulgo, workstate
dim currcartoonboxno, isnewcartoonbox ,research, i, j
dim suminnerboxweight, sumcartoonboxNweight, sumcartoonboxweight, sumemsprice
	menupos = requestCheckVar(request("menupos"),10)
	idx = requestCheckVar(request("idx"),10)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	showmichulgo = requestCheckVar(request("showmichulgo"),10)
	research = requestCheckVar(request("research"),2)

if (page = "") then
	page = 1
end if

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
    workstate = "6,7"
end if

dim ocartoonboxmaster
set ocartoonboxmaster = new CCartoonBox
	ocartoonboxmaster.FRectMasterIdx = idx
	ocartoonboxmaster.FRectShopid = shopid
	ocartoonboxmaster.FRectWorkState = workstate
	ocartoonboxmaster.GetMasterOne

if (ocartoonboxmaster.FResultCount < 1) then
	response.write CTX_The_wrong_approach
	response.end
end if

dim ocartoonboxdetail
set ocartoonboxdetail = new CCartoonBox
	ocartoonboxdetail.FRectMasterIdx = idx
	ocartoonboxdetail.FRectShopid = ocartoonboxmaster.FOneItem.Fshopid
	ocartoonboxdetail.GetDetailList

dim oinnerboxlist
set oinnerboxlist = new CCartoonBox
	oinnerboxlist.FRectMasterIdx = -1
	oinnerboxlist.FRectShopid = ocartoonboxmaster.FOneItem.Fshopid
	oinnerboxlist.GetInnerBoxList

%>

<script type='text/javascript'>

function downloadOrder(masteridx, baljucode, shopid, cartoonboxmasteridx) {
	var popwin = window.open("/common/popOrderSheet_foreign_excel.asp?masteridx=" + masteridx + "&baljucode=" +baljucode + "&shopid=" +shopid + "&cartoonboxmasteridx=" +cartoonboxmasteridx,"ExcelOfflineOrderSheet","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PrintDetailItemList(jungsanidx, shopid, shopname) {
	var popwin;
	popwin = window.open('/admin/fran/popcartonboxitemlist_print.asp?jungsanidx=' + jungsanidx + '&shopid=' + shopid + '&shopname=' + shopname + '&xl=Y','PrintDetailItemList','width=850,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopOpenInvoice(invoiceidx) {
	var popwin;
	popwin = window.open('/admin/fran/popoffinvoice_print.asp?idx=' + invoiceidx + '&xl=Y','PopOpenInvoice','width=850,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopOpenPackingList(invoiceidx) {
	var popwin;
	popwin = window.open('/admin/fran/popoffinvoice_print_packinglist.asp?idx=' + invoiceidx + '&xl=Y','PopOpenPackingList','width=850,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopBoxItemList(shopid, yyyy, mm, dd, boxno) {
	var popurl = "/common/offshop/shop_jumunbyboxitemlist.asp?research=on&shopid=" + shopid + "&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=" + dd + "&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + dd + "&boxno=" + boxno;

	var w = window.open(popurl);
	w.focus();
}

function popSubmaster(iid){
	var popwin = window.open('/offshop/jungsan/franmeaippopsubmaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr height="30" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong><%= CTX_Ordering_Information %></strong></font>
			    </td>
				<td align=right>
					<input type="button" class="button" value="BACK" onclick="history.back()">
			    </td>
			</tr>
		</table>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >IDX</td>
	<td>
		<%= ocartoonboxmaster.FOneItem.Fidx %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" ><%= CTX_title %></td>
	<td>
		<%= ocartoonboxmaster.FOneItem.Ftitle %>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" ><%= CTX_SHOP %></td>
	<td><%= ocartoonboxmaster.FOneItem.Fshopid %></td>
	<td bgcolor="<%= adminColor("tabletop") %>" ><%= CTX_shopname %></td>
	<td>
		<%= ocartoonboxmaster.FOneItem.Fshopname %>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_Status %></td>
	<td>
		<%= ocartoonboxmaster.FOneItem.GetStateName %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" ><%= CTX_Shipment_Date %></td>
	<td>
		<input type="text" class="text_ro" name="deliverdt" value="<%= ocartoonboxmaster.FOneItem.Fdeliverdt %>" size=10 readonly >
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_Invoice_Number %></td>
	<td>
		<%= ocartoonboxmaster.FOneItem.GetDeliverMethodName %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >EMS&nbsp;<%= CTX_cost %>(<%= CTX_WON %>)</td>
	<td>
		<input type="text" class="text_ro" name="deliverpay" value="<%= FormatNumber(ocartoonboxmaster.FOneItem.Fdeliverpay,0) %>" size=15 maxlength=100 readonly>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" ><%= CTX_Account %>&nbsp;<%= CTX_code %></td>
	<td>
		<a href="javascript:popSubmaster('<%= ocartoonboxmaster.FOneItem.Fjungsanidx %>')"><%= ocartoonboxmaster.FOneItem.Fjungsanidx %></a>

		<% if idx <> "" and idx <> "0" then %>
			&nbsp;
			<!--<input type="button" class="button" value="<%'= CTX_ITEM %>&nbsp;EXCEL&nbsp;<%'= CTX_Printer %>" onClick="PrintDetailItemList(<%'= ocartoonboxmaster.FOneItem.Fjungsanidx %>, '<%'= ocartoonboxmaster.FOneItem.Fshopid %>', '<%'= ocartoonboxmaster.FOneItem.Fshopname %>')">-->
			<input type="button" onclick="downloadOrder('','','<%= ocartoonboxmaster.FOneItem.Fshopid %>','<%= idx %>');" value="<%= CTX_ITEM %>&nbsp;EXCEL&nbsp;<%= CTX_Printer %>" class="button">
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" ><%= CTX_Invoice %>&nbsp;IDX</td>
	<td>
		<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>
		<% if (ocartoonboxmaster.FOneItem.Finvoiceidx <> "") then %>
			&nbsp;
			<input type="button" class="button" value="<%= CTX_Invoice %>" onClick="PopOpenInvoice(<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>)">
			&nbsp;
			<input type="button" class="button" value="<%= CTX_Packing_List_Carton %>" onClick="PopOpenPackingList(<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>)">
			&nbsp;
			(* 인쇄시 좌/우 여백을 1cm 이하로 조절하세요)
		<% end if %>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" ><%= CTX_writer %></td>
	<td>
		<% if (ocartoonboxmaster.FOneItem.Freguserid = "") then %>
			<%= session("ssBctid") %>
		<% else %>
			<%= ocartoonboxmaster.FOneItem.Freguserid %>
		<% end if %>
		<input type="hidden" name="reguserid" value="<%= session("ssBctid") %>">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" ><%=CTX_registration%><%=ctx_date%></td>
	<td>
		<% if (ocartoonboxmaster.FOneItem.Fregdate = "") then %>
			<%= now %>
		<% else %>
			<%= ocartoonboxmaster.FOneItem.Fregdate %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_etc %></td>
	<td colspan="3"><textarea class="textarea" name="comment" cols="80" rows="6" readonly><%= ocartoonboxmaster.FOneItem.FComment %></textarea>
	</td>
</tr>
</form>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="13" align="right">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td>
			</td>
			<td>
				<%= CTX_search_result %> : <b><%= ocartoonboxdetail.FResultCount %></b>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="110"><%= CTX_SHOP %></td>
    <td width="80"><%= CTX_Real_Order_Date %></td>
    <td width="60">InnerBox<br>No</td>
    <td width="80">InnerBox<br>Weight(KG)</td>
    <td></td>
    <td width="60">CartonBOX<br>No</td>
    <td width="80">CartonBOX<br>N.Weight(KG)</td>
    <td width="80">CartonBOX<br>G.Weight(KG)</td>
	<td width="100">CartonBOX<br>Type</td>
    <td width="90">EMS<br><%= CTX_transport %>&nbsp;<%= ctx_cost %><br>(WON)</td>
    <td><%= CTX_Invoice_Number %></td>
	<td><%= CTX_Note %></td>
</tr>
<%
currcartoonboxno = ""
suminnerboxweight = 0
sumcartoonboxNweight = 0
sumcartoonboxweight = 0
sumemsprice = 0
j = 0
%>
<% for i=0 to ocartoonboxdetail.FResultCount-1 %>
	<%
	if (ocartoonboxdetail.FItemList(i).Fcartoonboxno <> currcartoonboxno) then
		isnewcartoonbox = true
		currcartoonboxno = ocartoonboxdetail.FItemList(i).Fcartoonboxno
	else
		isnewcartoonbox = false
	end if

	if IsNull(ocartoonboxdetail.FItemList(i).FcartoonboxNweight) then
		ocartoonboxdetail.FItemList(i).FcartoonboxNweight = 0
	end if

	if (isnewcartoonbox = true) then
		sumcartoonboxNweight = sumcartoonboxNweight + ocartoonboxdetail.FItemList(i).FcartoonboxNweight
		sumcartoonboxweight = sumcartoonboxweight + ocartoonboxdetail.FItemList(i).Fcartoonboxweight
		sumemsprice = sumemsprice + ocartoonboxdetail.FItemList(i).Femsprice
	end if

	suminnerboxweight = suminnerboxweight + ocartoonboxdetail.FItemList(i).Finnerboxweight
	%>

<tr align="center" bgcolor="#FFFFFF">
	<td><%= ocartoonboxdetail.FItemList(i).Fshopid %></td>
	<td><%= ocartoonboxdetail.FItemList(i).Fbaljudate %></td>
	<td>
		<%= ocartoonboxdetail.FItemList(i).Finnerboxno %>
	</td>
	<td>
		<%= FormatNumber(ocartoonboxdetail.FItemList(i).Finnerboxweight, 2) %>
	</td>
	<td>
		<input type="button" class="button" value="<%= CTX_ITEM %>&nbsp;<%=CTX_DETAILVIEW%>" onClick="PopBoxItemList('<%= ocartoonboxdetail.FItemList(i).Fshopid %>', '<%= Left(ocartoonboxdetail.FItemList(i).Fbaljudate, 4) %>', '<%= Right(Left(ocartoonboxdetail.FItemList(i).Fbaljudate, 7), 2) %>', '<%= Right(ocartoonboxdetail.FItemList(i).Fbaljudate, 2) %>', <%= ocartoonboxdetail.FItemList(i).Finnerboxno %>)">
	</td>
	<td>
		<%= ocartoonboxdetail.FItemList(i).Fcartoonboxno %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<%= FormatNumber(ocartoonboxdetail.FItemList(i).FcartoonboxNweight, 2) %>
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<%= FormatNumber(ocartoonboxdetail.FItemList(i).Fcartoonboxweight, 2) %>
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<%= ocartoonboxdetail.FItemList(i).GetCartoonBoxTypeName %>
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<%= FormatNumber(ocartoonboxdetail.FItemList(i).Femsprice, 0) %>
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<%= ocartoonboxdetail.FItemList(i).Fcartonboxsongjangno %>
		<% end if %>
	</td>
	<td>
		<!--
		<% if (isnewcartoonbox = true) then %>
			&nbsp;
			<input type="button" class="button" value=" 무게자동계산 " onClick="CalcCartoonboxWeight(frmModiPrc_<%= i %>)">
		<% end if %>
		-->
	</td>
</tr>
<% next %>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan=3></td>
    <td><%= FormatNumber(suminnerboxweight, 2) %></td>
    <td colspan=2></td>
    <td><%= FormatNumber(sumcartoonboxNweight, 2) %></td>
    <td><%= FormatNumber(sumcartoonboxweight, 2) %></td>
	<td></td>
    <td><%= FormatNumber(sumemsprice, 0) %></td>
    <td colspan=2></td>
</tr>
</table>

<%
set ocartoonboxdetail = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
