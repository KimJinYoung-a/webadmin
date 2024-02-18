<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ������ŷ����Ʈ(�ڽ���)
' History : �̻� ����
'			2017.04.11 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<%

menupos = requestCheckVar(request("menupos"),10)

dim shopid, statecd, idx
dim research, i

shopid = requestCheckVar(request("shopid"),32)
idx = requestCheckVar(request("idx"),10)
research = requestCheckVar(request("research"),2)

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
    statecd = "7"
end if

dim ocoffinvoice
set ocoffinvoice = new COffInvoice
	ocoffinvoice.FRectShopid = shopid
	'//ocoffinvoice.FRectStateCD = statecd
	ocoffinvoice.FRectMasterIdx = idx
	ocoffinvoice.GetMasterOne

if (ocoffinvoice.FResultCount < 1) then
	response.write "<script type='text/javascript'>alert('�߸��� �����Դϴ�.');history.back();</script>"
	response.end
end if

'================================================================================
dim ocoffinvoicedetail
set ocoffinvoicedetail = new COffInvoice
	ocoffinvoicedetail.FRectMasterIdx = idx
	ocoffinvoicedetail.FRectShopid = ocoffinvoice.FOneItem.Fshopid
	ocoffinvoicedetail.GetDetailList

'================================================================================
dim ocoffinvoiceproductdetail
set ocoffinvoiceproductdetail = new COffInvoice
	ocoffinvoiceproductdetail.FRectMasterIdx = idx
	ocoffinvoiceproductdetail.FRectShopid = ocoffinvoice.FOneItem.Fshopid
	ocoffinvoiceproductdetail.GetProductDetailList

%>

<script type='text/javascript'>

function FormatNumber(nStr) {
	nStr += '';

	x = nStr.split('.');
	x1 = x[0];
	x2 = x.length > 1 ? '.' + x[1] : '';

	var rgx = /(\d+)(\d{3})/;
	while (rgx.test(x1)) {
		x1 = x1.replace(rgx, '$1' + ',' + '$2');
	}

	return x1 + x2;
}

function RecalcPrice(frm) {
	var exchangerate 		= frm.exchangerate.value.replace(/,/g, "");
	var totalboxprice 		= frm.totalboxprice.value.replace(/,/g, "");
	var totalgoodsprice 	= frm.totalgoodsprice.value.replace(/,/g, "");
	var totalprice			= 0;

	var priceunit 			= frm.priceunit.value;

	var totalgoodspricecalc = 0;
	var totalboxpricecalc 	= 0;
	var totalpricecalc 		= 0;

	var pointprice = 2;		// �Ҽ������� ���ڸ����� ���

	frm.totalprice.value = totalprice.toFixed(pointprice);
	frm.totalgoodspricecalc.value = totalgoodspricecalc.toFixed(pointprice);
	frm.totalboxpricecalc.value = totalboxpricecalc.toFixed(pointprice);
	frm.totalpricecalc.value = totalpricecalc.toFixed(pointprice);

	// ========================================================================
	if ((exchangerate == "") || (exchangerate*0 != 0) || (exchangerate*1 == 0)) {
		return "ȯ���� �Է��ϼ���.";
	}

	if ((totalgoodsprice == "") || (totalgoodsprice*0 != 0) || (totalgoodsprice*1 == 0)) {
		return "��ǰ�ݾ��� �Է��ϼ���";
	}

	if ((totalboxprice == "") || (totalboxprice*0 != 0)) {
		// ������ ���� �� �ִ�.;
		return "������ �Է��ϼ���";
	}

	if (priceunit == "") {
		return "�ۼ�ȭ�� �����ϼ���.";
	}

	// ========================================================================
	exchangerate = exchangerate*1;
	totalboxprice = totalboxprice*1;
	totalgoodsprice = totalgoodsprice*1;

	if (priceunit == "JPY") {
		// ��ȭ�� 100�� �����ش�.
		exchangerate = exchangerate*1 / 100;
		pointprice = 0;
	}

	totalgoodspricecalc = (totalgoodsprice / exchangerate).toFixed(pointprice);
	totalboxpricecalc = (totalboxprice / exchangerate).toFixed(pointprice);
	totalpricecalc = ((totalgoodsprice + totalboxprice) / exchangerate).toFixed(pointprice);

	// ========================================================================
	frm.exchangerate.value = FormatNumber(exchangerate);
	frm.totalboxprice.value = FormatNumber(totalboxprice);
	frm.totalgoodsprice.value = FormatNumber(totalgoodsprice);

	frm.totalboxpricecalc.value = FormatNumber(totalboxpricecalc);
	frm.totalgoodspricecalc.value = FormatNumber(totalgoodspricecalc);
	frm.totalpricecalc.value = FormatNumber(totalpricecalc);

	return "";
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopPrevInvoiceList(frm, mode) {
	var popwin;
	popwin = window.open('/admin/fran/popoffinvoice_list.asp?shopid=' + frm.shopid.value + '&frm=' + frm.name + '&mode=' + mode ,'PopPrevInvoiceList','width=1000,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PrintExportSheet(idx, mode){
	var popwin;
	if (mode == "INVOICE") {
		popwin = window.open('/admin/fran/popoffinvoice_print.asp?idx=' + idx + '&xl=Y','PrintExportSheet','width=850,height=600,scrollbars=yes,resizable=yes');
	} else {
		popwin = window.open('/admin/fran/popoffinvoice_print_packinglist.asp?idx=' + idx + '&xl=Y','PrintExportSheet','width=850,height=600,scrollbars=yes,resizable=yes');
	}

	popwin.focus();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="statecd" value="<%= ocoffinvoice.FOneItem.Fstatecd %>">
<input type="hidden" name="productdetailmode" value="">
<input type="hidden" name="productdetailidx" value="">
<input type="hidden" name="masteridx" value="<%= idx %>">

<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�κ��̽� �⺻����</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">IDX</td>
	<td><%= ocoffinvoice.FOneItem.Fidx %></td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�����</td>
	<td>
		<%= ocoffinvoice.FOneItem.Freguserid %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="200">�����̵�</td>
	<% if ocoffinvoice.FOneItem.Fshopid<>"" then %>
	<input type=hidden name="shopid" value="<%= ocoffinvoice.FOneItem.Fshopid %>">
	<td><%= ocoffinvoice.FOneItem.Fshopid %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td><%= ocoffinvoice.FOneItem.Fshopname %></td>
	<% else %>
	<td colspan=3></td>
	<% end if %>
</tr>

<!-- ��ܹ� ���� -->
<tr height="30" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�κ��̽� ��������</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >��۹��</td>
	<td>
		<%= ocoffinvoice.FOneItem.GetDeliverMethodName %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >���Ӻδ�</td>
	<td>
		<%= ocoffinvoice.FOneItem.GetExportMethodName %>
		&nbsp;
		* CFR : ���������ε�(��������) / FOB : �����ε�
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >������</td>
	<td>
		<%= ocoffinvoice.FOneItem.GetJungsanTypeName %>
		&nbsp;
		* TT : ����ȯ�۱�(������) / LC : �ſ���
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�ۼ�ȭ��</td>
	<td>
		<% drawSelectBoxPriceUnit "priceunit", ocoffinvoice.FOneItem.Fpriceunit %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >����ȯ��</td>
	<td>
		<input type="text" class="text_ro" name="exchangerate" value="<%= FormatNumber(ocoffinvoice.FOneItem.Fexchangerate, 2) %>" size=10>��
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" ></td>
	<td>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >��ǰ�ݾ�(��)</td>
	<td>
		<input type="text" class="text_ro" name="totalgoodsprice" value="<%= FormatNumber(ocoffinvoice.FOneItem.Ftotalgoodsprice, 2) %>" size=20 readonly>��
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >����(��)</td>
	<td>
		<input type="text" class="text_ro" name="totalboxprice" value="<%= FormatNumber(ocoffinvoice.FOneItem.Ftotalboxprice, 2) %>" size=20 readonly>��
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >��ǰ�ݾ�(��ȯ)</td>
	<td>
		<input type="text" class="text_ro" name="totalgoodspricecalc" value="" size=10 readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >����(��ȯ)</td>
	<td>
		<input type="text" class="text_ro" name="totalboxpricecalc" value="" size=10 readonly>
	</td>
</tr>

<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�κ��̽� ������</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">1. Shipper/Exporter(�������)</td>
	<td colspan="3"><textarea class="textarea" name="exporteraddr" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Fexporteraddr %></textarea>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">2. For account & Risk of Messers.<br>(���Ծ���)</td>
	<td colspan="3"><textarea class="textarea" name="riskmesseraddr" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Friskmesseraddr %></textarea>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">3. Notify party(��������ó)</td>
	<td colspan="3"><textarea class="textarea" name="notifyaddr" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Fnotifyaddr %></textarea>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >4. Port of loading(������)</td>
	<td>
		<input type="text" class="text_ro" name="portname" value="<%= ocoffinvoice.FOneItem.Fportname %>" size=20 readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >5. Final destination(��������)</td>
	<td>
		<input type="text" class="text_ro" name="destinationname" value="<%= ocoffinvoice.FOneItem.Fdestinationname %>" size=20 readonly>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >6. Carrier(�����̸�)</td>
	<td>
		<input type="text" class="text_ro" name="carriername" value="<%= ocoffinvoice.FOneItem.Fcarriername %>" size=20 readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >7. Sailing on or about(������)</td>
	<td>
		<input type="text" class="text_ro" name="carrierdate" value="<%= ocoffinvoice.FOneItem.Fcarrierdate %>" size=10 readonly >
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >8.No.& date of invoice<br>(�κ��̽�DATE)</td>
	<td>
		<input type="text" class="text_ro" name="invoicedate" value="<%= ocoffinvoice.FOneItem.Finvoicedate %>" size=10 readonly >
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�κ��̽�NO</td>
	<td>
		<input type="text" class="text_ro" name="invoiceno" value="<%= ocoffinvoice.FOneItem.Finvoiceno %>" size=30 readonly>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >9. No.& date of L/C(�ſ���)</td>
	<td>
		<input type="text" class="text_ro" name="lccomment" value="<%= ocoffinvoice.FOneItem.Flccomment %>" size=20 readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >10. L/C issuing bank(�߱�����)</td>
	<td>
		<input type="text" class="text_ro" name="lcbank" value="<%= ocoffinvoice.FOneItem.Flcbank %>" size=10 readonly>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">11. Remarks(���)</td>
	<td colspan="3"><textarea class="textarea" name="comment" cols="80" rows="6" readonly><%= ocoffinvoice.FOneItem.Fcomment %></textarea>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" width="40" align="center">ǥ��<br>����</td>
			<td bgcolor="<%= adminColor("tabletop") %>">
				12. Description of Goods<br>(��ǰ����)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">
				13. Q'ty / BOX<br>(ī��ڽ�����)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">
				14. Price / BOX<br>(��ջ�ǰ�ݾ�)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">
				15. Amount<br>(�ѻ�ǰ�ݾ�)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100" align="center">���</td>
		</tr>
		<input type="hidden" name="productdetailcount" value="<%= ocoffinvoiceproductdetail.FResultCount %>">
		<% for i=0 to ocoffinvoiceproductdetail.FResultCount-1 %>
			<%
			if (ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox <> "") then
				ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox = FormatNumber(ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox, 2)
			end if
			if (ocoffinvoiceproductdetail.FItemList(i).Ftotalprice <> "") then
				ocoffinvoiceproductdetail.FItemList(i).Ftotalprice = FormatNumber(ocoffinvoiceproductdetail.FItemList(i).Ftotalprice, 2)
			end if
			%>
		<input type="hidden" name="productdetailidx_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Fidx %>">
		<tr bgcolor="#FFFFFF">
			<td align="center"><input type="text" class="text_ro" name="orderno_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Forderno %>" size=2 readonly></td>
			<td><input type="text" class="text_ro" name="goodscomment_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Fgoodscomment %>" size=60 readonly></td>
			<td><input type="text" class="text_ro" name="totalboxno_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Ftotalboxno %>" size=10 readonly></td>
			<td>
				<input type="text" class="text_ro" name="priceperbox_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox %>" size=10 readonly>
			</td>
			<td><input type="text" class="text_ro" name="totalprice_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Ftotalprice %>" size=10 readonly></td>
			<td align="center">
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF">
			<td colspan=6 height=30></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="center"></td>
			<td>Total</td>
			<td>
				<input type="text" class="text_ro" name="totalboxno" value="<%= ocoffinvoice.FOneItem.Ftotalboxno %>" size=10 readonly>
			</td>
			<td></td>
			<td>
				<input type="hidden" name="totalprice" value="0">
				<input type="text" class="text_ro" name="totalpricecalc" value="<%= FormatNumber(ocoffinvoice.FOneItem.Ftotalprice, 2) %>" size=10 readonly>
			</td>
			<td align="center"></td>
		</tr>
		</table>
	</td>
</tr>

<tr height=40 bgcolor="#FFFFFF">
	<td colspan="4" align="center">
        <input type="button" class="button" value=" �κ��̽� �����ޱ� " onClick="PrintExportSheet(<%= idx %>, 'INVOICE')">
        &nbsp;
        <input type="button" class="button" value=" ��ŷ����Ʈ �����ޱ� " onClick="PrintExportSheet(<%= idx %>, 'PACKINGLIST')">
	</td>
</tr>

</form>
</table>

<br>

<!-- ��ܹ� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>��ŷ����Ʈ ����</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>10. Marks & number of pkgs<br>(�ڽ���ȣ)</td>
    <td>11. Description of Goods<br>(�ڽ�����)</td>
    <td>12. Quantity/unit</td>
	<td>13. N weight<br>(Inner�ڽ�����)</td>
	<td>14. G weight<br>(Carton�ڽ�����)</td>
	<td>���</td>
</tr>
<% for i=0 to ocoffinvoicedetail.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= ocoffinvoicedetail.FItemList(i).Fcartonboxno %> BOX</td>
	<td><%= ocoffinvoicedetail.FItemList(i).Fgoodscomment %></td>
	<td></td>
	<td><%= FormatNumber(ocoffinvoicedetail.FItemList(i).Fnweight, 2) %> Kgs</td>
	<td><%= FormatNumber(ocoffinvoicedetail.FItemList(i).Fgweight, 2) %> Kgs</td>
	<td></td>
</tr>
<% next %>
</table>

<script type='text/javascript'>

// ������ ���۽� �۵��ϴ� ��ũ��Ʈ
function getOnload(){
	var frm = document.frmMaster;

	RecalcPrice(frm);
}

window.onload = getOnload;

</script>

<%
set ocoffinvoice = Nothing
set ocoffinvoicedetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
