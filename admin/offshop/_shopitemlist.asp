<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������λ�ǰ ���
' Hieditor : 2009.04.07 ������ ����
'			 2010.06.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer,page,usingyn , research,pricediff,imageview, pricelow ,itemgubun, itemid, itemname
dim cdl, cdm, cds ,onexpire ,i, PriceDiffExists , IsDirectIpchulContractExistsBrand ,publicbarcode
dim centermwdiv, onlineMwDiv, readonlyyn, isupcheitemreg
	onlineMwDiv  	= RequestCheckVar(request("onlineMwDiv"),1)
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),9)
	usingyn     = RequestCheckVar(request("usingyn"),1)
	research    = RequestCheckVar(request("research"),9)
	pricediff   = RequestCheckVar(request("pricediff"),9)
	pricelow    = RequestCheckVar(request("pricelow"),9)
	imageview   = RequestCheckVar(request("imageview"),9)
	onexpire    = RequestCheckVar(request("onexpire"),9)
	itemgubun   = RequestCheckVar(request("itemgubun"),2)
	itemid      = RequestCheckVar(request("itemid"),9)
	itemname    = RequestCheckVar(request("itemname"),32)
	publicbarcode    = RequestCheckVar(request("publicbarcode"),20)
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
	centermwdiv = RequestCheckVar(request("centermwdiv"),3)
	if page="" then page=1
	if research<>"on" then usingyn="Y"

readonlyyn = "N"
isupcheitemreg = false

if C_ADMIN_USER then

'/����
elseif (C_IS_SHOP) then
	'//�������϶�
	if C_IS_OWN_SHOP then
	else
	end if

	readonlyyn = "Y"
else
	'/��ü�� ��� ���̵� �ھƳ���
	if C_IS_Maker_Upche then
		designer = session("ssBctId")
		IsDirectIpchulContractExistsBrand = fnIsDirectIpchulContractExistsBrand(designer)
		isupcheitemreg = getupcheitemregyn(designer)
	end if

	readonlyyn = "Y"
end if

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 100
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = designer
	ioffitem.FRectOnlyUsing = usingyn
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemID = itemid
	ioffitem.FRectItemName = html2db(itemname)
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.FRectOnlineExpiredItem = onexpire
	ioffitem.FRectpublicbarcode = publicbarcode
    ioffitem.FRectCenterMwdiv = centermwdiv
	ioffitem.FRectOnlineMwDiv = onlineMwDiv

	if pricediff="on" then
	    ioffitem.FRectPriceRow = pricelow
		ioffitem.GetOffShopPriceDiffItemList
	else
		ioffitem.GetOffNOnLineShopItemList
	end if

%>
<script type="text/javascript">

function NotUsingCheckAll(){
    var frm;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
		    if (frm.isusing[0].checked==true){
		        frm.isusing[1].checked = true;
		        frm.cksel.checked = true;
		        AnCheckClick(frm.cksel);
		    }
		}
	}
}

//����
function pop_itemedit_off_edit(ibarcode){
    <% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) and Not isupcheitemreg then %>
        alert('������ �����ϴ�. - ���� ���� �԰� �귣�常 ���� �����մϴ�.');
    	return;
	<% else %>
		var pop_itemedit_off_edit = window.open('/common/offshop/item/pop_itemedit_off_edit.asp?barcode=' + ibarcode,'pop_itemedit_off_edit','width=1024,height=768,resizable=yes,scrollbars=yes');
		pop_itemedit_off_edit.focus();
    <% end if %>
}

//���
function pop_itemedit_off_new(){
	var pop_itemedit_off_new;

    <% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) then %>
    	<% if not(isupcheitemreg) then %>
        	alert('������ �����ϴ�. - ���� �������������� ������ �ּ���.');
        	return;
		<% else %>
			if (confirm('�ڵ������ ��� ������ ������ ������ �¶��ο� ��ϵ��ְų�\n���������� ��ǰ��\n\n----------------����------------- \n\n������� ���� �ּ���. ����Ͻðڽ��ϱ�?')){
				pop_itemedit_off_new = window.open('/common/offshop/item/pop_itemedit_off_edit.asp','pop_itemedit_off_new','width=1024,height=768,scrollbars=yes,resizable=yes');
				pop_itemedit_off_new.focus();
			}
    	<% end if %>
	<% else %>
		if (confirm('�ڵ������ ��� ������ ������ ������ �¶��ο� ��ϵ��ְų�\n���������� ��ǰ��\n\n----------------����------------- \n\n������� ���� �ּ���. ����Ͻðڽ��ϱ�?')){
			pop_itemedit_off_new = window.open('/common/offshop/item/pop_itemedit_off_edit.asp','pop_itemedit_off_new','width=1024,height=768,scrollbars=yes,resizable=yes');
			pop_itemedit_off_new.focus();
		}
    <% end if %>
}

<% if C_ADMIN_USER then %>
	function pop_item_multi_add_off() {
		var pop_item_multi_add_off;

		if (confirm('�ڵ������ ��� ������ ������ ������ �¶��ο� ��ϵ��ְų�\n���������� ��ǰ��\n\n----------------����------------- \n\n������� ���� �ּ���. ����Ͻðڽ��ϱ�?')) {
			pop_item_multi_add_off = window.open('/common/offshop/item/pop_item_multi_add_off.asp','pop_item_multi_add_off','width=1024,height=768,scrollbars=yes,resizable=yes');
			pop_item_multi_add_off.focus();
		}
	}
<% end if %>

function ReSearch(page){
	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('��ǰ��ȣ�� ���ڸ� �����մϴ�.');
			frm.itemid.focus();
			return;
		}
	}

	frm.page.value = page;
	frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function ChargeIdAvail(ichargeid){
	var comp = document.frm.designer;

	if (ichargeid=="10x10"){
		return true
	}

	for (var i=0;i<comp.length;i++){
		if (comp[i].value==ichargeid){
			return true
		}
	}

	return false;
}

function ModiArr(){
	var upfrm = document.frmArrupdate;
	var frm; var str; var j; var checkStr;
	var pass = false;

<% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) then %>
        alert('������ �����ϴ�. - ���� ���� �԰� �귣�常 ���� �����մϴ�.');
        return;
<% else %>
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;
	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.itempricearr.value = "";
	upfrm.itemsuplyarr.value = "";
	upfrm.onofflinkynarr.value = "";
	upfrm.extbarcodearr.value = "";
	upfrm.shopbuypricearr.value = "";
    upfrm.orgsellpricearr.value = "";
    upfrm.isusingarr.value = "";
    upfrm.centermwdivarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (frm.extbarcode.value != ''){
					str = frm.extbarcode.value;
					for (j=0; j<str.length; j++){
						checkStr = str.charAt(j);
						if(/\W/.test(checkStr) && /[^\s]/.test(checkStr)){
							alert("������ڵ忡 Ư�����ڴ� ������� �ʽ��ϴ�.");
							frm.extbarcode.focus();
							return;
						}
					}

					if (frm.extbarcode.value.length < 8){
						alert('������ڵ� ���̰� �ʹ� ª���ϴ�(8�� �̻�).\n���� ���ڵ尡 �ִ°�츸 �Է��� �ּ���');
						frm.extbarcode.focus();
						return;
					}
				}

				if (frm.itemgubun.value == "80") {
					// ����ǰ
					if (frm.tx_sellcash.value > 0) {
						alert("����ǰ�� �ǸŰ��� 0���Ͽ��� �մϴ�.");
					    frm.tx_sellcash.focus();
					    return;
					}

					if (frm.tx_orgsellprice.value > 0) {
						alert("����ǰ�� �Һ��ڰ��� 0���Ͽ��� �մϴ�.");
					    frm.tx_orgsellprice.focus();
					    return;
					}
				}else if (frm.itemgubun.value == "60") {
	                if (frm.tx_orgsellprice.value.substr(0,1) != '-'){
						frm.tx_orgsellprice.value = "-"+frm.tx_orgsellprice.value
					}
	                if (frm.tx_sellcash.value.substr(0,1) != '-'){
						frm.tx_sellcash.value = "-"+frm.tx_sellcash.value
					}
				} else {
	                if (!IsDigit(frm.tx_orgsellprice.value)){
						alert('�Һ��ڰ����� ���ڸ� �����մϴ�.');
						frm.tx_orgsellprice.focus();
						return;
					}

					if (!IsDigit(frm.tx_sellcash.value)){
						alert('�ǸŰ��� ���ڸ� �����մϴ�.');
						frm.tx_sellcash.focus();
						return;
					}

                    <% if C_ADMIN_USER then %>
						if (!IsDigit(frm.tx_suplycash.value)){
							alert('���԰��� ���ڸ� �����մϴ�.');
							frm.tx_suplycash.focus();
							return;
						}

						if (!IsDigit(frm.tx_shopbuyprice.value)){
							alert('������ް��� ���ڸ� �����մϴ�.');
							frm.tx_shopbuyprice.focus();
							return;
						}
                    <% end if %>

					// �Ϲݻ�ǰ
					if (frm.tx_sellcash.value<10){
						if (!confirm('�ǸŰ��� 10������ Ŀ�� �մϴ�. ��� �����Ͻðڽ��ϱ�?')){
						    frm.tx_sellcash.focus();
						    return;
						}
					}

	                if (frm.tx_orgsellprice.value*1<frm.tx_sellcash.value*1){
						alert('�Һ��ڰ��� �ǸŰ����� Ŀ���մϴ�..');
						frm.tx_orgsellprice.focus();
						return;
					}

					<% if C_ADMIN_USER then %>
		                // ���԰� ���ް� üũ
		                if ((frm.tx_suplycash.value*1!=0)&&(frm.tx_suplycash.value*1!=0)){
		                    if ((frm.tx_suplycash.value*1>frm.tx_shopbuyprice.value*1)&&(frm.tx_shopbuyprice.value*1!=0)){  //���ް�0 ���԰� ���� �������� (���ް�0�ΰ�� ��ǥ��������)
		    					alert('�� ���ް��� ���԰����� Ŀ���մϴ�..');
		    					frm.tx_suplycash.focus();
		    					return;
		    				}
						}
					<% end if %>
				}

				if (frm.centermwdiv.value == ''){
					alert("���͸��Ա����� ������ �ȵǾ����ϴ�.");
					frm.centermwdiv.focus();
					return;
				}

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.orgsellpricearr.value = upfrm.orgsellpricearr.value + frm.tx_orgsellprice.value + "|";
				upfrm.itempricearr.value = upfrm.itempricearr.value + frm.tx_sellcash.value + "|";

				<% if C_ADMIN_USER then %>
					upfrm.itemsuplyarr.value = upfrm.itemsuplyarr.value + frm.tx_suplycash.value + "|";
					upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.tx_shopbuyprice.value + "|";

					if (frm.onofflinkyn[0].checked){
						upfrm.onofflinkynarr.value = upfrm.onofflinkynarr.value + frm.onofflinkyn[0].value + "|";
					}else if (frm.onofflinkyn[1].checked){
						upfrm.onofflinkynarr.value = upfrm.onofflinkynarr.value + frm.onofflinkyn[1].value + "|";
					}
				<% end if %>

				upfrm.centermwdivarr.value = upfrm.centermwdivarr.value + frm.centermwdiv.value + "|";
				upfrm.extbarcodearr.value = upfrm.extbarcodearr.value + frm.extbarcode.value + "|";

				<% if C_ADMIN_USER or C_IS_Maker_Upche then %>
					if (frm.isusing[0].checked){
						upfrm.isusingarr.value = upfrm.isusingarr.value + "Y" + "|";
					}else{
						upfrm.isusingarr.value = upfrm.isusingarr.value + "N" + "|";
					}
				<% end if %>
			}
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
		upfrm.mode.value = "arrmodi";
		upfrm.submit();
	}
<% end if %>
}

//���͸��Ա��� �ϰ�����
function CheckAllcentermwdiv(){
    var frmlist;
    var pass = false;

	if (frm.checkallcentermwdiv.value=="") {
	    alert('�ϰ������Ͻ� ���� ���͸��Ա����� �����ϼ���.');
	    frm.checkallcentermwdiv.focus();
	    return false;
	}

	for (var i=0;i<document.forms.length;i++){
		frmlist = document.forms[i];
		if (frmlist.name.substr(0,9)=="frmBuyPrc") {
			if (frmlist.cksel.checked){
				pass = true;
				frmlist.centermwdiv.value=frm.checkallcentermwdiv.value;
			}
		}
	}

	var ret;
	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}
}

function samePriceALL(){
    var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
		    samePrice(frm);
		}
	}
}

function samePrice(frm){
    frm.tx_orgsellprice.value=frm.oldonlineorgprice.value*1 + frm.oldonlineOptAddprice.value*1;  //�Һ��ڰ�
	frm.tx_sellcash.value=frm.oldonlineprice.value*1 + frm.oldonlineOptAddprice.value*1;         //�ǸŰ�

	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function EventPrice0(){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (frm.tx_discountsellprice.value!=0){
				frm.tx_discountsellprice.value=0;
				frm.cksel.checked=true;
				AnCheckClick(frm.cksel);
			}
		}
	}
}

function BuyPrice0(){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
    			if (frm.tx_suplycash.value!=0){
    				frm.tx_suplycash.value=0;
    				frm.cksel.checked=true;
    				AnCheckClick(frm.cksel);
    			}
			}
		}
	}
}

function ShopSuplyPrice0(){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
    			if (frm.tx_shopbuyprice.value!=0){
    				frm.tx_shopbuyprice.value=0;
    				frm.cksel.checked=true;
    				AnCheckClick(frm.cksel);
    			}
    		}
		}
	}
}

function qrcode_view(itemgubun, itemoption, itemid){
	var qrcode_view = window.open('/common/qrcode/qrcode_itemid_view.asp?itemgubun='+itemgubun+'&itemoption='+itemoption+'&itemid='+itemid,'qrcode_view','width=700,height=700,scrollbars=yes,resizable=yes');
	qrcode_view.focus();
}

function jsAlertNoAuth(msg) {
	alert(msg);
	//return false;
}

function downloadexcel(){
	alert("ok");
    document.frm.target = "view"; 
    document.frm.action = "/admin/offshop/shopitemlist_excel.asp";  
	document.frm.submit();
    document.frm.target = ""; 
    document.frm.action = "";  
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣�� :
		<% if C_IS_Maker_Upche then %>
			<%= designer %>
			<input type="hidden" name="designer" value="<%= designer %>">
		<% else %>
			<% drawSelectBoxDesignerwithName "designer",designer  %>
		<% end if %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;
		��ǰ����:<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
		&nbsp;
		ON���Ա��� :
		<select class="select" name="onlineMwDiv">
			<option value="">����</option>
			<option value="M" <% if (onlineMwDiv = "M") then %>selected<% end if %> >M</option>
			<option value="W" <% if (onlineMwDiv = "W") then %>selected<% end if %> >W</option>
			<option value="U" <% if (onlineMwDiv = "U") then %>selected<% end if %> >U</option>
			<option value="X" <% if (onlineMwDiv = "X") then %>selected<% end if %> >��Ÿ</option>
		</select>
     	&nbsp;
     	���͸��Ա���:
     	<select class="select" name="centermwdiv">
	        <option value="">����</option>
	        <option value="M" <%= CHKIIF(centermwdiv="M","selected","")%> >����</option>
	        <option value="W" <%= CHKIIF(centermwdiv="W","selected","")%> >Ư��</option>
	        <option value="X" <%= CHKIIF(centermwdiv="X","selected","")%> >������</option>
        </select>
		&nbsp;
		�������:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="ReSearch('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		��ǰ�ڵ� : <input type="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" style="IME-MODE: disabled" />
		&nbsp;
		��ǰ�� : <input type="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
		&nbsp;
		������ڵ� : <input type="text" name="publicbarcode" value="<%= publicbarcode %>" size="20" maxlength="20">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >�̹�������
		&nbsp;
		<input type="checkbox" name="pricediff" value="on" <% if pricediff="on" then response.write "checked" %> >���ݻ��̸� ����
		&nbsp;
		<input type="checkbox" name="pricelow" value="on" <% if pricelow="on" then response.write "checked" %> >�¶��κ��� ��������
		&nbsp;
		<input type="checkbox" name="onexpire" value="on" <% if onexpire="on" then response.write "checked" %> >ONǰ��+����+������(�Ż�ǰ����)
	</td>
</tr>
</table>
<!-- �˻� �� -->

<br>
�� ������ �����ǰ�� ���� �̹��� ����� �ʼ��� ����Ǿ����ϴ�.<br>
�� ��Ȱ�� �ֹ� ����ó���� ���� �̹��� ���� ��ǰ�� ���� <b>�̹����� ���</b>�� �ּ���<br>
�� ��ǰ�������� �����Ϸ��� ��ǰ��ȣ�� �����ּ���.
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<% if C_ADMIN_USER or C_IS_Maker_Upche then %>
			<% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) then %>
				<% if not(isupcheitemreg) then %>
					<input type="button" class="button" value="������������ ��ǰ���(��ü)" onclick="jsAlertNoAuth('������ �����ϴ�. - ���� �������������� ������ �ּ���.')">
				<% Else %>
					<input type="button" class="button" value="������������ ��ǰ���(��ü)" onclick="pop_itemedit_off_new()">
				<% End If %>
			<% Else %>
				<input type="button" class="button" value="������������ ��ǰ���(����,Ư��)" onclick="pop_itemedit_off_new()">
			<% End If %>
			<% if C_ADMIN_USER then %>
				<input type="button" class="button" value="������������ ��ǰ �ϰ����(����)" onclick="pop_item_multi_add_off()">
			<% end if %>
		<% end if %>
	</td>
	<td align="right">
		<input type="button" onclick="downloadexcel();" value="�����ٿ�ε�" class="button">
		<% if ioffitem.FresultCount>0 then %>
			<% if C_ADMIN_USER then %>
				<input type="button" class="button" value="���û�ǰ ���԰� 0 ����" onclick="BuyPrice0()">
				<input type="button" class="button" value="���û�ǰ �����ް� 0 ����" onclick="ShopSuplyPrice0()">
			<% end if %>
		<% end if %>
		<input type="button" class="button" value="���û�ǰ �ϰ�����" onclick="ModiArr()">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= ioffitem.FTotalcount %></b>
		&nbsp;&nbsp;
		<% if ioffitem.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>

		<b><%= page %> / <%= ioffitem.FTotalpage %></b>

		<% if (ioffitem.FTotalpage - ioffitem.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>

		<% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) then %>
        	�� ���� ���� �԰� �귣�常 ���� �����մϴ�.
        <% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>

	<% if (imageview<>"") then %>
		<td width="50">�̹���</td>
	<% end if %>

	<td width="70">�귣��ID</td>
	<td width="90">��ǰ�ڵ�</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>

	<% if C_ADMIN_USER then %>
		<td width="20"><input type="button" value=">" onclick="samePriceALL();"></td>
	<% end if %>

	<td width="60">�Һ��ڰ�</td>
	<td width="60">�ǸŰ�</td>
	<td width="40">������<br>(%)</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td width="60">���԰�</td>
		<td width="60">������ް�</td>
		<td width="30">����<br>����</td>
		<td width="30">����<br>����</td>
	<% end if %>

	<td width="30">ON<br>����<br>����</td>
	<td width="80">
		����<br>���Ա���
		
		<% if C_ADMIN_AUTH or C_OFF_AUTH then %>
			<Br>
	     	<select class="select" name="checkallcentermwdiv" >
		        <option value="">����</option>
		        <option value="M">����</option>
		        <option value="W">Ư��</option>
	        </select>
	        <br><input class="button" type="button" value="��������" onClick="CheckAllcentermwdiv();">
	    <% end if %>
	</td>
	<td width="30">ON<br>�Ǹ�</td>
	<td width="30">ON<br>����</td>
	<td width="100">������ڵ�</td>

	<% if C_ADMIN_USER or C_IS_Maker_Upche then %>
		<td width="60">��� ����<br><input class="button" type="button" value="������" onClick="NotUsingCheckAll();"></td>
	<% end if %>
	<% if C_ADMIN_USER then %>
		<td width="50">ON/OFF<br>���ݿ���</td>
	<% end if %>

	<td>���</td>
</tr>
</form>

<% if ioffitem.FresultCount>0 then %>
	<% for i=0 to ioffitem.FresultCount -1 %>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
	<input type="hidden" name="oldonlineprice" value="<%= ioffitem.FItemlist(i).FOnLineItemprice %>">
	<input type="hidden" name="oldonlineorgprice" value="<%= ioffitem.FItemlist(i).FOnLineItemOrgprice %>">
	<input type="hidden" name="oldonlineOptAddprice" value="<%= ioffitem.FItemlist(i).FOnlineOptaddprice %>">

	<% if ioffitem.FItemlist(i).Fisusing="N" then %>
		<tr bgcolor="#EEEEEE">
	<% else %>
		<tr bgcolor="#FFFFFF">
	<% end if %>

		<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>

		<% if (imageview<>"") then %>
			<td width="50">
				<img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0>
			</td>
		<% end if %>

		<td ><%= ioffitem.FItemlist(i).FMakerID %></td>
		<td>
		    <% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) and Not isupcheitemreg then %>
				<a href="javascript:jsAlertNoAuth('������ �����ϴ�. - ���� ���� �԰� �귣�常 ���� �����մϴ�');" onfocus="this.blur()">
			<% Else %>
				<a href="javascript:pop_itemedit_off_edit('<%= ioffitem.FItemlist(i).GetBarCode %>')" onfocus="this.blur()">
			<% End If %>

			<%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %>
			</a>
		</td>
		<td>
		    <% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) and Not isupcheitemreg then %>
				<a href="javascript:jsAlertNoAuth('������ �����ϴ�. - ���� ���� �԰� �귣�常 ���� �����մϴ�');" onfocus="this.blur()">
			<% Else %>
				<a href="javascript:pop_itemedit_off_edit('<%= ioffitem.FItemlist(i).GetBarCode %>')" onfocus="this.blur()">
			<% End If %>

			<%= ioffitem.FItemlist(i).FShopItemName %>
			</a>
		</td>
		<td>
			<%= ioffitem.FItemlist(i).FShopitemOptionname %>

			<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
			    <br>�ɼ��߰��ݾ�: <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>
			<% end if %>
		</td>
	    <% PriceDiffExists = false %>

	    <% if C_ADMIN_USER then %>
			<td align="center" >
			    <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
				    <% if (ioffitem.FItemlist(i).FOnlineitemorgprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemOrgprice) or (ioffitem.FItemlist(i).FOnLineItemprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemprice) then %>
					    <input type="button" class="button" value=">" onclick="samePrice(frmBuyPrc_<%= i %>);">
					    <% PriceDiffExists = true %>
				    <% end if %>
			    <% end if %>
			</td>
		<% end if %>

	    <td align="right" >
	        <input type="text" class="text" name="tx_orgsellprice" <% if readonlyyn = "Y" then response.write " readonly" %> value="<%= ioffitem.FItemlist(i).FShopItemOrgprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">

	        <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
		        <% if (ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemOrgprice)  then %>
		            <font color="red"><strong><%= ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
		        <% else %>
		            <% if (PriceDiffExists) then %>
						<%= ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice %>
		            <% end if %>
		        <% end if %>
	        <% end if %>
	    </td>
		<td align="right" >
		    <input type="text" class="text" name="tx_sellcash" <% if readonlyyn = "Y" then response.write " readonly" %> value="<%= ioffitem.FItemlist(i).FShopItemprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">

		    <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
		        <% if (ioffitem.FItemlist(i).FOnLineItemprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemprice)  then %>
			        <font color="red"><strong><%= ioffitem.FItemlist(i).FOnLineItemprice + ioffitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
			    <% else %>
			        <% if (PriceDiffExists) then %>
						<%= ioffitem.FItemlist(i).FOnLineItemprice + ioffitem.FItemlist(i).FOnlineOptaddprice %>
			        <% end if %>
		        <% end if %>
	        <% end if %>
		</td>
		<td align="center" >
	        <% if (ioffitem.FItemlist(i).FShopItemOrgprice<>0) then %>
	            <% if ioffitem.FItemlist(i).FShopItemOrgprice<>ioffitem.FItemlist(i).FShopItemprice then %>
					OFF:<font color="#FF3333"><%= CLng((ioffitem.FItemlist(i).FShopItemOrgprice-ioffitem.FItemlist(i).FShopItemprice)/ioffitem.FItemlist(i).FShopItemOrgprice*100*100)/100 %></font>
	            <% end if %>
		    <% end if %>

		    <% if (ioffitem.FItemlist(i).FOnlineitemorgprice<>0) then %>
		        <% if ioffitem.FItemlist(i).FOnlineitemorgprice<>ioffitem.FItemlist(i).FOnLineItemprice then %>
					ON:<font color="#FF3333"><%= CLng((ioffitem.FItemlist(i).FOnlineitemorgprice-ioffitem.FItemlist(i).FOnLineItemprice)/ioffitem.FItemlist(i).FOnlineitemorgprice*100*100)/100 %></font>
	            <% end if %>
		    <% end if %>
		</td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right" >
				<input type="text" name="tx_suplycash" <% if readonlyyn = "Y" then response.write " readonly" %> value="<%= ioffitem.FItemlist(i).Fshopsuplycash %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
			</td>
			<td align="right" >
				<input type="text" name="tx_shopbuyprice" <% if readonlyyn = "Y" then response.write " readonly" %> value="<%= ioffitem.FItemlist(i).Fshopbuyprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
			</td>
			<td align="right" >
				<% if (ioffitem.FItemlist(i).FShopItemprice<>0) and (ioffitem.FItemlist(i).Fshopsuplycash<>0) then %>
					<font color="blue"><%= CLng((ioffitem.FItemlist(i).FShopItemprice-ioffitem.FItemlist(i).Fshopsuplycash)/ioffitem.FItemlist(i).FShopItemprice*100) %>%</font>
				<% end if %>
			</td>
			<td align="right" >
				<% if (ioffitem.FItemlist(i).FShopItemprice<>0) and (ioffitem.FItemlist(i).Fshopbuyprice<>0) then %>
					<font color="blue"><%= CLng((ioffitem.FItemlist(i).FShopItemprice-ioffitem.FItemlist(i).Fshopbuyprice)/ioffitem.FItemlist(i).FShopItemprice*100) %>%</font>
				<% end if %>
		    </td>
		<% end if %>

		<td align="center" ><%= ioffitem.FItemlist(i).FmwDiv %></td>
	    <td align="center" >
	    	<% if ioffitem.FItemlist(i).Fstockitemid = 0 or C_ADMIN_AUTH or C_OFF_AUTH then %>
		     	<select class="select" name="centermwdiv">
			        <option value="">����</option>
			        <option value="M" <%= CHKIIF(ioffitem.FItemlist(i).Fcentermwdiv="M","selected","")%> >����</option>
			        <option value="W" <%= CHKIIF(ioffitem.FItemlist(i).Fcentermwdiv="W","selected","")%> >Ư��</option>
		        </select>
		    <% else %>
		    	<%= ioffitem.FItemlist(i).Fcentermwdiv %>
				<input type="hidden" name="centermwdiv" value="<%= ioffitem.FItemlist(i).Fcentermwdiv %>">
			<% end if %>

	        <% if (ioffitem.FItemlist(i).FmwDiv="W" or ioffitem.FItemlist(i).FmwDiv="M") and (ioffitem.FItemlist(i).FmwDiv<>ioffitem.FItemlist(i).FCenterMwDiv) then %>
	            <br><font color='red'>�¶��ΰ�����</font></strong>
	        <% end if %>
	    </td>
	    <td align="center" ><%= fnColor(ioffitem.FItemlist(i).Fsellyn,"sellyn") %></td>
	    <td align="center" ><%= fnColor(ioffitem.FItemlist(i).FonLineDanjongyn,"dj") %></td>
		<td align="right" >
			<input type="text" name="extbarcode" value="<%= ioffitem.FItemlist(i).FextBarcode %>" size="12" maxlength="20" style="border:1px #999999 solid; " onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
		</td>

		<% if C_ADMIN_USER or C_IS_Maker_Upche then %>
			<td align="left" >
				<% if ioffitem.FItemlist(i).Fisusing="Y" then %>
					<input type="radio" name="isusing" value="Y" checked onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
					<input type="radio" name="isusing" value="N" onclick="CheckThis(frmBuyPrc_<%= i %>)">N
				<% else %>
					<input type="radio" name="isusing" value="Y" onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
					<input type="radio" name="isusing" value="N" checked onclick="CheckThis(frmBuyPrc_<%= i %>)"><font color="red">N</font>
				<% end if %>
			</td>
		<% end if %>

		<% if C_ADMIN_USER then %>
			<td align="center">
				<input type="radio" name="onofflinkyn" value="Y" <% if ioffitem.FItemlist(i).fonofflinkyn="Y" then response.write " checked" %> onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
				<input type="radio" name="onofflinkyn" value="N" <% if ioffitem.FItemlist(i).fonofflinkyn="N" then response.write " checked" %> onclick="CheckThis(frmBuyPrc_<%= i %>)">N
			</td>
		<% end if %>

		<td align="center">
			<% if ioffitem.FItemlist(i).Fitemgubun="10" then %>
				<input type="button" onclick="qrcode_view('<%= ioffitem.FItemlist(i).Fitemgubun %>','<%= ioffitem.FItemlist(i).Fitemoption %>','<%= ioffitem.FItemlist(i).Fshopitemid %>');" value="QR" class="button">
			<% end if %>
		</td>
	</tr>
	</form>
	<% next %>

	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">
		<% if ioffitem.HasPreScroll then %>
			<a href="javascript:ReSearch('<%= ioffitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
			<% if i>ioffitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:ReSearch('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ioffitem.HasNextScroll then %>
			<a href="javascript:ReSearch('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<% if C_ADMIN_USER or C_IS_Maker_Upche then %>
			<% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) then %>
				<% if not(isupcheitemreg) then %>
					<input type="button" class="button" value="������������ ��ǰ���" onclick="alert('������ �����ϴ�. - ���� �������������� ������ �ּ���.');return false;">
				<% Else %>
					<input type="button" class="button" value="������������ ��ǰ���" onclick="pop_itemedit_off_new()">
				<% End If %>
			<% Else %>
				<input type="button" class="button" value="������������ ��ǰ���" onclick="pop_itemedit_off_new()">
			<% End If %>
		<% end if %>
	</td>
	<td align="right">
		<% if ioffitem.FresultCount>0 then %>
			<% if C_ADMIN_USER then %>
				<input type="button" class="button" value="���û�ǰ ���԰� 0 ����" onclick="BuyPrice0()">
				<input type="button" class="button" value="���û�ǰ �����ް� 0 ����" onclick="ShopSuplyPrice0()">
			<% end if %>
		<% end if %>
		<input type="button" class="button" value="���û�ǰ �ϰ�����" onclick="ModiArr()">
	</td>
</tr>
</table>
<!-- �׼� �� -->
<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<form name="frmArrupdate" method="post" action="/admin/offshop/shopitem_process.asp" style="margin:0px;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="onofflinkynarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="orgsellpricearr" value="">
	<input type="hidden" name="itempricearr" value="">
	<input type="hidden" name="itemsuplyarr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="extbarcodearr" value="">
	<input type="hidden" name="centermwdivarr" value="">
</form>
<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
