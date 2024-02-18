<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����� ���� ����̵�
' History : 2018.02.07 �̻� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%

'// �ݾ׼��� �۾� �ȵǾ�����!!!, skyer9, 2018-02-12
dim PriceEditEnable : PriceEditEnable = False

'// sellcash : �ǸŰ�, buycash : ���԰�, suplycash : �� ���ް�
'// �̿� �ٸ� ��Ī�� ������� �ʴ´�.


dim i, j, k
dim shopid, moveshopid, makerid
dim itemgubunarr, itemidarr, itemoptionarr, sellcasharr, buycasharr, suplycasharr, itemnoarr, itemnamearr, itemoptionnamearr
dim scheduledt, songjangdiv, songjangno, comment
dim masteridx, mode
dim ojumunmaster, ojumundetail

masteridx  = requestCheckVar(request("masteridx"), 32)
mode  = requestCheckVar(request("mode"), 32)

shopid  = requestCheckVar(request("shopid"), 32)
moveshopid  = requestCheckVar(request("moveshopid"), 32)
makerid  = requestCheckVar(request("makerid"), 32)

scheduledt  = requestCheckVar(request("scheduledt"), 32)
songjangdiv  = requestCheckVar(request("songjangdiv"), 32)
songjangno  = requestCheckVar(request("songjangno"), 32)
comment  = request("comment")

if (C_IS_Maker_Upche) then
	makerid = session("ssBctID")
end if

set ojumunmaster = new COrderSheet
set ojumundetail = new COrderSheet

if (masteridx <> "") then

	ojumunmaster.FRectIdx = masteridx
	ojumunmaster.GetOneOrderSheetMaster_IMSI

	shopid = ojumunmaster.FOneItem.Ftargetid
	moveshopid = ojumunmaster.FOneItem.Fbaljuid

	if (scheduledt = "") then
		scheduledt = ojumunmaster.FOneItem.Fscheduledate
	end if

	if (songjangdiv = "") then
		songjangdiv = ojumunmaster.FOneItem.Fsongjangdiv
	end if

	if (songjangno = "") then
		songjangno = ojumunmaster.FOneItem.Fsongjangno
	end if

	if (comment = "") then
		comment = ojumunmaster.FOneItem.Fcomment
	end if

	ojumundetail.FRectIdx = masteridx
	ojumundetail.GetOrderSheetDetail_IMSI
end if


if (shopid = "") then
	response.write "�߸��� �����Դϴ�."
	dbget.close : response.end
end if


dim IsAddItemOK : IsAddItemOK = False
dim errMsg : errMsg = ""

if (shopid <> "") and (moveshopid <> "") and (makerid <> "") then
	IsAddItemOK = IsSameShopContract(shopid, moveshopid, makerid)
	if (IsAddItemOK <> True) then
		errMsg = "����� ���ų� �θ����� �귣�� ��ึ���� �ٸ��ϴ�."
		response.write "<script>alert('" & errMsg & "')</script>"
	end if
end if


if (mode = "additem") then
	'
end if

%>
<script>
function jsSetShopidMove() {
	var frm = document.frm;
	if (frm.moveshopidsel.value == "") {
		alert("���� ����(��������)�� �����ϼ���.");
		return;
	}

	if (frm.shopid.value == frm.moveshopidsel.value) {
		alert("����!!\n\n��߸���� ���������� �����մϴ�.");
		return;
	}

	frm.moveshopid.value = frm.moveshopidsel.value;
	frm.moveshopidsel.value = "";

	frm.submit();
}

function jsSetMakerid() {
	var frm = document.frm;
	if (frm.makeridsel.value == "") {
		alert("���� �귣�带 �����ϼ���.");
		return;
	}

	if (frm.makerid.value == frm.makeridsel.value) {
		alert("����!!\n\n������ �귣���Դϴ�.");
		return;
	}

	frm.makerid.value = frm.makeridsel.value;
	frm.makeridsel.value = "";

	frm.submit();
}

function jsChkForm(frm) {
	if (frm.shopid.value == "") {
		alert("����!!\n\n��߸��� �����ȵ�.");
		return false;
	}

	if (frm.moveshopid.value == "") {
		alert("����!!\n\n�������� �����ȵ�.");
		return false;
	}

	if (frm.scheduledt.value.length<1){
		alert('����̵����� �Է��ϼ���');
		calendarOpen3(frm.scheduledt,'����̵����� �Է��ϼ���','');
		return false;
	}

	if (frm.songjangdiv.value.length<1){
		alert('�ù�縦 ���� �ϼ���');
		frm.songjangdiv.focus();
		return false;
	}

	if (frm.songjangno.value.length<1){
		alert('���� ��ȣ�� �Է� �ϼ���');
		msfrm.songjangno.focus();
		return false;
	}

	return true;
}

function jsStockMove() {
	var frm = document.frm;

	if (jsChkForm(frm) != true) { return; }

	<% if (ojumundetail.FResultCount < 1) then %>
	if (frm.itemgubunarr.value == "") {
		alert("�߰��� ��ǰ�� �����ϴ�.");
		return;
	}
	<% end if %>

	var ret = confirm('�Է��ϽŴ�� ��� �̵�ó�� �Ͻðڽ��ϱ�?');
	if (ret) {
		frm.mode.value = "saveorder";
		frm.method = "post";
		frm.action = "pop_jumun_move_process.asp";
		frm.submit();
	}
}

function jsSaveModified() {
	var frm = document.frm;

	if (jsChkForm(frm) != true) { return; }

	frm.itemgubunarr.value = "";
	frm.itemidarr.value = "";
	frm.itemoptionarr.value = "";
	frm.itemnoarr.value = "";

	for (var i = 0; i < document.forms.length ; i++){
		o = document.forms[i];
		if (o.name.substr(0,9)=="frmBuyPrc") {
			if (!IsInteger(o.itemno.value)){
				alert('������ ������ �����մϴ�.');
				o.itemno.focus();
				return;
			}

		    if (o.itemno.value < 0){
				alert("������ 0�̻� ��� �˴ϴ�.");
				o.itemno.focus();
				return;
			}

			frm.itemgubunarr.value = frm.itemgubunarr.value + o.itemgubun.value + "|";
			frm.itemidarr.value = frm.itemidarr.value + o.itemid.value + "|";
			frm.itemoptionarr.value = frm.itemoptionarr.value + o.itemoption.value + "|";
			frm.itemnoarr.value = frm.itemnoarr.value + o.itemno.value + "|";
			frm.sellcasharr.value = frm.sellcasharr.value + o.sellcash.value + "|";
			frm.buycasharr.value = frm.buycasharr.value + o.buycash.value + "|";
			frm.suplycasharr.value = frm.suplycasharr.value + o.suplycash.value + "|";
		}
	}

	var ret = confirm('�����Ͻðڽ��ϱ�?');
	if (ret) {
		frm.mode.value = "additem";
		frm.method = "post";
		frm.action = "pop_jumun_move_process.asp";
		frm.submit();
	}
}

function jsButtonDisabled() {
	document.getElementById("btnMove").disabled = true;
}

function jsAddItems() {
	var frm = document.frm;

	if (frm.shopid.value == "") {
		alert("����!!\n\n��߸��� �����ȵ�.");
		return;
	}

	if (frm.makerid.value == "") {
		alert("���� �귣�带 �����ϼ���.");
		return;
	}

	var popwin;
	popwin = window.open('/common/offshop/popshopitemV2.asp?shopid=' + frm.shopid.value + '&chargeid=' + frm.makerid.value,'jsAddItems','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ReActItems(igubun,iitemid,iitemoption,isellcash,isuplycash,ishopbuyprice,iitemno,iitemname,iitemoptionname,iitemdesigner) {
	var frm = document.frm;

	<% '// sellcash : �ǸŰ�, buycash : ���԰�, suplycash : �� ���ް� %>
	frm.itemgubunarr.value = igubun;
	frm.itemidarr.value = iitemid;
	frm.itemoptionarr.value = iitemoption;
	frm.sellcasharr.value = isellcash;
	frm.buycasharr.value = isuplycash;
	frm.suplycasharr.value = ishopbuyprice;
	frm.itemnoarr.value = iitemno;

	frm.mode.value = "additem";
	frm.method = "post";
	frm.action = "pop_jumun_move_process.asp";
	frm.submit();
}

function AddItemsBarCode() {
	var frm = document.frm;

	if (frm.shopid.value == "") {
		alert("����!!\n\n��߸��� �����ȵ�.");
		return;
	}

	if (frm.makerid.value == "") {
		alert("���� �귣�带 �����ϼ���.");
		return;
	}

	var popwin = window.open('popshopjumunitemBybarcode.asp?shopid=' + frm.shopid.value + '&suplyer=' + frm.makerid.value + '&digitflag=MV','AddItemsBarCode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<form name="frm" method="get" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="moveshopid" value="<%= moveshopid %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="25" width="150" bgcolor="<%= adminColor("tabletop") %>">��߸���</td>
	<td>
		<%= shopid %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="25" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td>
		<% if (moveshopid = "") then %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("moveshopidsel",moveshopid, "21") %>
		<% if (moveshopid = "") then %>
		&nbsp;
		<input type="button" value="������������" onClick="jsSetShopidMove()" class="button">
		* <font color="red">���� ���������� �����ϼ���.</font>
		<% end if %>
		<% else %>
		<%= moveshopid %>
		<% end if %>
	</td>
</tr>
<% if (shopid <> "") and (moveshopid <> "") then %>
<tr bgcolor="#FFFFFF">
	<td height="25" bgcolor="<%= adminColor("tabletop") %>">���� �귣��</td>
	<td>
		<% if (makerid = "") then %>
		<% Call drawSelectBoxDesignerwithName("makeridsel",makerid) %>
		&nbsp;
		<input type="button" value="�귣�� ����" onClick="jsSetMakerid()" class="button">
		* <font color="red">���� �귣�带 �����ϼ���.</font>
		<% else %>
		<%= makerid %>
		<% if (errMsg <> "") then %>
		&nbsp;
		<font color="red">* <%= errMsg %></font>
		<% end if %>
		<% end if %>
	</td>
</tr>
<% if (makerid <> "") then %>
<tr bgcolor="#FFFFFF">
	<td height="25" bgcolor="<%= adminColor("tabletop") %>">�귣�� ����</td>
	<td>
		<% Call drawSelectBoxDesignerwithName("makeridsel","") %>
		<% if (shopid <> "") and (moveshopid <> "") then %>
		&nbsp;
		<input type="button" value="�귣�� ����" onClick="jsSetMakerid()" class="button">
		* <font color="red">�ٸ� �귣��</font>�� ��ǰ�� �����Ϸ��� ���� �귣�带 �����ϼ���.
		<% end if %>
	</td>
</tr>
<% end if %>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">
		����̵���
	</td>
	<td>
		<input type="text" class="text" name="scheduledt" value="<%= scheduledt %>" size=10 readonly ><a href="javascript:calendarOpen(frm.scheduledt);">
		<img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>

			�ù�� :<% drawSelectBoxDeliverCompany "songjangdiv", songjangdiv %>
			�����ȣ:<input type="text" class="text" name="songjangno" size=14 maxlength=16 value="<%= songjangno %>" >
			<br>
			(�ù�� ������ ������� �ù��:��Ÿ���� �����ȣ:�����, ������� ���� �Է� �Ͻø� �˴ϴ�.)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ��û����</td>
	<td>
		<textarea name="comment" class="textarea" cols="80" rows="6"><%= comment %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" colspan="2" align="center">
	    <input type="button" value="����̵�ó��" onClick="jsStockMove()" class="button" id="btnMove">
		&nbsp;
		<input type="button" value="�����ϱ�" onClick="jsSaveModified()" class="button" id="btnSave">
	</td>
</tr>
</table>
</form>

<% if (IsAddItemOK) then %>
<p></p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left" valign="bottom">
		�� ������ �÷����� �ֹ� �Ͻø�, ��߸���(<font color="red">���̳ʽ��ֹ�</font>)�� ��������(<font color="red">�԰��ֹ�</font>)�� �ֹ��� ���� �����˴ϴ�
	</td>
	<td align="right">
		<input type="button" class="button" value="����(���ڵ�)" onclick="AddItemsBarCode()">
		&nbsp;
		<input type="button" class="button" value="��ǰ�߰�" onclick="jsAddItems()">
	</td>
</tr>
</table>
<!-- �׼� �� -->
<% else %>
<p />
* ���� �������� �� �귣�带 �����ϼ���.
<% end if %>

<p></p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= ojumundetail.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">���ڵ�</td>
	<td>�귣��</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width="80">�ǸŰ�</td>
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	    <td width="60">�ٹ�����<br>���԰�</td>
	    <td width="60">����<br>���ް�</td>
	<% elseif (C_IS_Maker_Upche) then %>
		<td width="60">�ٹ�����<br>���ް�</td>
	<% else %>
		<td width="60">����<br>���ް�</td>
	<% end if %>
	<td width="60">����</td>
</tr>
<% if (ojumundetail.FResultCount > 0) then %>
<% for i = 0 to ojumundetail.FResultCount - 1 %>
<form name="frmBuyPrc_<%= i %>" method="post" action="">
<input type="hidden" name="detailidx" value="<%= ojumundetail.FItemList(i).Fidx %>">
<input type="hidden" name="itemgubun" value="<%= ojumundetail.FItemList(i).Fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ojumundetail.FItemList(i).Fitemid %>">
<input type="hidden" name="itemoption" value="<%= ojumundetail.FItemList(i).Fitemoption %>">
<% if Not (PriceEditEnable) then %>
<input type="hidden" name="sellcash" value="-1">
<input type="hidden" name="buycash" value="-1">
<input type="hidden" name="suplycash" value="-1">
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td ><%= ojumundetail.FItemList(i).Fitemgubun %><%= CHKIIF(ojumundetail.FItemList(i).Fitemid >= 1000000,format00(8,ojumundetail.FItemList(i).Fitemid),format00(6,ojumundetail.FItemList(i).Fitemid)) %><%= ojumundetail.FItemList(i).Fitemoption %></td>
	<td ><%= ojumundetail.FItemList(i).FMakerid %></td>
	<td align="left"><%= ojumundetail.FItemList(i).Fitemname %></td>
	<td ><%= ojumundetail.FItemList(i).Fitemoptionname %></td>
	<% if Not (PriceEditEnable) then %>
		<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></td>
		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fbuycash,0) %></td><!-- ���԰� -->
			<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fsuplycash,0) %></td><!--���� ���ް�-->
		<% elseif (C_IS_Maker_Upche) then %>
			<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fbuycash,0) %></td><!-- ���԰� -->
		<% else %>
			<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fsuplycash,0) %></td><!--���� ���ް�-->
		<% end if %>
	<% else %>
		<td ><input type="text" class="text" name="sellcash" value="<%= ojumundetail.FItemList(i).Fsellcash %>" size="8" maxlength="8"></td>
		<td ><input type="text" class="text" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>" size="8" maxlength="8"></td>
		<td ><input type="text" class="text" name="suplycash" value="<%= ojumundetail.FItemList(i).Fsuplycash %>" size="8" maxlength="8"></td>
	<% end if %>
	<td ><input type="text" class="text" name="itemno" value="<%= ojumundetail.FItemList(i).Fbaljuitemno %>"  size="4" maxlength="4" onKeyDown="jsButtonDisabled()"></td>
</tr>
</form>
<% next %>
<% end if %>

</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
