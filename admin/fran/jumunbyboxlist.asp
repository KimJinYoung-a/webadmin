<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ������ŷ����(�ڽ���)
' History : 2011.01.18 �̻� ����
'			2012.08.14 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_balju.asp"-->
<%
dim page, shopid, chulgoyn, showdeleted, showmichulgo, michulgoreason ,statecd, itemid, brandid
dim day5chulgo, shortchulgo, tempshort, danjong, etcshort ,research, i, shopdiv, baljucode ,tmpcartonboxbarcode
dim innerboxno, innerboxsongjangno, cartoonboxno, cartonboxsongjangno ,innerboxbarcode , cartonboxbarcode
dim yyyy1,mm1 , dd1, yyyy2, mm2, dd2, fromDate, toDate ,siteSeq
dim dateType, tplgubun
	menupos = request("menupos")
	page = request("page")
	shopid = request("shopid")
	chulgoyn = request("chulgoyn")
	showdeleted = request("showdel")		'������ ������Ʈ�� �Ķ������ delete ������ �ִ� ��� ���´�.
	showmichulgo = request("showmichulgo")
	michulgoreason = request("michulgoreason")
	statecd = request("statecd")
	itemid = request("itemid")
	brandid = request("brandid")
	shopdiv = request("shopdiv")
	baljucode = request("baljucode")
	day5chulgo = request("day5chulgo")
	shortchulgo = request("shortchulgo")
	tempshort = request("tempshort")
	danjong = request("danjong")
	etcshort = request("etcshort")
	research = request("research")
	innerboxno 			= request("innerboxno")
	innerboxsongjangno 	= request("innerboxsongjangno")
	innerboxbarcode = request("innerboxbarcode")
	cartoonboxno 		= request("cartoonboxno")
	cartonboxsongjangno = request("cartonboxsongjangno")
	cartonboxbarcode = request("cartonboxbarcode")
	dateType = requestCheckVar(request("dateType"),1)
	tplgubun = requestCheckVar(request("tplgubun"),16)

siteSeq = "10"
if (page = "") then
	page = 1
end if

if (research = "") then
	showdeleted = "N"
	michulgoreason = "all"
end if

michulgoreason = "|"
if (day5chulgo = "Y") then
	'5�ϳ����
	michulgoreason = michulgoreason + "5|"
end if
if (shortchulgo = "Y") then
	'������
	michulgoreason = michulgoreason + "S|"
end if
if (tempshort = "Y") then
	'�Ͻ�ǰ��
	michulgoreason = michulgoreason + "T|"
end if
if (danjong = "Y") then
	'����
	michulgoreason = michulgoreason + "D|"
end if
if (etcshort = "Y") then
	'��Ÿ
	michulgoreason = michulgoreason + "E|"
end if

if dateType="" then dateType="B"

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oshopbalju
set oshopbalju = new CShopBalju
	oshopbalju.FRectFromDate = fromDate
	oshopbalju.FRectToDate = toDate
	oshopbalju.FRectDateType = dateType
	oshopbalju.FRectBaljuId = shopid
	oshopbalju.FRectItemid = itemid
	oshopbalju.FRectBrandid = brandid
	oshopbalju.FRectShopdiv = shopdiv
	oshopbalju.FRectBaljucode = baljucode
	oshopbalju.FRectBoxno = innerboxno
	oshopbalju.FRectCartonBoxno = cartoonboxno
	oshopbalju.FRectBoxsongjangno = innerboxsongjangno
	oshopbalju.FRectCartonBoxsongjangno = cartonboxsongjangno
	oshopbalju.frectinnerboxbarcode = innerboxbarcode
	oshopbalju.frectcartonboxbarcode = cartonboxbarcode
	oshopbalju.FtplGubun = tplgubun

	if (statecd = "A") then
		oshopbalju.FRectChulgoYN = "N"
	else
		oshopbalju.FRectStatecd = statecd
	end if

	oshopbalju.FRectShowDeleted = "N"
	'oshopbalju.FRectMichulgoReason = michulgoreason
	oshopbalju.FCurrPage = page
	oshopbalju.Fpagesize = 25
	''oshopbalju.GetShopBaljuByBox
	oshopbalju.GetShopBaljuByBoxNEW
%>

<script language='javascript'>

function regsubmit(page){

	if (frm.innerboxbarcode.value != ''){
		if (frm.innerboxbarcode.value.length != 19){
			alert('InnerBox���ڵ带 ��Ȯ�� �Է��� �ּ���');
			return;
		}

		if (!IsDouble(frm.innerboxbarcode.value)){
			alert('InnerBox���ڵ�� ���ڸ� �Է°��� �մϴ�');
			frm.innerboxbarcode.focus();
			return;
		}
	}

	if (frm.cartonboxbarcode.value != ''){
		if (frm.cartonboxbarcode.value.length != 19){
			alert('cartonBox���ڵ带 ��Ȯ�� �Է��� �ּ���');
			return;
		}

		if (!IsDouble(frm.cartonboxbarcode.value)){
			alert('cartonBox���ڵ�� ���ڸ� �Է°��� �մϴ�');
			frm.cartonboxbarcode.focus();
			return;
		}
	}

	frm.page.value=page;
	frm.submit();
}

function MakeJumun(){
	location.href="jumuninput.asp";
}

function PopSegumil(frm,iidx,comp){
	if (calendarOpen2(comp)){
		if (confirm('������ : ' + comp.value + ' OK?')){
			frm.idx.value = iidx;
			frm.mode.value = "segumil";
			frm.submit();
		}
	};
}

function PopIpgumil(frm,iidx,comp){
	if (calendarOpen2(comp)){
		if (confirm('�Ա��� : ' + comp.value + ' OK?')){
			frm.idx.value = iidx;
			frm.mode.value="ipkumil";
			frm.submit();
		}
	};
}

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('popshopjumunsheet2.asp?idx=' + v ,'shopjumunsheet','width=740,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function ExcelSheet(v){
	window.open('popshopjumunsheet2.asp?idx=' + v + '&xl=on');
}

function MakeReJumun(iidx){
	if (!calendarOpen2(frmMaster.datestr)){ return };

	if (!confirm('��� ������ : ' + frmMaster.datestr.value + ' OK?')){ return };

	if (confirm('�̹�� �ֹ����� �ۼ� �Ͻðڽ��ϱ�?')){
		frmMaster.idx.value = iidx;
		frmMaster.mode.value = "remijumun";
		frmMaster.target = "_blank";
		frmMaster.submit();
	}
}

function MakeReturn(iidx){
	if (!calendarOpen2(frmMaster.datestr)){ return };

	if (!confirm('��� ������ : ' + frmMaster.datestr.value + ' OK?')){ return };

	if (confirm('��ǰ �ֹ����� �ۼ� �Ͻðڽ��ϱ�?')){
		frmMaster.idx.value = iidx;
		frmMaster.mode.value = "returnjumun";
		frmMaster.target = "_blank";
		frmMaster.submit();
	}
}

function MakeDuplicateJumun(iidx){
	if (!calendarOpen2(frmMaster.datestr)){ return };

	if (!confirm('��� ������ : ' + frmMaster.datestr.value + ' OK?')){ return };

	if (confirm('���� �ֹ����� �ۼ� �Ͻðڽ��ϱ�?')){
		frmMaster.idx.value = iidx;
		frmMaster.mode.value = "duplicatejumun";
		frmMaster.target = "_blank";
		frmMaster.submit();
	}
}

function Popbalju(){
	var frm = document.frmlist;
	var idxarr = "";
	for (var i=0;i<frm.elements.length;i++){
		if ((frm.elements[i].name=="ck_all") && (frm.elements[i].checked)){
        	idxarr = idxarr + frm.elements[i].value + ",";
      	}
	}
	if (idxarr==""){
		alert('�ֹ����� �����ϼ���.');
		return;
	}else{
		frm.idxarr.value= idxarr;
		frm.target="_blank";
		frm.action="popoffbaljulist.asp"
		frm.submit();
	}
}

function ModifyBox(frm) {
	if (CheckBox(frm) == true) {
		/*
		if (frm.detailidx.value =="") {
			alert("���������� �Էµ� ������ ���ؼ��� ������ �����մϴ�.");
			return;
		}
		*/
		if (confirm("�Է��Ͻðڽ��ϱ�?") == true) {
			frm.submit();
		}
	}
}

function SetRecv(frm) {
	if (confirm("����Ȯ���Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "setrecv";
		frm.submit();
	}
}

function CheckBox(frm) {
	if (frm.cartoonboxno.value == "") {
		alert("Carton�ڽ���ȣ�� �Է��ϼ���.");
		frm.cartoonboxno.focus();
		return false;
	}

	if (frm.cartoonboxno.value*0 != 0) {
		alert("Carton�ڽ���ȣ�� ���ڸ� �����մϴ�.");
		frm.cartoonboxno.focus();
		return false;
	}

	if (frm.innerboxweight.value == "") {
		frm.innerboxweight.value = 0;
	}

	if (frm.innerboxweight.value*0 != 0) {
		alert("Inner�ڽ� ���Դ� ���ڸ� �����մϴ�.");
		frm.innerboxweight.focus();
		return false;
	}

	if (frm.cartoonboxweight.value == "") {
		frm.cartoonboxweight.value = 0;
	}

	if (frm.cartoonboxweight.value*0 != 0) {
		alert("Carton�ڽ� ���Դ� ���ڸ� �����մϴ�.");
		frm.cartoonboxweight.focus();
		return false;
	}

	return true;
}

function DeleteBox(frm) {
	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "deletedetail";
		frm.submit();
	}
}

function downloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/fran/jumunbyboxlist_excel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		ShopID : 
		<% 'drawSelectBoxOffShop "shopid",shopid %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<select class="select" name="dateType">
			<option value="B" <%= CHKIIF(dateType="B", "selected", "") %> >������</option>
			<option value="C" <%= CHKIIF(dateType="C", "selected", "") %> >�����</option>
		</select> :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		�ֹ����� :
		<select name="statecd" class="select">
			<option value="">��ü
			<option value=" " <% if statecd=" " then response.write "selected" %> >�ۼ���
			<option value="0" <% if statecd="0" then response.write "selected" %> >�ֹ�����
			<option value="1" <% if statecd="1" then response.write "selected" %> >�ֹ�Ȯ��
			<option value="2" <% if statecd="2" then response.write "selected" %> >�Աݴ��
			<option value="5" <% if statecd="5" then response.write "selected" %> >����غ�
			<option value="6" <% if statecd="6" then response.write "selected" %> >�����
			<option value="7" <% if statecd="7" then response.write "selected" %> >���Ϸ�
			<option value="8" <% if statecd="8" then response.write "selected" %> >�԰���
			<option value="9" <% if statecd="9" then response.write "selected" %> >�԰�Ϸ�
			<option value="">========
			<option value="A" <% if statecd="A" then response.write "selected" %> >���������ü
		</select>
	</td>
	<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="regsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�ֹ��ڵ� : <input type="text" class="text" name="baljucode" value="<%= baljucode %>" size="10" maxlength="8">
		�귣�� : <% drawSelectBoxDesignerwithName "brandid", brandid %>
		��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="12">
		3PL ���� : <% Call drawSelectBoxTPLGubunNew("tplgubun", tplgubun) %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		Inner�ڽ���ȣ : <input type="text" class="text" name="innerboxno" value="<%= innerboxno %>" size="10" maxlength="10">
		Inner������ȣ : <input type="text" class="text" name="innerboxsongjangno" value="<%= innerboxsongjangno %>" size="20" maxlength="20">
		InnerBox���ڵ� : <input type="text" class="text" name="innerboxbarcode" value="<%= innerboxbarcode %>" onKeyPress="if(window.event.keyCode==13) regsubmit('');" size="23" maxlength="19">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		Carton�ڽ���ȣ : <input type="text" class="text" name="cartoonboxno" value="<%= cartoonboxno %>" size="10" maxlength="10">
		Carton������ȣ : <input type="text" class="text" name="cartonboxsongjangno" value="<%= cartonboxsongjangno %>" size="20" maxlength="20">
		CartonBox���ڵ� : <input type="text" class="text" name="cartonboxbarcode" value="<%= cartonboxbarcode %>" onKeyPress="if(window.event.keyCode==13) regsubmit('');" size="23" maxlength="19">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
     	����SHOP���� :
     	<input type="radio" name="shopdiv" value="" <% if shopdiv="" then response.write "checked" %> >��ü
		<input type="radio" name="shopdiv" value="direct" <% if shopdiv="direct" then response.write "checked" %> >����
		<input type="radio" name="shopdiv" value="franchisee" <% if shopdiv="franchisee" then response.write "checked" %> >������
		<input type="radio" name="shopdiv" value="foreign" <% if shopdiv="foreign" then response.write "checked" %> >�ؿ�
		<input type="radio" name="shopdiv" value="buy" <% if shopdiv="buy" then response.write "checked" %> >����
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<p>

* �̳ʹڽ� ���Ը� 0 ���� �Է��ϸ� <font color=red>�ؿ������� > ������ �ڽ�</font> ���� ���ܵ˴ϴ�.
<input type="text" name="page" size="2">*10000/<%= oshopbalju.FTotalCount %>&nbsp;<input type="button" onclick="downloadexcel();" value="�����ٿ�ε�" class="button">
<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="17">
			�˻���� : <b><%= oshopbalju.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= oshopbalju.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�����̵�</td>
		<td width=80>������</td>
		<td width=60>Inner<br>�ڽ���ȣ</td>
		<td width=60>Inner<br>�ڽ��߷�</td>
		<td width=60>Carton<br>�ڽ���ȣ</td>
		<td width=60>Carton<br>�ڽ��߷�</td>
		<td>�����ڵ�</td>
		<td>�ֹ��ڵ�</td>
		<td>���ް�</td>
		<td>������</td>
		<td>�����</td>
		<td>Inner<br>������ȣ</td>
		<td>Carton<br>�ù��</td>
		<td>Carton<br>������ȣ</td>
		<td>����Ȯ��</td>
		<td>���</td>
		<td>���ڵ�<br>���</td>
	</tr>
	<% if oshopbalju.FResultCount >0 then %>
	<% for i=0 to oshopbalju.FResultcount-1 %>
	<form name="frmModiPrc_<%= i %>" method="post" action="cartoonbox_process.asp">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="mode" value="modifycartoondetail">
		<input type="hidden" name="masteridx" value="<%= oshopbalju.FItemList(i).Fcartoonmasteridx %>">
		<input type="hidden" name="detailidx" value="<%= oshopbalju.FItemList(i).Fcartoondetailidx %>">
		<input type="hidden" name="baljudate" value="<%= oshopbalju.FItemList(i).Fbaljudate %>">
		<input type="hidden" name="shopid" value="<%= oshopbalju.FItemList(i).Fbaljuid %>">
		<input type="hidden" name="innerboxno" value="<%= oshopbalju.FItemList(i).Fboxno %>">
		<input type="hidden" name="baljunum" value="<%= oshopbalju.FItemList(i).Fbaljunum %>">
		<input type="hidden" name="page" value="<%= page %>">
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oshopbalju.FItemList(i).Fbaljuid %><br><%= oshopbalju.FItemList(i).Fbaljuname %></td>
			<td align="center"><%= oshopbalju.FItemList(i).Fbaljudate %></td>
			<td align="center">
				<%
				if (oshopbalju.FItemList(i).Fboxno <> "0") then
					response.write oshopbalju.FItemList(i).Fboxno
				end if
				%>
			</td>
			<td align="center">
				<%
				if (oshopbalju.FItemList(i).Finnerboxweight <> "") then
					oshopbalju.FItemList(i).Finnerboxweight = FormatNumber(oshopbalju.FItemList(i).Finnerboxweight, 2)
				end if
				%>
				<input type="text" class="text" name="innerboxweight" value="<%= oshopbalju.FItemList(i).Finnerboxweight %>" size="3" maxlength="6" style="text-align:right">
			</td>
			<td align="center">
				<input type="text" class="text" name="cartoonboxno" value="<%= oshopbalju.FItemList(i).Fcartoonboxno %>" size="3" maxlength="6" style="text-align:right">
			</td>
			<td align="center">
				<%
				if (oshopbalju.FItemList(i).Fcartoonboxweight <> "") then
					oshopbalju.FItemList(i).Fcartoonboxweight = FormatNumber(oshopbalju.FItemList(i).Fcartoonboxweight, 2)
				end if
				%>
				<input type="text" class="text" name="cartoonboxweight" value="<%= oshopbalju.FItemList(i).Fcartoonboxweight %>" size="3" maxlength="6" style="text-align:right">
			</td>
			<td align="center"><%= oshopbalju.FItemList(i).Fbaljunum %></td>
			<td align="center"><%= oshopbalju.FItemList(i).Fbaljucode %></td>
			<td align="center"><%= FormatNumber(oshopbalju.FItemList(i).Ftotsuplycash, 0) %></td>
			<td align="center">
				<font color="<%= oshopbalju.FItemList(i).GetStateColor %>"><%= oshopbalju.FItemList(i).GetStateName %></font>
			</td>
			<td align="center"><%= oshopbalju.FItemList(i).Fchulgodate %></td>
			<td align="center">
				<input type="text" class="text" name="innerboxsongjangno" value="<%= oshopbalju.FItemList(i).Fboxsongjangno %>" size="16" maxlength="20" style="text-align:right">
			</td>
			<td align="center">
				<% drawSelectBoxDeliverCompany "cartonboxsongjangdiv", oshopbalju.FItemList(i).Fcartonboxsongjangdiv %>
			</td>
			<td align="center">
				<input type="text" class="text" name="cartonboxsongjangno" value="<%= oshopbalju.FItemList(i).Fcartonboxsongjangno %>" size="16" maxlength="20" style="text-align:right">
				<% if (oshopbalju.FItemList(i).Ffindurl <> "") then %>
				<input type="button" class="button" value="����" onclick="window.open('<%= oshopbalju.FItemList(i).Ffindurl + oshopbalju.FItemList(i).Fcartonboxsongjangno %>');">
				<% end if %>
			</td>
			<td align="center">
				<% If Not IsNull(oshopbalju.FItemList(i).FshopReceive) Then %>
				<% If (oshopbalju.FItemList(i).FshopReceive = "N") Then %>
				<input type="button" class="button" value=" Ȯ�� " onClick="SetRecv(frmModiPrc_<%= i %>)">
				<% Else %>
				<%= oshopbalju.FItemList(i).FshopReceiveUserID %>
				<% End If %>
				<% End If %>
			</td>
			<td align="center">
				<input type="button" class="button" value=" ���� " onClick="ModifyBox(frmModiPrc_<%= i %>)">
				<% if (oshopbalju.FItemList(i).Fcartoondetailidx <> "") then %>
				&nbsp;
				<!--
					 <input type="button" class="button" value=" ���� " onClick="DeleteBox(frmModiPrc_<%= i %>)">
				   -->
				<% end if %>
			</td>
			<td align="center">
				<input type="button" class="button" value="���" onclick="printbarcode_off('PACKING', '', '', '', '', '', '<%= oshopbalju.FItemList(i).Fordermasteridx %>', '<%= oshopbalju.FItemList(i).Fboxno %>', '');">
			</td>
		</tr>
	</form>
	<% next %>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="17" align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
	<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="17" align="center">
			<%
			dim strparam
			strparam = "&shopid=" + CStr(shopid) + "&yyyy1=" + CStr(yyyy1) + "&mm1=" + CStr(mm1) + "&dd1=" + CStr(dd1) + "&yyyy2=" + CStr(yyyy2) + "&mm2=" + CStr(mm2) + "&dd2=" + CStr(dd2)

			strparam = strparam + "&menupos=" + CStr(menupos)
			strparam = strparam + "&chulgoyn=" + CStr(chulgoyn)
			strparam = strparam + "&showdel=" + CStr(showdeleted)
			strparam = strparam + "&showmichulgo=" + CStr(showmichulgo)
			strparam = strparam + "&michulgoreason=" + Server.URLEncode(CStr(michulgoreason))

			strparam = strparam + "&statecd=" + CStr(statecd)
			strparam = strparam + "&itemid=" + CStr(itemid)
			strparam = strparam + "&brandid=" + CStr(brandid)
			strparam = strparam + "&shopdiv=" + CStr(shopdiv)
			strparam = strparam + "&baljucode=" + CStr(baljucode)

			strparam = strparam + "&day5chulgo=" + CStr(day5chulgo)
			strparam = strparam + "&shortchulgo=" + CStr(shortchulgo)
			strparam = strparam + "&tempshort=" + CStr(tempshort)
			strparam = strparam + "&danjong=" + CStr(danjong)
			strparam = strparam + "&etcshort=" + CStr(etcshort)

			%>
			<% if oshopbalju.HasPreScroll then %>
			<a href="?page=<%= oshopbalju.StartScrollPage-1 %>&research=on<%= strparam %>">[pre]</a>
			<% else %>
			[pre]
			<% end if %>

			<% for i=0 + oshopbalju.StartScrollPage to oshopbalju.FScrollCount + oshopbalju.StartScrollPage - 1 %>
			<% if i>oshopbalju.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&research=on<%= strparam %>">[<%= i %>]</a>
			<% end if %>
			<% next %>

			<% if oshopbalju.HasNextScroll then %>
			<a href="?page=<%= i %>&research=on<%= strparam %>">[next]</a>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set oshopbalju = Nothing
%>
<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
