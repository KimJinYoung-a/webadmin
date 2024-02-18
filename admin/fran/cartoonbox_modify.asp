<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������ �ֹ��� �ۼ�
' History : 2009.04.07 ������ ����
'			2022.07.22 �ѿ�� ����(Ȧ���� ī��ڽ� ���� �߰�, ���Ȱ�ȭ, �ҽ�ǥ��ȭ)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<%
dim idx, mode, listpageurl, enclistpageurl, i, j, currcartoonboxno, isnewcartoonbox, currbaljudate
dim suminnerboxweight, sumcartoonboxNweight, sumcartoonboxweight, sumemsprice, sumsupplyPrice, sumcartoonboxcbm
dim shopid
	menupos = RequestCheckVar(request("menupos"),32)
	idx = RequestCheckVar(request("idx"),32)
	mode = RequestCheckVar(request("mode"),32)
	enclistpageurl = RequestCheckVar(request("enclistpageurl"), 256)
	shopid = RequestCheckVar(request("shopid"),32)

if (enclistpageurl = "") then
	listpageurl = request.ServerVariables("HTTP_REFERER")
	enclistpageurl = request.ServerVariables("HTTP_REFERER")

	enclistpageurl = Replace(enclistpageurl, "?", "_Q_")
	enclistpageurl = Replace(enclistpageurl, "&", "_A_")
	enclistpageurl = Replace(enclistpageurl, "=", "_E_")
	enclistpageurl = Replace(enclistpageurl, "/", "_B_")
else
	listpageurl = enclistpageurl

	listpageurl = Replace(listpageurl, "_Q_", "?")
	listpageurl = Replace(listpageurl, "_A_", "&")
	listpageurl = Replace(listpageurl, "_E_", "=")
	listpageurl = Replace(listpageurl, "_B_", "/")
end if

if idx="" then idx=0

dim ocartoonboxmaster
set ocartoonboxmaster = new CCartoonBox
	ocartoonboxmaster.FRectMasterIdx = idx
	ocartoonboxmaster.GetMasterOne

dim ocartoonboxdetail
set ocartoonboxdetail = new CCartoonBox
	ocartoonboxdetail.FRectMasterIdx = idx
	ocartoonboxdetail.FRectShopid = ocartoonboxmaster.FOneItem.Fshopid
	ocartoonboxdetail.GetDetailList   ''[db_storage].dbo.uf_getCartonBoxPrice �̺κ� �Լ� ������ ����.

dim oinnerboxlist
set oinnerboxlist = new CCartoonBox
	oinnerboxlist.FRectMasterIdx = -1
	oinnerboxlist.FRectShopid = ocartoonboxmaster.FOneItem.Fshopid

	if (idx = 0) then
		oinnerboxlist.GetInnerBoxList  '// ��������, ''�ּ�ó�� 2016/08/31 eastone
	end if

dim oBaljuList
set oBaljuList = new CCartoonBox
oBaljuList.FRectMasterIdx = idx
oBaljuList.GetBaljuList

%>

<script type="text/javascript">

function SaveMaster(frm) {
	/*
	if (frm.title.value == "") {
		alert("�۾����� �Է��ϼ���.");
		frm.title.focus();
		return;
	}
	*/
	if (frm.shopid.value == "") {
		alert("���� �����ϼ���.");
		frm.shopid.focus();
		return;
	}
	if (frm.deliverpay.value != "") {
		frm.deliverpay.value = frm.deliverpay.value.replace(/,/g, "");

		if (frm.deliverpay.value*0 != 0) {
			alert("EMS����� ���ڸ� �Է� �����մϴ�.");
			frm.deliverpay.focus();
			return;
		}
	}
	if (frm.workstate[2].checked == true) {
		if (frm.deliverdt.value == "") {
			alert('������� �Է��� �ּ���.');
			frm.deliverdt.focus();
			if (!calendarOpen2(frm.deliverdt)) { return };
		}
	}

	if (frm.masteridx.value*1 == 0) {
		frm.detailidxarr.value = "-1";
		for (var i=0;i<document.forms.length;i++){
			frmarr = document.forms[i];
			if (frmarr.name.substr(0,12)=="frmSelectPrc") {
				if (frmarr.cksel.checked == true) {
					if (frm.shopid.value != frmarr.shopid.value) {
						alert("���� �ٸ� ������ ���õǾ����ϴ�.");
						return;
					}
					frm.detailidxarr.value = frm.detailidxarr.value + "," + frmarr.detailidx.value;
				}
			}
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret == true) {
		if (frm.masteridx.value*1 == 0) {
			frm.mode.value="newmaster";
		} else {
			frm.mode.value="savemaster";
		}
		frm.submit();
	}
}

function DelMaster(frm) {
	var ret = confirm('��ü���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.mode.value="delmaster";
		frm.submit();
	}
}

function GotoListPage() {
	if ("<%= listpageurl %>" != "") {
		location.href = "<%= listpageurl %>";
	}
}

/*
function ModifyBox(frm) {
	if (CheckBox(frm) == true) {
		frm.submit();
	}
}
*/

function CheckBox(frm) {
	if (frm.cartoonboxno.value == "") {
		alert("Cartoon�ڽ���ȣ�� �Է��ϼ���.");
		frm.cartoonboxno.focus();
		return false;
	}

	if (frm.innerboxno.value == "") {
		alert("Inner�ڽ���ȣ�� �Է��ϼ���.");
		frm.innerboxno.focus();
		return false;
	}

	if (frm.cartoonboxno.value*0 != 0) {
		alert("Cartoon�ڽ���ȣ�� ���ڸ� �����մϴ�.");
		frm.cartoonboxno.focus();
		return false;
	}

	if (frm.innerboxno.value*0 != 0) {
		alert("Inner�ڽ���ȣ�� ���ڸ� �����մϴ�.");
		frm.innerboxno.focus();
		return false;
	}

	if (frm.cartoonboxweight.value == "") {
		frm.cartoonboxweight.value = 0;
	}

	if (frm.cartoonboxweight.value*0 != 0) {
		alert("Cartoon�ڽ� ���Դ� ���ڸ� �����մϴ�.");
		frm.cartoonboxweight.focus();
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

	return true;
}

function SaveSelectArr() {
	var upfrm = document.frmadd;
	var frm;
	var pass = false;
	var ret;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,12)=="frmSelectPrc") {
			pass = ((pass) || (frm.cksel.checked == true));
		}
	}

	if (pass != true) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.detailidxarr.value = "";
	upfrm.cartoonboxnoarr.value = "";
	upfrm.innerboxnoarr.value = "";
	upfrm.innerboxweightarr.value = "";

	upfrm.baljudatearr.value = "";
	upfrm.shopidarr.value = "";


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,12)=="frmSelectPrc") {
			if (frm.cksel.checked == true) {
				if (CheckBox(frm) != true) {
					return;
				}

				upfrm.detailidxarr.value = upfrm.detailidxarr.value + "|" + frm.detailidx.value;
				upfrm.cartoonboxnoarr.value = upfrm.cartoonboxnoarr.value + "|" + frm.cartoonboxno.value;
				upfrm.innerboxnoarr.value = upfrm.innerboxnoarr.value + "|" + frm.innerboxno.value;
				upfrm.innerboxweightarr.value = upfrm.innerboxweightarr.value + "|" + frm.innerboxweight.value;

				upfrm.baljudatearr.value = upfrm.baljudatearr.value + "|" + frm.baljudate.value;
				upfrm.shopidarr.value = upfrm.shopidarr.value + "|" + frm.shopid.value;
			}
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		upfrm.mode.value = "saveselectedbox";
		upfrm.submit();
	}
}

function DeleteSelectArr() {
	var upfrm = document.frmadd;
	var frm;
	var pass = false;
	var ret;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,12)=="frmSelectPrc") {
			pass = ((pass) || (frm.cksel.checked == true));
		}
	}

	if (pass != true) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.detailidxarr.value = "0";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,12)=="frmSelectPrc") {
			if (frm.cksel.checked == true) {
				upfrm.detailidxarr.value = upfrm.detailidxarr.value + "," + frm.detailidx.value;
			}
		}
	}

	if (confirm('������ ���� �Ͻðڽ��ϱ�?')){
		upfrm.mode.value = "deleteselectedbox";
		upfrm.submit();
	}
}

function DeselectArr() {
	var upfrm = document.frmadd;
	var frm;
	var pass = false;
	var ret;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,10)=="frmModiPrc") {
			pass = ((pass) || (frm.cksel.checked == true));
		}
	}

	if (pass != true) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.detailidxarr.value = "0";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,10)=="frmModiPrc") {
			if (frm.cksel.checked == true) {
				upfrm.detailidxarr.value = upfrm.detailidxarr.value + "," + frm.detailidx.value;
			}
		}
	}

	if (confirm('������ ���� �Ͻðڽ��ϱ�?')){
		upfrm.mode.value = "deselectbox";
		upfrm.submit();
	}
}

function SaveDetailArr() {
	var upfrm = document.frmadd;
	var frm;
	var ret;

	upfrm.detailidxarr.value = "";
	upfrm.cartoonboxnoarr.value = "";
	upfrm.cartoonboxweightarr.value = "";
	upfrm.cartoonboxTypearr.value = "";
	upfrm.cartonboxsongjangnoarr.value = "";
	upfrm.innerboxnoarr.value = "";
	upfrm.innerboxweightarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,10)=="frmModiPrc") {

			if (CheckBox(frm) != true) {
				return;
			}

			frm.cksel.checked = true;

			upfrm.detailidxarr.value = upfrm.detailidxarr.value + "|" + frm.detailidx.value;

			upfrm.cartoonboxnoarr.value = upfrm.cartoonboxnoarr.value + "|" + frm.cartoonboxno.value;
			upfrm.cartoonboxweightarr.value = upfrm.cartoonboxweightarr.value + "|" + frm.cartoonboxweight.value;
			upfrm.cartoonboxTypearr.value = upfrm.cartoonboxTypearr.value + "|" + frm.cartoonboxType.value;
			upfrm.cartonboxsongjangnoarr.value = upfrm.cartonboxsongjangnoarr.value + "|" + frm.cartonboxsongjangno.value;
			upfrm.innerboxnoarr.value = upfrm.innerboxnoarr.value + "|" + frm.innerboxno.value;
			upfrm.innerboxweightarr.value = upfrm.innerboxweightarr.value + "|" + frm.innerboxweight.value;
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		upfrm.mode.value = "savedetailarr";
		upfrm.submit();
	}
}

function CalcCartoonboxWeight(frmforcalc) {
	var upfrm = document.frmadd;
	var frm;
	var pass = false;
	var ret;
	var cartoonboxno = frmforcalc.cartoonboxno.value*1;
	var sumweight = 0;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,10)=="frmModiPrc") {
			if (frm.cartoonboxno.value*1 == cartoonboxno) {
				sumweight = sumweight + frm.innerboxweight.value*1;
			}
		}
	}

	frmforcalc.cartoonboxweight.value = sumweight.toFixed(2);
}

function PopBoxItemList(shopid, yyyy, mm, dd, boxno) {
	var popurl = "/admin/fran/jumunbyboxitemlist.asp?research=on&shopid=" + shopid + "&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=" + dd + "&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + dd + "&boxno=" + boxno;

	var w = window.open(popurl);
	w.focus();
}

function PopBoxSelect(masteridx) {
	var popurl = "popoffinvoice_selectbox.asp?masteridx=" + masteridx;

	var w = window.open(popurl);
	w.focus();
}

function chkAllitem(frmname, frmlength) {
    var frm;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,frmlength)==frmname) {
		    frm.cksel.checked = true;
		    AnCheckClick(frm.cksel);
		}
	}
}

function jsDisableOtherShop() {
	var isCheckExist = false;
	var checkedShopid;

	for (var i=0;i<document.forms.length;i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,12)=="frmSelectPrc") {
			if (frm.cksel.checked == true) {
				isCheckExist = true;
				checkedShopid = frm.shopid.value;
				break;
			}
		}
	}

	for (var i=0;i<document.forms.length;i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,12)=="frmSelectPrc") {
			if (isCheckExist == false) {
				frm.cksel.disabled = false;
			} else if (frm.shopid.value != checkedShopid) {
				frm.cksel.disabled = true;
			}
		}
	}

	if (isCheckExist == true) {
		document.frmMaster.shopid.value = checkedShopid;
	} else {
		document.frmMaster.shopid.value = "";
	}
}

function PrintDetailItemList(jungsanidx, shopid, shopname) {
	var popwin;
	popwin = window.open('/admin/fran/popcartonboxitemlist_print.asp?jungsanidx=' + jungsanidx + '&shopid=' + shopid + '&shopname=' + shopname + '&xl=Y','PrintDetailItemList','width=850,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

/*
function PopOpenInvoice(invoiceidx) {
	var popwin;
	popwin = window.open('/admin/fran/offinvoice_modify.asp?idx=' + invoiceidx,'PopOpenInvoice','width=850,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
*/

function PopOpenInvoice(invoiceidx,isxl) {
	var popwin;

	popwin = window.open( '/admin/fran/popoffinvoice_print.asp?idx=' + invoiceidx + '&xl='+isxl,'PopOpenInvoice','width=850,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopOpenPackingList(invoiceidx,isxl) {
	var popwin;

	popwin = window.open('/admin/fran/popoffinvoice_print_packinglist.asp?idx=' + invoiceidx + '&xl='+isxl,'PopOpenPackingList','width=850,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function downloadOrder(masteridx, baljucode, shopid, cartoonboxmasteridx) {
    var chkimg = (document.frmMaster.chkimg.checked) ? "on" : ""; //2017/06/26 �߰�

	var popwin = window.open("/common/popOrderSheet_foreign_excel.asp?masteridx=" + masteridx + "&baljucode=" +baljucode + "&shopid=" +shopid + "&cartoonboxmasteridx=" +cartoonboxmasteridx + "&chkimg="+chkimg,"ExcelOfflineOrderSheet","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// qs �˾�
function PopOpenQS(invoiceidx, jungsanidx, workidx, loginsite, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/quotationsheet.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenQSList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// pi �˾�
function PopOpenPI(invoiceidx, jungsanidx, workidx, loginsite, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/proformainvoice.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite+ '&cunit='+cunit+'&xl=Y&tpl='+tpl,'PopOpenInvoice','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ci �˾�
function PopOpenCI(invoiceidx, jungsanidx, workidx, loginsite, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/commercialinvoice.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite + '&cunit='+cunit+'&tpl='+tpl,'PopOpenCIList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// pl �˾�
function PopOpenPL(invoiceidx, jungsanidx, workidx, loginsite,boxidx, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/packlinglist.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite+'&boxidx='+boxidx+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenPLList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// pl ��ǰ �˾�
function PopOpenPLItem(invoiceidx,loginsite,boxidx, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/packlingItemlist.asp?idx=' + invoiceidx+'&ls='+ loginsite+'&boxidx='+boxidx+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenPLIList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popJungsanMaster(iid){
	var popwin = window.open('/admin/offshop/franmeaippopsubmaster.asp?idx=' + iid,'popjungsan','width=1280, height=960, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popinvocesMaster(iid, workidx, shopid){
	if (workidx != "") {
		<% if IsNull(ocartoonboxmaster.FOneItem.Ftotsuplycash) then %>
		alert('���� ���ް�(����) �� ���ΰ�ħ�ϼ���.');
		return;
		<% end if %>
	}
	var url = '/admin/fran/offinvoice_modify.asp?menupos=<%= menupos %>';
	if (iid != '') {
		url = url + '&idx=' + iid;
	}
	if (workidx != '') {
		url = url + '&workidx=' + workidx;
	}
	if (shopid != '') {
		url = url + '&shopid=' + shopid;
	}
	var popwin = window.open(url,'popinvoces','width=1280, height=960, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popViewBalju(baljunum, baljuid) {
	var popwin = window.open('/admin/fran/baljufinish_offline_new.asp?baljunum=' + baljunum + '&baljuid=' + baljuid,'popViewBalju','width=1280, height=960, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function jsRefreshSupplyCash() {
	var frm = document.frmMaster;
	frm.mode.value = "refreshsupplycash";
	frm.submit();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="cartoonbox_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="masteridx" value="<%= idx %>">
<input type="hidden" name="detailidxarr" value="">
<input type="hidden" name="enclistpageurl" value="<%= enclistpageurl %>">

<!-- ��ܹ� ���� -->
<tr height="30" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�۾�����</strong></font>
			    </td>
				<td align=right>
					<input type="button" class="button" value="������� �̵�" onclick="GotoListPage()">
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >IDX</td>
	<td>
		<%= ocartoonboxmaster.FOneItem.Fidx %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�۾���</td>
	<td>
		<% if (ocartoonboxmaster.FOneItem.Fidx = "") then %>
			�۾����� ����� �ڵ��Էµ˴ϴ�.
			<input type="hidden" name="title" value="">
		<% else %>
			<input type="text" class="text" name="title" value="<%= ocartoonboxmaster.FOneItem.Ftitle %>" size=60 maxlength=100>
		<% end if %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >�����̵�</td>
	<% if idx <> 0 then %>
	<input type=hidden name="shopid" value="<%= ocartoonboxmaster.FOneItem.Fshopid %>">
	<td><%= ocartoonboxmaster.FOneItem.Fshopid %></td>
	<% else %>
	<td><% drawSelectBoxOffShopNot000 "shopid", ocartoonboxmaster.FOneItem.Fshopid %></td>
	<% end if %>
	<td bgcolor="<%= adminColor("tabletop") %>" >����</td>
	<td>
		<%= ocartoonboxmaster.FOneItem.Fshopname %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
	<td>
		<input type=radio name="workstate" value="5" <% if (ocartoonboxmaster.FOneItem.Fworkstate="5" or ocartoonboxmaster.FOneItem.Fworkstate="") then response.write "checked" %> >��ŷ��
		<input type=radio name="workstate" value="6" <% if ocartoonboxmaster.FOneItem.Fworkstate="6" then response.write "checked" %> >�����
		<input type=radio name="workstate" value="7" <% if ocartoonboxmaster.FOneItem.Fworkstate="7" then response.write "checked" %> >���Ϸ�
		<% if (ocartoonboxmaster.FOneItem.Fjungsanidx <> "") then %>
			&nbsp;
			<font color=blue>�����Է�</font>
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�����</td>
	<td>
		<input type="text" class="text" name="deliverdt" value="<%= ocartoonboxmaster.FOneItem.Fdeliverdt %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.deliverdt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
	</td>
</tr>

<% if ocartoonboxmaster.FOneItem.getcartoonboxpaymentstatus<>"" then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">wholesale��������</td>
		<td colspan="3">
			<%= ocartoonboxmaster.FOneItem.getcartoonboxpaymentstatus %>
			<br>���ڹ߼�:
			<br><%= left(ocartoonboxmaster.FOneItem.fsmssenddate,10) %>
			<br><%= mid(ocartoonboxmaster.FOneItem.fsmssenddate,12,22) %>
		</td>
	</tr>
<% end if %>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���ù���</td>
	<td>
		<% for i=0 to oBaljuList.FResultCount-1 %>
		<a href="javascript:popViewBalju(<%= oBaljuList.FItemList(i).Fbaljukey %>, '<%= oBaljuList.FItemList(i).Fshopid %>')">
			<%= CHKIIF(i>0, ", ", "") %><%= oBaljuList.FItemList(i).Fbaljukey %><%= CHKIIF(oBaljuList.FItemList(i).FnotfinishCnt>0, "(�������)", "") %>
		</a>
		<% next %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" ></td>
	<td>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��۹��</td>
	<td>
		<select class="select" name="delivermethod">
			<option value=''>����</option>
			<option value='E' <% if (ocartoonboxmaster.FOneItem.Fdelivermethod = "E") then %>selected<% end if %>>EMS</option>
			<option value='D' <% if (ocartoonboxmaster.FOneItem.Fdelivermethod = "D") then %>selected<% end if %>>DHL</option>
			<option value='F' <% if (ocartoonboxmaster.FOneItem.Fdelivermethod = "F") then %>selected<% end if %>>�װ�</option>
			<option value='S' <% if (ocartoonboxmaster.FOneItem.Fdelivermethod = "S") then %>selected<% end if %>>�ؿ�</option>
			<option value='P' <% if (ocartoonboxmaster.FOneItem.Fdelivermethod = "P") then %>selected<% end if %>>��������(����)</option>
			<option value='T' <% if (ocartoonboxmaster.FOneItem.Fdelivermethod = "T") then %>selected<% end if %>>�����ù�</option>
		</select>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >EMS���</td>
	<td>
		<input type="text" class="text" name="deliverpay" value="<%= ocartoonboxmaster.FOneItem.Fdeliverpay %>" size=15 maxlength=100>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >���ް�(����)</td>
	<td>
		<% if Not IsNull(ocartoonboxmaster.FOneItem.Ftotsuplycash) then %>
			<%= FormatNumber(ocartoonboxmaster.FOneItem.Ftotsuplycash, 0) %>
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�ؿܰ��ް�(����)</td>
	<td>
		<% if Not IsNull(ocartoonboxmaster.FOneItem.Ftotforeign_suplycash) then %>
			<%= FormatNumber(ocartoonboxmaster.FOneItem.Ftotforeign_suplycash,2) %>
		<% end if %>
		<% if Not IsNull(ocartoonboxmaster.FOneItem.FjumuncurrencyUnit) then %>
			&nbsp;<%= ocartoonboxmaster.FOneItem.FjumuncurrencyUnit %>
		<% end if %>
		<% if idx <> "" and idx <> "0" then %>
		<input type="button" class="button" value="���ΰ�ħ" onClick="jsRefreshSupplyCash();">
		<% end if %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >�����ڵ�</td>
	<td>
		<%= ocartoonboxmaster.FOneItem.Fjungsanidx %>

		<% if (ocartoonboxmaster.FOneItem.Fjungsanidx <> "") then %>
			&nbsp;
			<input type="button" class="button" value="��ȸ�ϱ�" onClick="popJungsanMaster('<%= ocartoonboxmaster.FOneItem.Fjungsanidx %>');">
		<% else %>
			* ���� : ������ ���� ���
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�κ��̽�IDX</td>
	<td>
		<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>

		<% if (ocartoonboxmaster.FOneItem.Finvoiceidx <> "") then %>
			&nbsp;
			<input type="button" class="button" value="��ȸ�ϱ�" onClick="popinvocesMaster('<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>', '', '');">
		<% elseif idx <> "" and idx <> "0" then %>
			<input type="button" class="button" value="�ۼ��ϱ�" onClick="popinvocesMaster('<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>', '<%= idx %>', '<%= ocartoonboxmaster.FOneItem.Fshopid %>');">
		<% end if %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >����</td>
	<td colspan="3">
		<% ''if (ocartoonboxmaster.FOneItem.Fjungsanidx <> "") then %>
		&nbsp;
		<!--<input type="button" class="button" value=" ��ǰ��� ������� " onClick="PrintDetailItemList(<%'= ocartoonboxmaster.FOneItem.Fjungsanidx %>, '<%'= ocartoonboxmaster.FOneItem.Fshopid %>', '<%'= ocartoonboxmaster.FOneItem.Fshopname %>')">-->
		<input type="checkbox" name="chkimg" >�̹���

		<% if idx <> "" and idx <> "0" then %>
			<input type="button" onclick="downloadOrder('','','<%= ocartoonboxmaster.FOneItem.Fshopid %>','<%= idx %>');" value="��ǰ��� �������" class="button">
		<% end if %>

		<% ''end if %>
		&nbsp;
		<% if (ocartoonboxmaster.FOneItem.Finvoiceidx <> "") then %>
		<input type="button" class="button" value="QS" onClick="PopOpenQS('<%= ocartoonboxmaster.FOneItem.Finvoiceidx%>', '<%=ocartoonboxmaster.FOneItem.Fjungsanidx%>','<%=ocartoonboxmaster.FOneItem.Fidx%>','<%=ocartoonboxmaster.FOneItem.Floginsite%>','<%=ocartoonboxmaster.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
		<input type="button" class="button" value="PI" onClick="PopOpenPI('<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>', '<%=ocartoonboxmaster.FOneItem.Fjungsanidx%>','<%=ocartoonboxmaster.FOneItem.Fidx%>','<%=ocartoonboxmaster.FOneItem.Floginsite%>','<%=ocartoonboxmaster.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
		<input type="button" class="button" value="CI" onClick="PopOpenCI('<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>', '<%=ocartoonboxmaster.FOneItem.Fjungsanidx%>','<%=ocartoonboxmaster.FOneItem.Fidx%>','<%=ocartoonboxmaster.FOneItem.Floginsite%>','<%=ocartoonboxmaster.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
		<input type="button" class="button" value="PL" onClick="PopOpenPL('<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>', '<%=ocartoonboxmaster.FOneItem.Fjungsanidx%>','<%=ocartoonboxmaster.FOneItem.Fidx%>','<%=ocartoonboxmaster.FOneItem.Floginsite%>','<%= idx%>','<%=ocartoonboxmaster.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">

		<input type="button" class="button" value="PL_Item" onClick="PopOpenPLItem('<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>','<%=ocartoonboxmaster.FOneItem.Floginsite%>','<%= idx%>','<%=ocartoonboxmaster.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">

 		&nbsp;
 		<input type="button" class="button" value=" �κ��̽� " onClick="PopOpenInvoice(<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>,'')">
		<input type="button" class="button" value=" �κ��̽� Excel" onClick="PopOpenInvoice(<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>,'Y')">
		&nbsp;
		<input type="button" class="button" value=" ��ŷ����Ʈ " onClick="PopOpenPackingList(<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>,'')">
		<input type="button" class="button" value=" ��ŷ����Ʈ Excel" onClick="PopOpenPackingList(<%= ocartoonboxmaster.FOneItem.Finvoiceidx %>,'Y')">
		&nbsp;

		(* �μ�� ��/�� ������ 1cm ���Ϸ� �����ϼ���)
		<% end if %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >�ۼ���</td>
	<td>
		<% if (ocartoonboxmaster.FOneItem.Freguserid = "") then %>
			<%= session("ssBctid") %>
		<% else %>
			<%= ocartoonboxmaster.FOneItem.Freguserid %>
		<% end if %>
		<input type="hidden" name="reguserid" value="<%= session("ssBctid") %>">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�����</td>
	<td>
		<% if (ocartoonboxmaster.FOneItem.Fregdate = "") then %>
			<%= now %>
		<% else %>
			<%= ocartoonboxmaster.FOneItem.Fregdate %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ�޸�</td>
	<td colspan="3"><textarea class="textarea" name="comment" cols="80" rows="6"><%= ocartoonboxmaster.FOneItem.FComment %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center">
		<input type="button" class="button" value=" �����ϱ� " onClick="SaveMaster(frmMaster)">

		<% if (idx <> 0) then %>
		<input type="button" class="button" value=" ��ü���� " onClick="DelMaster(frmMaster)">
		<% end if %>
	</td>
</tr>
</form>
</table>

<p>

<% if (idx <> 0) then %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="right">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="#FFFFFF">
				<td>
					<input type="button" class="button" value="���ùڽ�����" onClick="DeselectArr()">
					<input type="button" class="button" value=" �ڽ��߰� " onClick="PopBoxSelect('<%= idx %>')">
				</td>
				<td align="right">
					�ѰǼ�:  <%= ocartoonboxdetail.FResultCount %>
				</td>
			</tr>
		</table>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="cksel" onClick="chkAllitem('frmModiPrc', 10)"></td>
    <td width="110">�����̵�</td>
    <td width="70">CartonBOX<br>��ȣ</td>
    <td width="90">CartonBOX<br>N.Weight(KG)</td>
    <td width="90">CartonBOX<br>G.Weight(KG)</td>
	<td width="150">CartonBOX<br>Type</td>
	<td width="90">CBM</td>
    <td width="80">���</td>
	<td width="80">����EMS<br>��ۺ��</td>
    <td width="150">������ȣ</td>
	<td width="80">������</td>
	<td width="70">InnerBOX<br>��ȣ</td>
    <td width="70">InnerBOX<br>����(KG)</td>
    <td width="100">InnerBOX<br>��ǰ����</td>
	<td width="100">InnerBOX<br>���ް�</td>
	<td>���</td>
</tr>
<%
currcartoonboxno = ""
currbaljudate = ""
suminnerboxweight = 0
sumcartoonboxNweight = 0
sumcartoonboxweight = 0
sumemsprice = 0
sumsupplyPrice = 0
sumcartoonboxcbm = 0

j = 0
%>
<% for i=0 to ocartoonboxdetail.FResultCount-1 %>
	<%
	if (ocartoonboxdetail.FItemList(i).Fcartoonboxno <> currcartoonboxno) then
		isnewcartoonbox = true
		currcartoonboxno = ocartoonboxdetail.FItemList(i).Fcartoonboxno
		currbaljudate = ocartoonboxdetail.FItemList(i).Fbaljudate
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
		sumsupplyPrice = sumsupplyPrice + FormatNumber(ocartoonboxdetail.FItemList(i).FsupplyPrice, 2)

		if ocartoonboxdetail.FItemList(i).FcartoonboxType <> "" then
			sumcartoonboxcbm = sumcartoonboxcbm + getcartoonboxtype(ocartoonboxdetail.FItemList(i).FcartoonboxType, 1)
		end if
	end if

	suminnerboxweight = suminnerboxweight + ocartoonboxdetail.FItemList(i).Finnerboxweight

	%>

<form name="frmModiPrc_<%= i %>" method="post" action="cartoonbox_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="orgcartoonboxno"  value="<%= ocartoonboxdetail.FItemList(i).Fcartoonboxno %>">
<input type="hidden" name="mode" value="modifybox">
<input type="hidden" name="masteridx" value="<%= idx %>">
<input type="hidden" name="enclistpageurl" value="<%= enclistpageurl %>">
<input type="hidden" name="detailidx" value="<%= ocartoonboxdetail.FItemList(i).Fidx %>">
<!--
<input type="hidden" name="cartoonboxno" value="<%= ocartoonboxdetail.FItemList(i).Fcartoonboxno %>">
-->
<input type="hidden" name="innerboxno" value="<%= ocartoonboxdetail.FItemList(i).Finnerboxno %>">
<input type="hidden" name="innerboxweight" value="<%= ocartoonboxdetail.FItemList(i).Finnerboxweight %>">
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><%= ocartoonboxdetail.FItemList(i).Fshopid %></td>
	<td>
		<input type="text" name="cartoonboxno" value="<%= ocartoonboxdetail.FItemList(i).Fcartoonboxno %>" size="2" maxlength="3"><!--�����û 2016/09/12-->
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<%= FormatNumber(ocartoonboxdetail.FItemList(i).FcartoonboxNweight, 2) %>
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<input type="text" class="text" name="cartoonboxweight" value="<%= FormatNumber(ocartoonboxdetail.FItemList(i).Fcartoonboxweight, 2) %>" size="6" maxlength="6" style="text-align:right">
		<% else %>
			<input type="hidden" name="cartoonboxweight" value="-1">
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<% drawcartoonboxtype "cartoonboxType", ocartoonboxdetail.FItemList(i).FcartoonboxType, "", "Y", "N", "N" %>
		<% else %>
			<input type="hidden" name="cartoonboxType" value="-1">
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) and ocartoonboxdetail.FItemList(i).FcartoonboxType <> "" then %>
			<%= getcartoonboxtype(ocartoonboxdetail.FItemList(i).FcartoonboxType, 1) %>
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<%= FormatNumber(ocartoonboxdetail.FItemList(i).FsupplyPrice, 2) %>
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<%= FormatNumber(ocartoonboxdetail.FItemList(i).Femsprice, 0) %>
		<% end if %>
	</td>
	<td>
		<% if (isnewcartoonbox = true) then %>
			<input type="text" class="text" name="cartonboxsongjangno" value="<%= ocartoonboxdetail.FItemList(i).Fcartonboxsongjangno %>" size="15" maxlength="15" style="text-align:right">
		<% else %>
			<input type="hidden" name="cartonboxsongjangno" value="-1">
		<% end if %>
	</td>
	<td><%= ocartoonboxdetail.FItemList(i).Fbaljudate %></td>
	<td>
		<%= ocartoonboxdetail.FItemList(i).Finnerboxno %>
	</td>
	<td>
		<%= FormatNumber(ocartoonboxdetail.FItemList(i).Finnerboxweight, 2) %>
	</td>
	<td>
		<input type="button" class="button" value=" ��ǰ���� " onClick="PopBoxItemList('<%= ocartoonboxdetail.FItemList(i).Fshopid %>', '<%= Left(ocartoonboxdetail.FItemList(i).Fbaljudate, 4) %>', '<%= Right(Left(ocartoonboxdetail.FItemList(i).Fbaljudate, 7), 2) %>', '<%= Right(ocartoonboxdetail.FItemList(i).Fbaljudate, 2) %>', <%= ocartoonboxdetail.FItemList(i).Finnerboxno %>)">
	</td>
	<td>
		<%= FormatNumber(ocartoonboxdetail.FItemList(i).FinnerSupplyPrice, 0) %>
	</td>
	<td>
		<!--
		<% if (isnewcartoonbox = true) then %>
			&nbsp;
			<input type="button" class="button" value=" �����ڵ���� " onClick="CalcCartoonboxWeight(frmModiPrc_<%= i %>)">
		<% end if %>
		-->
	</td>
</tr>
</form>
<% next %>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan=3></td>
    <td><%= FormatNumber(sumcartoonboxNweight, 2) %></td>
    <td><%= FormatNumber(sumcartoonboxweight, 2) %></td>
	<td></td>
	<td><%= sumcartoonboxcbm %></td>
    <td><%= FormatNumber(sumsupplyPrice, 2) %></td>
	<td><%= FormatNumber(sumemsprice, 0) %></td>
	<td colspan=3></td>
	<td><%= FormatNumber(suminnerboxweight, 2) %></td>
    <td colspan=3></td>
</tr>

<tr bgcolor="#FFFFFF">
	<td colspan="16" align=center height=30>
	<% if ocartoonboxmaster.FOneItem.Fworkstate="9" then %>
		<b>�԰� �Ϸ�� ������ ���� �Ͻ� �� �����ϴ�.</b>
	<% elseif (ocartoonboxmaster.FOneItem.Fworkstate>"6") then %>
		<b>��� �Ϸ�� ������ ���� �Ͻ� �� �����ϴ�.</b>
	<% else %>
		<input type="button" class="button" value=" ��ü���� " onclick="SaveDetailArr()">
		<!--
		&nbsp;
		<input type="button" class="button" value=" ��ü���� " onclick="DelMaster(frmMaster)">
		-->
	<% end if %>
	</td>
</tr>
<form name="frmadd" method=post action="cartoonbox_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="shopid" value="<%= ocartoonboxmaster.FOneItem.Fshopid %>">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="enclistpageurl" value="<%= enclistpageurl %>">
	<input type="hidden" name="detailidxarr" value="">
	<input type="hidden" name="cartoonboxnoarr" value="">
	<input type="hidden" name="cartoonboxweightarr" value="">
	<input type="hidden" name="cartoonboxTypearr" value="">
	<input type="hidden" name="cartonboxsongjangnoarr" value="">
	<input type="hidden" name="innerboxnoarr" value="">
	<input type="hidden" name="innerboxweightarr" value="">
	<input type="hidden" name="baljudatearr" value="">
	<input type="hidden" name="shopidarr" value="">
</form>
</table>

<p>

* �ϳ��� �ֹ����� �ϳ��� ��� ��� ��ϵǾ�� �մϴ�.<br>
* ������ �Ϻ���� �䱸�ϴ� ��쿡�� �ϳ��� ���ǿ� ��� ����ϰ� ����� �Ϻ���� �ؾ� �մϴ�.

<% else %>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="right">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr bgcolor="#FFFFFF" height=20>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
						<font color="red"><strong>�������ڽ�</strong></font>
					</td>
					<td align="right">
						�ѰǼ�:  <%= oinnerboxlist.FResultCount %>
					</td>
				</tr>
			</table>
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><!--<input type="checkbox" name="cksel" onClick="chkAllitem('frmSelectPrc', 12)">--></td>
		<td width="110">�����̵�</td>
		<td width="250">���̸�</td>
		<td width="120">������</td>
		<td width="80">Inner<br>�ڽ���ȣ</td>
		<td width="80">Inner<br>����(KG)</td>
		<td width="80">Carton<br>�ڽ���ȣ</td>
		<td width="80">�����</td>
		<td>���</td>
	</tr>
	<% for i=0 to oinnerboxlist.FResultCount-1 %>
	<form name="frmSelectPrc_<%= i %>" method="post" action="cartoonbox_process.asp">
		<input type="hidden" name="detailidx" value="<%= oinnerboxlist.FItemList(i).Fidx %>">
		<input type="hidden" name="baljudate" value="<%= oinnerboxlist.FItemList(i).Fbaljudate %>">
		<input type="hidden" name="shopid" value="<%= oinnerboxlist.FItemList(i).Fshopid %>">
		<input type="hidden" name="innerboxweight" value="<%= oinnerboxlist.FItemList(i).Finnerboxweight %>">
		<input type="hidden" name="cartoonboxno" value="<%= oinnerboxlist.FItemList(i).Fcartoonboxno %>">
		<tr align="center" bgcolor="#FFFFFF">
			<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this); jsDisableOtherShop();"></td>
			<td><%= oinnerboxlist.FItemList(i).Fshopid %></td>
			<td><%= oinnerboxlist.FItemList(i).Fshopname %></td>
			<td><%= oinnerboxlist.FItemList(i).Fbaljudate %></td>
			<td>
				<%= oinnerboxlist.FItemList(i).Finnerboxno %>
				<input type="hidden" name="innerboxno" value="<%= oinnerboxlist.FItemList(i).Finnerboxno %>">
			</td>
			<td>
				<%
				if (oinnerboxlist.FItemList(i).Finnerboxweight <> "") then
					oinnerboxlist.FItemList(i).Finnerboxweight = FormatNumber(oinnerboxlist.FItemList(i).Finnerboxweight, 2)
				end if
				%>
				<%= oinnerboxlist.FItemList(i).Finnerboxweight %>
			</td>
			<td>
				<%= oinnerboxlist.FItemList(i).Fcartoonboxno %>
			</td>
			<td>
				<%= oinnerboxlist.FItemList(i).Fbeasongdate %>
				<input type="hidden" name="cartoonboxweight" value="0">
			</td>
			<td>
				<input type="button" class="button" value=" ��ǰ���� " onClick="PopBoxItemList('<%= oinnerboxlist.FItemList(i).Fshopid %>', '<%= Left(oinnerboxlist.FItemList(i).Fbaljudate, 4) %>', '<%= Right(Left(oinnerboxlist.FItemList(i).Fbaljudate, 7), 2) %>', '<%= Right(oinnerboxlist.FItemList(i).Fbaljudate, 2) %>', <%= oinnerboxlist.FItemList(i).Finnerboxno %>)">
			</td>
		</tr>
	</form>
	<% next %>
</table>

<% end if %>

<%
set ocartoonboxmaster = Nothing
set ocartoonboxdetail = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
