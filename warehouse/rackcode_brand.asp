<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �귣�巺�ڵ����
' Hieditor : �̻� ����
'			 2020.01.09 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/rackipgocls.asp"-->
<%
dim isusing, research, page
dim rackcode2, rackcode, makerid
dim maeipdiv
dim searchtype, fromrackcode2, torackcode2, purchasetype
dim warehouseCd

page        = request("page")
isusing     = request("isusing")
research    = request("research")
rackcode2   = requestCheckvar(request("rackcode2"),4)
maeipdiv  	= requestCheckvar(request("maeipdiv"),1)

searchtype  	= requestCheckvar(request("searchtype"),1)
fromrackcode2  	= requestCheckvar(request("fromrackcode2"),4)
torackcode2  	= requestCheckvar(request("torackcode2"),4)
purchasetype  	= requestCheckvar(request("purchasetype"),2)
warehouseCd  	= requestCheckvar(request("warehouseCd"),3)

makerid     = request("makerid")

if research="" and isusing="" then isusing="Y"
if page="" then page=1
'if searchtype="" then searchtype = "F"

dim orackcode_brand
set orackcode_brand = new CRackIpgo
orackcode_brand.FCurrpage = page
orackcode_brand.FPageSize = 30
orackcode_brand.FRectMakerid = makerid
orackcode_brand.FRectRackCode = rackcode2
orackcode_brand.FRectIsUsingYN = isusing
orackcode_brand.FRectMaeipDiv = maeipdiv
orackcode_brand.FRectSearchType = searchtype
orackcode_brand.FRectFromRackcode2 = fromrackcode2
orackcode_brand.FRectToRackcode2 = torackcode2
orackcode_brand.FRectPurchaseType = purchasetype
orackcode_brand.FRectWarehouseCd = warehouseCd

orackcode_brand.GetRackBrandList

dim i
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/ttpbarcode.js"></script>
<script type="text/javascript" src="/js/barcode.js"></script>
<script type="text/javascript" src="/js/DOSHIBAbarcode.js"></script>
<script type='text/javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popBrandRackCodeEdit(imakerid){
    var popwin = window.open('pop_BrandRackCodeEdit.asp?makerid=' + imakerid,'popBrandRackCodeEdit','width=500,height=200,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// ���ⷢ�ڵ��ε������		// 2020.01.09 �ѿ�� ����
function IndexSudongrackBarcodePrint(){
	var popwin = window.open('/common/barcode/sudongrackindexprint.asp?menupos=<%=menupos%>','IndexSudongrackBarcodePrint','width=1024,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function toggleChecked(status) {
    $('[name="check"]').each(function () {
        $(this).prop("checked", status);
    });
}

function IndexrackBarcodePrint() {
	var arrbarcode = new Array();
    var isforeignprint; var domainname; var showdomainyn; var showdomainyn; var ttptype;
    var barcodetype;
    var shopbrandyn; var currencychar; var showpriceyn;

	isforeignprint = "N";

	shopbrandyn		= "N";
	ttptype			= "TTP-243_80x50";
	barcodetype		= "";
	var paperwidth = "80";
	var paperheight = "50";
	var papermargin = "3";
	var heightoffset = 0;
	showpriceyn = "N";

	currencychar = "��";
	domainname		= "www.10x10.co.kr";
	showdomainyn	= "Y";

	var fontName = "10X10";
    var rackcode = "";
	var itemno = document.frmArr.itemno.value;
	var MAX_MESSAGE_LENGTH = 8;

    if ($('input[name="check"]:checked').length == 0) {
        alert('���� �������� �����ϴ�.');
        return;
    }
    if (itemno=="" || itemno==0){
        alert("����Ͻ� ������ �Է����ּ���.");
        frmArr.itemno.focus();
        return;
    }

	$('input[name="check"]:checkbox:checked').each(function () {
		rackcode = $(this).attr('rackcode');

		if (rackcode!=""){
			if (rackcode.replace("\r", "").length > MAX_MESSAGE_LENGTH) {
				alert("\n========== ���� ==========\n\n��¸޽����� " + MAX_MESSAGE_LENGTH + "���ڸ� �̻��� ���� �� �����ϴ�. ");
				return;
			}

			var v = new BarcodeDataClass_udong(rackcode,itemno,fontName);
			arrbarcode.push(v);
		}
	});

	//TEC B-FV4		//2020.01.09 �ѿ�� ����
	if (TEC_DO3.IsDriver == 1){
		if (confirm("�����Ͻ� ���ڵ� �ε����� ����մϴ�.\n\nTEC B-FV4 �� ����Ͻðڽ��ϱ�?") == true) {
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_PAPERWIDTH = 800;
			TOSHIBA_PAPERHEIGHT = 500;
			TOSHIBA_PAPERMARGIN = 3;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = barcodetype;

			printTOSHIBAMultiRackIndexLabel(arrbarcode);
		}

	// /js/barcode.js ����
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) == true) {
		if (confirm("�����Ͻ� ���ڵ� �ε����� ����մϴ�.\n\nTTP-243 �� ����Ͻðڽ��ϱ�?") == true) {
			printTTPMultiRackIndexLabel(arrbarcode);
		}

	}else {
	    alert("TTP-243(��)�� TEC B-FV4 ����̹��� ��ġ�� �ּ���");
	}

	return;
}

function jsSetChecked(index) {
    var check, warehouseCd;

    check = document.getElementById('check' + index);
    check.checked = true;
    AnCheckClick(check);
}

function jsSetWarehouseCd() {

    var frmAct, mode, makeridArr, warehouseCdArr;
    var makerid, check, warehouseCd;

    frmAct = document.frmAct;
    makeridArr = "";
    warehouseCdArr = "";

    if ($('input[name="check"]:checked').length == 0) {
        alert('���� �������� �����ϴ�.');
        return;
    }

    for (var i = 0; ; i++) {
        check = document.getElementById('check' + i);
        makerid = document.getElementById('makerid' + i);
        warehouseCd = document.getElementById('warehouseCd' + i);

        if (check == undefined) { break; }
        if (check.checked != true) { continue; }

        if (warehouseCd.value == '') {
            alert('�����Ӽ��� �����ϼ���.');
            warehouseCd.focus();
            return;
        }

        makeridArr = makeridArr + "," + makerid.value;
        warehouseCdArr = warehouseCdArr + "," + warehouseCd.value;
    }

    if (confirm('�����Ͻðڽ��ϱ�?')) {
        frmAct.mode.value = "setwarehousecd";
        frmAct.makeridArr.value = makeridArr;
        frmAct.warehouseCdArr.value = warehouseCdArr;
        frmAct.submit();
    }
}

$(document).ready(function () {
    var checkAllBox = $("#ckall");

    checkAllBox.click(function () {
        var status = checkAllBox.prop('checked');
        toggleChecked(status);
    });
});

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		���ڵ�(4�ڸ�) :
		<input type="radio" name="searchtype" value="F" <% if (searchtype = "F") then %>checked<% end if %> >
		<input type="text" name=rackcode2 value="<%= rackcode2 %>" maxlength="4" size="4" class="text">
		&nbsp;
		<input type="radio" name="searchtype" value="R" <% if (searchtype = "R") then %>checked<% end if %> >
		<input type="text" name=fromrackcode2 value="<%= fromrackcode2 %>" maxlength="4" size="4" class="text">
		~
		<input type="text" name=torackcode2 value="<%= torackcode2 %>" maxlength="4" size="4" class="text">
		&nbsp;
		��� : <% drawSelectBoxUsingYN "isusing", isusing %>
		&nbsp;
		�¶��α⺻���� : <% DrawBrandMWUCombo "maeipdiv", maeipdiv %>
		&nbsp;
		�������� : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
        &nbsp;
		�����Ӽ� :
        <select class="select" name="warehouseCd">
            <option></option>
            <option value="AGV" <%= CHKIIF(warehouseCd = "AGV", "selected", "") %>>AGV</option>
            <option value="BLK" <%= CHKIIF(warehouseCd = "BLK", "selected", "") %>>BLK</option>
            <option value="NUL" <%= CHKIIF(warehouseCd = "NUL", "selected", "") %>>������</option>
        </select>
	</td>
</tr>
</table>
</form>

<br>

<!-- �׼� ���� -->
<form name="frmArr" action="" method="get" style="margin:0px;">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="���ڵ庯��" onClick="popBrandRackCodeEdit('');">
        &nbsp;
        <input type="button" class="button" value="�����Ӽ�����" onClick="jsSetWarehouseCd();">
	</td>
	<td align="right">
		<input type="text" name="itemno" value="1" size=3 maxlength=5>
		<input type="button" class="button" value="���÷��ڵ��ε������" onClick="IndexrackBarcodePrint();">
		&nbsp;
		<input type="button" class="button" value="���ⷢ�ڵ��ε������" onClick="IndexSudongrackBarcodePrint();">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" id="ckall" onclick="totalCheck()"></td>
	<td width="90">���ڵ�</td>
	<td width="180">�귣��ID</td>
    <td width="60">�����Ӽ�</td>
    <td>��Ʈ��Ʈ��</td>
	<td width="50">���</td>
	<!--
	<td width="70">���(����)</td>
	-->
	<td width="50">���ڽ�<br>����</td>
	<td width="120">��������ǰ</td>
	<td width="120">��ǰ����ǰ</td>
	<td width="150">�Ͻ�ǰ��������ǰ</td>
	<td width="120">�Ͻ�ǰ����ǰ</td>
</tr>
<% for i=0 to orackcode_brand.FResultCount - 1 %>
<tr <%= chkIIF( orackcode_brand.FItemList(i).FBrandUsing="Y","bgcolor='#FFFFFF'","bgcolor='#CCCCCC'") %>>
	<td align="center"><input type="checkbox" id="check<%= i %>" name="check" rackcode="<%= orackcode_brand.FItemList(i).Frackcode %>" onClick="AnCheckClick(this);"></td>
    <input type="hidden" id="makerid<%= i %>" name="makerid" value="<%= orackcode_brand.FItemList(i).FMakerid %>">
	<td align="center"><%= orackcode_brand.FItemList(i).Frackcode %></td>
	<td><a href="javascript:popBrandRackCodeEdit('<%= orackcode_brand.FItemList(i).FMakerid %>');"><%= orackcode_brand.FItemList(i).FMakerid %></a></td>
    <td align="center">
        <select class="select" id="warehouseCd<%= i %>" name="warehouseCd" onchange="jsSetChecked(<%= i %>)">
            <option></option>
            <option value="AGV" <%= CHKIIF(orackcode_brand.FItemList(i).FwarehouseCd = "AGV", "selected", "") %>>AGV</option>
            <option value="BLK" <%= CHKIIF(orackcode_brand.FItemList(i).FwarehouseCd = "BLK", "selected", "") %>>BLK</option>
        </select>
    </td>
    <td><%= orackcode_brand.FItemList(i).Fmakername %></td>
	<td align="center"><%= ChkIIF(orackcode_brand.FItemList(i).FBrandUsing="Y","O","X") %></td>
	<!--
	<td align="center"><%= ChkIIF(orackcode_brand.FItemList(i).FBrandUsingExt="Y","O","X") %></td>
	-->
	<td align="right"><%= orackcode_brand.FItemList(i).Frackboxno %></td>
	<td align="center">
		<input type="button" class="button" value="��������ǰ" onclick="javascript:window.open('/admin/stock/brandcurrentstock.asp?menupos=708&research=on&page=&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&onoffgubun=on&mwdiv=MW&returnitemgubun=rackdisp');">
	</td>
	<td align="center">
		<!-- <input type="button" class="button" value="��ǰ����ǰ" onclick="javascript:window.open('/admin/stock/return_item.asp?menupos=983&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&realstocknotzero=on');"> -->
		<input type="button" class="button" value="��ǰ����ǰ" onclick="javascript:window.open('/admin/stock/brandcurrentstock.asp?menupos=708&research=on&page=&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&onoffgubun=on&mwdiv=MW&returnitemgubun=reton');">
	</td>
	<td align="center">
		<input type="button" class="button" value="�Ͻ�ǰ��������ǰ" onclick="javascript:window.open('/admin/stock/brandcurrentstock.asp?menupos=708&research=on&page=&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&onoffgubun=on&sellyn=S&usingyn=&danjongyn=YM&mwdiv=MW');">
	</td>

	<td align="center">
		<!-- �Ͻ�ǰ��/�������� -->
		<input type="button" class="button" value="�Ͻ�ǰ����ǰ" onclick="javascript:window.open('/admin/shopmaster/danjong_set.asp?menupos=1053&research=on&page=1&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&mwdiv=MW');">
	</td>
</tr>
<% next %>

<tr bgcolor="#FFFFFF">
	<td colspan="11" align="center">
		<% if orackcode_brand.HasPreScroll then %>
		<a href="javascript:NextPage('<%= orackcode_brand.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + orackcode_brand.StarScrollPage to orackcode_brand.FScrollCount + orackcode_brand.StarScrollPage - 1 %>
			<% if i>orackcode_brand.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if orackcode_brand.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
</form>

<%
set orackcode_brand = Nothing
%>

<form name="frmAct" method="post" action="brandRackCode_process.asp">
<input type="hidden" name="mode" />
<input type="hidden" name="makeridArr" />
<input type="hidden" name="warehouseCdArr" />
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
