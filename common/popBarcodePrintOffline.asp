<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��ǰ ���ڵ� ���
' Hieditor : 2010.10.26 ������ ����
'			 2011.02.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchulproductcls.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchullocationcls.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchulbarcodecls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim i,page,research, shopid ,useforeigndata, currencyunit , ipchul, currencyChar
dim prdcode, prdname, itemid, generalbarcode, makerid ,olocation , printername
dim printpriceyn, isforeignprint ,shopitemname ,cdl, cdm, cds ,oproduct , listgubun
dim currentstockexist, realstockonemore, shopitemnameinserted, displayrealstockno , makeriddispyn
	listgubun 		= requestCheckVar(request("listgubun"), 32)
	prdcode 		= requestCheckVar(request("prdcode"),32)
	prdname 		= requestCheckVar(request("prdname"),32)
	itemid 			= requestCheckVar(request("itemid"),255)
	generalbarcode 	= requestCheckVar(request("generalbarcode"),32)
	makerid = requestCheckVar(request("makerid"),32)
	shopid 		= requestCheckVar(request("shopid"),32)
	shopitemname 	= RequestCheckVar(request("shopitemname"),32)
	cdl         	= RequestCheckVar(request("cdl"),3)
	cdm         	= RequestCheckVar(request("cdm"),3)
	cds         	= RequestCheckVar(request("cds"),3)
	page 			= requestCheckVar(request("page"),32)
	printpriceyn 	= requestCheckVar(request("printpriceyn"),32)
	isforeignprint 	= requestCheckVar(request("isforeignprint"),32)
	currentstockexist 		= requestCheckVar(request("currentstockexist"),32)
	realstockonemore 		= requestCheckVar(request("realstockonemore"),32)
	shopitemnameinserted 	= requestCheckVar(request("shopitemnameinserted"),32)
	displayrealstockno 		= requestCheckVar(request("displayrealstockno"),32)
	printername = requestCheckVar(request("printername"),32)
	ipchul 		= requestCheckVar(request("ipchul"),32)
	research 		= requestCheckVar(request("research"),32)
	makeriddispyn 			= requestCheckVar(request("makeriddispyn"),32)

response.write "�ŸŴ� ������, ������� �Ŵ� �Դϴ�. �������� ���� �ϼ���."
response.end

'/�����ϰ�� ���� ���常 ��밡��
if (C_IS_SHOP) then
	'/���α��� ���� �̸�
	'if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	'end if
else
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then

		else

		end if
	end if
end if

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if
if displayrealstockno = "" and research <> "on" then
	displayrealstockno = "Y"
end if

if printername = "" then printername = "TEC_B-FV4"
if listgubun = "" then listgubun = "ITEM"
if makeriddispyn = "" then makeriddispyn = "Y"

if page = "" then page = 1
useforeigndata = "N"
currencyunit = "WON"
currencyChar = "��"

set oproduct = new CProduct
	oproduct.FCurrpage = page
	oproduct.FPageSize = 100
	oproduct.FRectLocationId = shopid				'�̵�ó
	oproduct.FRectLocationIdMaker = makerid
	oproduct.FRectPrdCode = prdcode
	oproduct.FRectItemID = itemid
	oproduct.FRectPrdName = html2db(prdname)
	oproduct.FRectGeneralBarcode = generalbarcode
	''oproduct.FRectUseYN = "Y"                         ''��뱸�� ������� ��ü ǥ��
	oproduct.FRectCDL = cdl
	oproduct.FRectCDM = cdm
	oproduct.FRectCDS = cds
	oproduct.FRectShopItemName = html2db(shopitemname)
	oproduct.FRectCurrentStockExist = currentstockexist
	oproduct.FRectRealStockOneMore = realstockonemore
	oproduct.FRectShopItemNameInserted = shopitemnameinserted
	oproduct.frectipchul = ipchul

	if shopid <> "" then
		if listgubun = "ITEM" then
			oproduct.GetProductListOffline()
		elseif listgubun = "JUMUN" then

			if ipchul <> "" then
				oproduct.GetipchulListOffline()
			end if
		end if
	else
	    response.write "<script language='javascript'>"
	    response.write "    alert('������ ������ �ּ���');"
	    response.write "</script>"
	end if

set olocation = new CLocation
	olocation.FRectlocationid = shopid

	if (shopid <> "") then
		olocation.GetOneLocation

		useforeigndata = olocation.FOneItem.Fuseforeigndata
		if (isforeignprint = "") then
			isforeignprint = useforeigndata
		end if
		currencyunit = olocation.FOneItem.Fcurrencyunit
		currencyChar = olocation.FOneItem.FcurrencyChar
	end if

Function AddSpace(byval str)
	if ((str = "") or (IsNull(str))) then
		AddSpace = "&nbsp;"
	else
		AddSpace = str
	end if
End Function
%>

<script type="text/javascript" src="/js/ttpbarcode.js"></script>
<script type="text/javascript" src="/js/barcode.js"></script>
<script type="text/javascript" src="/js/DOSHIBAbarcode.js"></script>
<script type="text/javascript">

function reg(page){
//	if(frm.itemid.value!=""){
//		if (!IsDouble(frm.itemid.value)){
//			alert("��ǰ�ڵ�� ���ڸ� �����մϴ�.");
//			frm.itemid.focus();
//			return;
//		}
//	}

	frm.page.value=page;
	frm.submit();
}

function SearchByItemId(frm) {

	frm.prdcode.value = "";
	frm.submit();
}

function SearchByPrdcode(frm) {

	frm.itemid.value = "";
	frm.submit();
}

function ipchulview(v,onload){
	if (onload =="ONLOAD"){
		if (frm.ipchul.disabled){
			frm.listgubun[1].disabled = true;
		}else{
			if (v == "ITEM"){
				frm.ipchul.disabled = true;
			} else if (v == "JUMUN") {
				frm.ipchul.disabled = false;
			}
		}
	}else{
		if (v == "ITEM"){
			frm.ipchul.disabled = true;
		} else if (v == "JUMUN") {
			frm.ipchul.disabled = false;
		}
	}
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function jsSetSelectBoxColor() {
	var frm = document.frm;

	frm.printpriceyn.style.background = "";
	frm.makeriddispyn.style.background = "";
	frm.isforeignprint.style.background = "";

	if (frm.printpriceyn.value == "N") {
		frm.printpriceyn.style.background = "orange";
	}

	if (frm.makeriddispyn.value == "N") {
		frm.makeriddispyn.style.background = "orange";
	}

	if (frm.isforeignprint.value == "Y") {
		frm.isforeignprint.style.background = "orange";
	}
}

//�ε��� ���
function IndexBarcodePrint() {
	var arr = new Array();
	var isforeignprint; var printpriceval; var domainname; var showdomainyn; var ttptype;
	var prdname; var itemoptionname; var customerprice; var barcodetype;
	var skipnotinserted; var printno; var shopbrandyn; var currencychar; var showpriceyn;

	isforeignprint = document.frm.isforeignprint.value;
	skipnotinserted = false;

	shopbrandyn		= "Y";
	ttptype			= "TTP-243_80x50";
	barcodetype		= "2";								// T or G or 2(�ٹ����� ���ڵ� or ������ڵ� or ���ٹ��ڵ�_������ڵ�)
	var paperwidth = frm.paperwidth.value;
	var paperheight = frm.paperheight.value;
	var papermargin = frm.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frm.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "��";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	domainname		= "www.10x10.co.kr";
	showdomainyn	= frm.makeriddispyn.value;

	for (var i = 0; ; i++) {
		chk = document.getElementById("chk_" + i);

		if (chk == undefined) {
			break;
		}

		if (chk.checked == true) {
			if (isforeignprint == "N") {
				itemname = document.getElementById("itemname_" + i).value.trim();
				itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();
				customerprice = document.getElementById("customerprice_" + i).value.trim();
			} else {
				itemname = document.getElementById("itemname_foreign_" + i).value.trim();
				itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
				customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();
			}

			itembarcode = document.getElementById("itembarcode_" + i).value.trim();
			publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
			itembarcode = itembarcode + "_" + publicbarcode;

			makerid = document.getElementById("makerid_" + i).value.trim();
			//printno = document.getElementById("printno_" + i).value.trim();
			printno = 1;

			if (printno*1 != 0) {
				var v = new TTPBarcodeDataClass(itembarcode, makerid, itemname, itemoptionname, customerprice, printno);
				arr.push(v);
			}
		}
	}

	//TEC B-FV4		//2016.11.24 �ѿ�� ����
	if (frm.printername.value=='TEC_B-FV4'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���!!"); return; }
		if (confirm("���� ��ǰ�� �ε����� ����մϴ�.\n\nTEC B-FV4 �� ����Ͻðڽ��ϱ�?") == true) {
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_PAPERWIDTH = 800;
			TOSHIBA_PAPERHEIGHT = 500;
			TOSHIBA_PAPERMARGIN = 3;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = barcodetype;

			printTOSHIBAMultiItemLabel(arr);
		}

	// /js/barcode.js ����
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) == true) {
		if (confirm("���� ��ǰ�� �ε����� ����մϴ�.\n\nTTP-243 �� ����Ͻðڽ��ϱ�?") == true) {
			printTTPMultiItemLabel(arr);
		}

	}else {
	    alert("TTP-243(��)�� TEC B-FV4 ����̹��� ��ġ�� �ּ���");
	}
	return;
}

//��ǰ�ڵ� ���
function BarcodePrint(btype) {
	var arrbarcode = new Array();
	var chk, itembarcode, makerid, itemname, itemoptionname, customerprice, printno; var ttptype;
	var found = false;
	var isforeignprint; var currencychar; var domainname; var showdomainyn; var showpriceyn; var shopbrandyn; var publicbarcode;

	isforeignprint = frm.isforeignprint.value;

	shopbrandyn		= frm.makeriddispyn.value;
	ttptype			= "TTP-243_45x22";
	var paperwidth = frm.paperwidth.value;
	var paperheight = frm.paperheight.value;
	var papermargin = frm.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frm.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "��";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	domainname		= "www.10x10.co.kr";
	showdomainyn	= "Y";

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("chk_" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		if (isforeignprint == "N") {
			itemname = document.getElementById("itemname_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();
			customerprice = document.getElementById("customerprice_" + i).value.trim();
		} else {
			itemname = document.getElementById("itemname_foreign_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
			customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();
		}

		itembarcode = document.getElementById("itembarcode_" + i).value.trim();
		publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
		itembarcode = itembarcode + "_" + publicbarcode;
		makerid = document.getElementById("makerid_" + i).value.trim();
		printno = document.getElementById("printno_" + i).value.trim();

		var v = new TTPBarcodeDataClass(itembarcode, makerid, itemname, itemoptionname, customerprice, printno);
		arrbarcode.push(v);
	}

	if (found == false) {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	//TEC B-FV4		//2016.11.24 �ѿ�� ����
	if (frm.printername.value=='TEC_B-FV4'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���!!"); return; }
		if (confirm("���� ��ǰ�� ���ڵ带 ����մϴ�.\n\nTEC B-FV4 �� ����Ͻðڽ��ϱ�?") == true) {
			TOSHIBA_PAPERWIDTH = paperwidth;
			TOSHIBA_PAPERHEIGHT = paperheight;
			TOSHIBA_PAPERMARGIN = papermargin;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = btype;

			printTOSHIBAMultiBarcode(arrbarcode);
		}

	// /js/barcode.js ����
	}else if (initTTPprinter(ttptype, btype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, 0) == true) {
		if (confirm("���� ��ǰ�� ���ڵ带 ����մϴ�.\n\nTTP-243 �� ����Ͻðڽ��ϱ�?") == true) {
			printTTPMultiBarcode(arrbarcode);
		}

	}else {
	    alert("TTP-243(��)�� TEC B-FV4 ����̹��� ��ġ�� �ּ���");
	}
	return;
}

String.prototype.trim = function() {
    return this.replace(/^\s+|\s+$/g, "");
};

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<%
		'����/������
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ���� : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			<% end if %>
		<% else %>
			<% if (C_IS_Maker_Upche) then %>
				* ���� : <% drawBoxDirectIpchulOffShopByMakerchfg "shopid", shopid, makerid, " onchange='reg("""");'", " 'B011','B012','B013'" %>
			<% else %>
				* ���� : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			<% end if %>
		<% end if %>
		&nbsp;&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>

	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<% if (C_IS_Maker_Upche) then %>
			* �귣�� : <%= makerid %>
			<input type="hidden" name="makerid" value="<%= makerid %>">
			&nbsp;&nbsp;
		<% else %>
			* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;&nbsp;
		<% end if %>

		* ��ǰ�ڵ� : 
		<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea><!--onKeyPress="if (event.keyCode == 13) 	reg('');"-->
		&nbsp;&nbsp;
		* ��ǰ�� : <input type="text" class="text" name="prdname" value="<%= prdname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;&nbsp;
		* �����ڵ� :
		<input type="text" class="text" name="prdcode" value="<%= prdcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">

	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<% if not(C_IS_Maker_Upche) then %>
			* ���庰��ǰ�� :
			<input type="text" class="text" name="shopitemname" value="<%= shopitemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
			&nbsp;&nbsp;
			<input type="checkbox" class="checkbox" name="currentstockexist" value="Y" onClick="reg('');" <% if (currentstockexist = "Y") then %>checked<% end if %>> �԰��������ǰ
			&nbsp;&nbsp;
			<input type="checkbox" class="checkbox" name="realstockonemore" value="Y" onClick="reg('');" <% if (realstockonemore = "Y") then %>checked<% end if %>> ��������ǰ
			&nbsp;&nbsp;
			<input type="checkbox" class="checkbox" name="shopitemnameinserted" value="Y" onClick="reg('');" <% if (shopitemnameinserted = "Y") then %>checked<% end if %>> ���庰��ǰ���ϻ�ǰ��
			&nbsp;&nbsp;
		<% end if %>

		<% if listgubun = "ITEM" then %>
			* ������ڵ� :
			<input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
			&nbsp;&nbsp;

			<% if not(C_IS_Maker_Upche) then %>
				<input type="checkbox" class="checkbox" name="displayrealstockno" value="Y" onClick="reg('');" <% if (displayrealstockno = "Y") then %>checked<% end if %>>���� ��������
			<% end if %>
		<% elseif listgubun = "JUMUN" then %>
			<input type="checkbox" class="checkbox" name="displayrealstockno" value="Y" onClick="reg('');" <% if (displayrealstockno = "Y") then %>checked<% end if %>>
			����������� ��������
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�� ������ ���� :
		<select name="printername" onchange="reg('');">
			<option value="TTP-243_45x22" <% if printername = "TTP-243_45x22" then response.write " selected" %>>TTP-243 (�԰�45x22)</option>
			<option value="TTP-243_80x50" <% if printername = "TTP-243_80x50" then response.write " selected" %>>TTP-243 (�԰�80x50)</option>
			<option value="TEC_B-FV4" <% if printername = "TEC_B-FV4" then response.write " selected" %>>TEC B-FV4</option>
		</select>
		&nbsp;&nbsp;
		* ǥ�û�ǰ�� :
		<select name="isforeignprint" onchange="reg('');">
			<option value="N" <% if (isforeignprint = "N") then %>selected<% end if %>>������ǰ��</option>
			<option value="Y" <% if (isforeignprint = "Y") then %>selected<% end if %>>������ǰ��</option>
		</select>
		&nbsp;&nbsp;
		* �ݾ�ǥ�ÿ��� :
		<select name="printpriceyn" onchange="reg('');">
			<option value="Y" <% if (printpriceyn = "Y") then %>selected<% end if %>>�ݾ�ǥ��</option>
			<option value="N" <% if (printpriceyn = "N") then %>selected<% end if %>>�ݾ�ǥ�þ���</option>
		</select>
		&nbsp;&nbsp;
		* �귣��ǥ�� :
		<select name="makeriddispyn" onchange="reg('');">
			<option value="Y" <% if (makeriddispyn = "Y") then %>selected<% end if %>>�귣��ǥ��</option>
			<option value="N" <% if (makeriddispyn = "N") then %>selected<% end if %>>�귣��ǥ�þ���</option>
		</select>
        <br>
        * ���ڵ� ���� �԰� -
		<% if printername = "TTP-243_45x22" then %>
			����:<input type="text" name="paperwidth" value="45" size="4" maxlength=9>
			����:<input type="text" name="paperheight" value="22" size="4" maxlength=9>
		<% elseif printername = "TTP-243_80x50" then %>
			����:<input type="text" name="paperwidth" value="80" size="4" maxlength=9>
			����:<input type="text" name="paperheight" value="50" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4" then %>
			����:<input type="text" name="paperwidth" value="450" size="4" maxlength=9>
			����:<input type="text" name="paperheight" value="220" size="4" maxlength=9>
		<% end if %>

		����:<input type="text" name="papermargin" value="3" size="4" maxlength=9>

		<script language="javascript">
			jsSetSelectBoxColor();
		</script>
	</td>
</tr>
</table>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<a href="http://imgstatic.10x10.co.kr/offshop/sample/print/���ù�_TEC B-FV4_�������ڵ�_���ù�.docx" target="_blank">TEC B-FV4 ���ð� �ٿ�ε�</a>

		<% if shopid <> "" then %>
			<br><font color="red">����Ʈ���� :</font>
			<input type="radio" name="listgubun" value="ITEM" onClick="ipchulview(this.value,'');" <% if listgubun = "ITEM" then response.write " checked" %>>��ǰ����Ʈ

			<input type="radio" name="listgubun" value="JUMUN" onClick="ipchulview(this.value,'');" <% if listgubun = "JUMUN" then response.write " checked" %>>�ֹ�����Ʈ
			<% drawipchulmaster "ipchul",ipchul,shopid ,makerid ," onchange=reg('');","" %>
			<script language="javascript">
				ipchulview("<%=listgubun%>","ONLOAD");
			</script>
		<% end if %>
	</td>
	<td align="right">
		<% if printername = "TTP-243_45x22" then %>
			<input type="button" class="button" value="��ǰ�ڵ����(TTP-243)" onClick="BarcodePrint('T')">
			<input type="button" class="button" value="������ڵ����(TTP-243)" onClick="BarcodePrint('G')">
		<% elseif printername = "TTP-243_80x50" then %>
			<% if not(C_IS_Maker_Upche) then %>
				<input type="button" class="button" value="�ε������(TTP-243)" onClick="IndexBarcodePrint();">
			<% end if %>
		<% else %>
			<input type="button" class="button" value="��ǰ�ڵ����(TEC B-FV4)" onClick="BarcodePrint('T')">
			<input type="button" class="button" value="������ڵ����(TEC B-FV4)" onClick="BarcodePrint('G')">

			<% if not(C_IS_Maker_Upche) then %>
				<input type="button" class="button" value="�ε������(TEC B-FV4)" onClick="IndexBarcodePrint();">
			<% end if %>
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        �˻���� : <b><%= oproduct.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= oproduct.FTotalpage %></b>
        &nbsp;&nbsp;�� ��ǰ�� Ư�����ڰ� �ִ� ��� �˻����� �ʽ��ϴ�.
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>�̹���</td>
	<td>�����ڵ�<br><font color=blue>[������ڵ�]</font></td>
	<td>�귣��</td>
	<td>
		��ǰ��<font color=blue>[�ɼǸ�]</font>
		<% if (useforeigndata = "Y") then %>
			<br>������ǰ��<font color=blue>[�����ɼǸ�]</font>
		<% end if %>
	</td>
	<td>
		�Һ��ڰ�
		<% if (useforeigndata = "Y") then %>
			<br>[�����ݾ�]
		<% end if %>
	</td>
	<td>����</td>
	<td>���</td>
</tr>
<% if oproduct.FresultCount > 0 then %>
<% for i=0 to oproduct.FresultCount-1 %>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="location_name" value="<%= oproduct.FItemList(i).Flocation_name %>">
<input type="hidden" id="makerid_<%= i %>" name="locationid" value="<%= oproduct.FItemList(i).Flocationid %>">
<input type="hidden" id="itembarcode_<%= i %>" name="prdcode" value="<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>">
<input type="hidden" id="publicbarcode_<%= i %>" name="generalbarcode" value="<%= oproduct.FItemList(i).Fgeneralbarcode %>">
<input type="hidden" id="itemname_<%= i %>" name="prdname" value="<%= (oproduct.FItemList(i).Fprdname) %>">
<input type="hidden" id="itemoptionname_<%= i %>" name="itemoptionname" value="<%= (oproduct.FItemList(i).Fitemoptionname) %>">
<input type="hidden" id="customerprice_<%= i %>" name="customerprice" value="<%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %>">
<input type="hidden" id="itemname_foreign_<%= i %>" name="lcitemname" value="<%= (oproduct.FItemList(i).Flcitemname) %>">
<input type="hidden" id="itemoptionname_foreign_<%= i %>" name="lcitemoptionname" value="<%= (oproduct.FItemList(i).Flcitemoptionname) %>">
<input type="hidden" id="customerprice_foreign_<%= i %>" name="lcprice" value="<%= round(oproduct.FItemList(i).Flcprice,2) %>">
<tr align="center" bgcolor="#FFFFFF">
	<td width=20><input type="checkbox" id="chk_<%= i %>" name="cksel" onClick="AnCheckClick(this);"></td>
	<td width=50><img src="<%= oproduct.FItemList(i).Fmainimageurl %>" width=50 height=50></td>
	<td width=120>
		<%= AddSpace(oproduct.FItemList(i).Fprdcode) %>
		<% if oproduct.FItemList(i).Fgeneralbarcode <> "" then %>
			<br><font color=blue>[<%= AddSpace(oproduct.FItemList(i).Fgeneralbarcode) %>]</font>
		<% end if %>
	</td>
	<td>
		<%= oproduct.FItemList(i).Flocationid %>
	</td>
	<td align="left">
		<%= AddSpace(oproduct.FItemList(i).Fprdname) %>
		<% if oproduct.FItemList(i).Fitemoptionname <> "" then %>
			<font color=blue>[<%= AddSpace(oproduct.FItemList(i).Fitemoptionname) %>]</font>
		<% end if %>
		<% if (useforeigndata = "Y") then %>
			<br><%= AddSpace(oproduct.FItemList(i).Flcitemname) %><font color=blue>[<%= AddSpace(oproduct.FItemList(i).Flcitemoptionname) %>]</font>
		<% end if %>
	</td>
	<td align="right" width=80>
		<%= AddSpace(FormatNumber(oproduct.FItemList(i).Fcustomerprice,0)) %> ��
		<% if (useforeigndata = "Y") then %>
			<br><%= AddSpace(oproduct.FItemList(i).Flcprice) %> <%= currencyChar %>
		<% end if %>
	</td>
	<td width=60>
		<% if listgubun = "ITEM" then %>
			<input type="text" id="printno_<%= i %>" class="text" name="fixedno" value="<% if (displayrealstockno = "Y") then %><%= oproduct.FItemList(i).Frealstockno %><% else %>1<% end if %>" size="4" maxlength="4" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
		<% elseif listgubun = "JUMUN" then %>
			<input type="text" id="printno_<%= i %>" class="text" name="fixedno" value="<% if (displayrealstockno = "Y") then %><%= oproduct.FItemList(i).Fitemno %><% else %>1<% end if %>" size="4" maxlength="4" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
		<% end if %>
	</td>
	<td width=80>
		<!--<input type="checkbox" name="IsSellPricePrint" checked>�������-->
	</td>
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if oproduct.HasPreScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=oproduct.StartScrollPage-1%>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oproduct.StartScrollPage to oproduct.StartScrollPage + oproduct.FScrollCount - 1 %>
			<% if (i > oproduct.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oproduct.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oproduct.HasNextScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="25" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>

<%
	set oproduct = nothing
	set olocation = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
