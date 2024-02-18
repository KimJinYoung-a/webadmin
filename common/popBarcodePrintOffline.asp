<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 상품 바코드 출력
' Hieditor : 2010.10.26 서동석 생성
'			 2011.02.10 한용민 수정
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

response.write "신매뉴 생성후, 사용중지 매뉴 입니다. 개발팀에 문의 하세요."
response.end

'/매장일경우 본인 매장만 사용가능
if (C_IS_SHOP) then
	'/어드민권한 점장 미만
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
currencyChar = "￦"

set oproduct = new CProduct
	oproduct.FCurrpage = page
	oproduct.FPageSize = 100
	oproduct.FRectLocationId = shopid				'이동처
	oproduct.FRectLocationIdMaker = makerid
	oproduct.FRectPrdCode = prdcode
	oproduct.FRectItemID = itemid
	oproduct.FRectPrdName = html2db(prdname)
	oproduct.FRectGeneralBarcode = generalbarcode
	''oproduct.FRectUseYN = "Y"                         ''사용구분 상관없이 전체 표시
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
	    response.write "    alert('매장을 선택해 주세요');"
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
//			alert("상품코드는 숫자만 가능합니다.");
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

//인덱스 출력
function IndexBarcodePrint() {
	var arr = new Array();
	var isforeignprint; var printpriceval; var domainname; var showdomainyn; var ttptype;
	var prdname; var itemoptionname; var customerprice; var barcodetype;
	var skipnotinserted; var printno; var shopbrandyn; var currencychar; var showpriceyn;

	isforeignprint = document.frm.isforeignprint.value;
	skipnotinserted = false;

	shopbrandyn		= "Y";
	ttptype			= "TTP-243_80x50";
	barcodetype		= "2";								// T or G or 2(텐바이텐 바코드 or 범용바코드 or 텐텐바코드_범용바코드)
	var paperwidth = frm.paperwidth.value;
	var paperheight = frm.paperheight.value;
	var papermargin = frm.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frm.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "￦";
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

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("선택 상품의 인덱스를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
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

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) == true) {
		if (confirm("선택 상품의 인덱스를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPMultiItemLabel(arr);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

//상품코드 출력
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
		currencychar = "￦";
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
		alert("선택된 상품이 없습니다.");
		return;
	}

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("선택 상품의 바코드를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
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

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, btype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, 0) == true) {
		if (confirm("선택 상품의 바코드를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPMultiBarcode(arrbarcode);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

String.prototype.trim = function() {
    return this.replace(/^\s+|\s+$/g, "");
};

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			<% end if %>
		<% else %>
			<% if (C_IS_Maker_Upche) then %>
				* 매장 : <% drawBoxDirectIpchulOffShopByMakerchfg "shopid", shopid, makerid, " onchange='reg("""");'", " 'B011','B012','B013'" %>
			<% else %>
				* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			<% end if %>
		<% end if %>
		&nbsp;&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>

	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<% if (C_IS_Maker_Upche) then %>
			* 브랜드 : <%= makerid %>
			<input type="hidden" name="makerid" value="<%= makerid %>">
			&nbsp;&nbsp;
		<% else %>
			* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;&nbsp;
		<% end if %>

		* 상품코드 : 
		<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea><!--onKeyPress="if (event.keyCode == 13) 	reg('');"-->
		&nbsp;&nbsp;
		* 상품명 : <input type="text" class="text" name="prdname" value="<%= prdname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;&nbsp;
		* 물류코드 :
		<input type="text" class="text" name="prdcode" value="<%= prdcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">

	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<% if not(C_IS_Maker_Upche) then %>
			* 매장별상품명 :
			<input type="text" class="text" name="shopitemname" value="<%= shopitemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
			&nbsp;&nbsp;
			<input type="checkbox" class="checkbox" name="currentstockexist" value="Y" onClick="reg('');" <% if (currentstockexist = "Y") then %>checked<% end if %>> 입고내역존재상품
			&nbsp;&nbsp;
			<input type="checkbox" class="checkbox" name="realstockonemore" value="Y" onClick="reg('');" <% if (realstockonemore = "Y") then %>checked<% end if %>> 재고존재상품
			&nbsp;&nbsp;
			<input type="checkbox" class="checkbox" name="shopitemnameinserted" value="Y" onClick="reg('');" <% if (shopitemnameinserted = "Y") then %>checked<% end if %>> 매장별상품명등록상품만
			&nbsp;&nbsp;
		<% end if %>

		<% if listgubun = "ITEM" then %>
			* 범용바코드 :
			<input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
			&nbsp;&nbsp;

			<% if not(C_IS_Maker_Upche) then %>
				<input type="checkbox" class="checkbox" name="displayrealstockno" value="Y" onClick="reg('');" <% if (displayrealstockno = "Y") then %>checked<% end if %>>재고로 수량설정
			<% end if %>
		<% elseif listgubun = "JUMUN" then %>
			<input type="checkbox" class="checkbox" name="displayrealstockno" value="Y" onClick="reg('');" <% if (displayrealstockno = "Y") then %>checked<% end if %>>
			입출고내역으로 수량설정
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		※ 프린터 설정 :
		<select name="printername" onchange="reg('');">
			<option value="TTP-243_45x22" <% if printername = "TTP-243_45x22" then response.write " selected" %>>TTP-243 (규격45x22)</option>
			<option value="TTP-243_80x50" <% if printername = "TTP-243_80x50" then response.write " selected" %>>TTP-243 (규격80x50)</option>
			<option value="TEC_B-FV4" <% if printername = "TEC_B-FV4" then response.write " selected" %>>TEC B-FV4</option>
		</select>
		&nbsp;&nbsp;
		* 표시상품명 :
		<select name="isforeignprint" onchange="reg('');">
			<option value="N" <% if (isforeignprint = "N") then %>selected<% end if %>>국내상품명</option>
			<option value="Y" <% if (isforeignprint = "Y") then %>selected<% end if %>>샵별상품명</option>
		</select>
		&nbsp;&nbsp;
		* 금액표시여부 :
		<select name="printpriceyn" onchange="reg('');">
			<option value="Y" <% if (printpriceyn = "Y") then %>selected<% end if %>>금액표시</option>
			<option value="N" <% if (printpriceyn = "N") then %>selected<% end if %>>금액표시안함</option>
		</select>
		&nbsp;&nbsp;
		* 브랜드표시 :
		<select name="makeriddispyn" onchange="reg('');">
			<option value="Y" <% if (makeriddispyn = "Y") then %>selected<% end if %>>브랜드표시</option>
			<option value="N" <% if (makeriddispyn = "N") then %>selected<% end if %>>브랜드표시안함</option>
		</select>
        <br>
        * 바코드 용지 규격 -
		<% if printername = "TTP-243_45x22" then %>
			넓이:<input type="text" name="paperwidth" value="45" size="4" maxlength=9>
			높이:<input type="text" name="paperheight" value="22" size="4" maxlength=9>
		<% elseif printername = "TTP-243_80x50" then %>
			넓이:<input type="text" name="paperwidth" value="80" size="4" maxlength=9>
			높이:<input type="text" name="paperheight" value="50" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4" then %>
			넓이:<input type="text" name="paperwidth" value="450" size="4" maxlength=9>
			높이:<input type="text" name="paperheight" value="220" size="4" maxlength=9>
		<% end if %>

		여백:<input type="text" name="papermargin" value="3" size="4" maxlength=9>

		<script language="javascript">
			jsSetSelectBoxColor();
		</script>
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<a href="http://imgstatic.10x10.co.kr/offshop/sample/print/도시바_TEC B-FV4_물류바코드_셋팅법.docx" target="_blank">TEC B-FV4 셋팅값 다운로드</a>

		<% if shopid <> "" then %>
			<br><font color="red">리스트기준 :</font>
			<input type="radio" name="listgubun" value="ITEM" onClick="ipchulview(this.value,'');" <% if listgubun = "ITEM" then response.write " checked" %>>상품리스트

			<input type="radio" name="listgubun" value="JUMUN" onClick="ipchulview(this.value,'');" <% if listgubun = "JUMUN" then response.write " checked" %>>주문리스트
			<% drawipchulmaster "ipchul",ipchul,shopid ,makerid ," onchange=reg('');","" %>
			<script language="javascript">
				ipchulview("<%=listgubun%>","ONLOAD");
			</script>
		<% end if %>
	</td>
	<td align="right">
		<% if printername = "TTP-243_45x22" then %>
			<input type="button" class="button" value="상품코드출력(TTP-243)" onClick="BarcodePrint('T')">
			<input type="button" class="button" value="범용바코드출력(TTP-243)" onClick="BarcodePrint('G')">
		<% elseif printername = "TTP-243_80x50" then %>
			<% if not(C_IS_Maker_Upche) then %>
				<input type="button" class="button" value="인덱스출력(TTP-243)" onClick="IndexBarcodePrint();">
			<% end if %>
		<% else %>
			<input type="button" class="button" value="상품코드출력(TEC B-FV4)" onClick="BarcodePrint('T')">
			<input type="button" class="button" value="범용바코드출력(TEC B-FV4)" onClick="BarcodePrint('G')">

			<% if not(C_IS_Maker_Upche) then %>
				<input type="button" class="button" value="인덱스출력(TEC B-FV4)" onClick="IndexBarcodePrint();">
			<% end if %>
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        검색결과 : <b><%= oproduct.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= oproduct.FTotalpage %></b>
        &nbsp;&nbsp;※ 상품명에 특수문자가 있는 경우 검색되지 않습니다.
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>이미지</td>
	<td>물류코드<br><font color=blue>[범용바코드]</font></td>
	<td>브랜드</td>
	<td>
		상품명<font color=blue>[옵션명]</font>
		<% if (useforeigndata = "Y") then %>
			<br>샵별상품명<font color=blue>[샵별옵션명]</font>
		<% end if %>
	</td>
	<td>
		소비자가
		<% if (useforeigndata = "Y") then %>
			<br>[샵별금액]
		<% end if %>
	</td>
	<td>수량</td>
	<td>비고</td>
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
		<%= AddSpace(FormatNumber(oproduct.FItemList(i).Fcustomerprice,0)) %> 원
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
		<!--<input type="checkbox" name="IsSellPricePrint" checked>가격출력-->
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
	<td colspan="25" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>

<%
	set oproduct = nothing
	set olocation = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
