<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 주문서 상품 바코드 출력
' Hieditor : 2010.10.21 서동석 생성
'			 2011.02.10 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchulbarcodecls.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchullocationcls.asp"-->
<%
dim i,page,research, masteridx ,boxno , ocstoragemaster ,divcd ,companyid ,userid
dim defaultlocationid ,printername ,barcodetype ,barcodetypestring ,ocstoragedetail
dim locationidfrom ,locationnamefrom ,locationidto ,locationnameto ,maxboxno, cartonboxno, maxcartonboxno
dim IsOneOrderOnly ,siteSeq ,innerboxidx ,innerboxweight ,cartonboxweight ,shopseq
	divcd = requestCheckVar(request("divcd"),32)
	'companyid = requestCheckVar(trim(request("companyid")),32)
	companyid = requestCheckVar(session("ssBctID"), 32)
	masteridx = requestCheckVar(request("masteridx"),32)
	barcodetype = requestCheckVar(request("barcodetype"),32)
	boxno = requestCheckVar(request("boxno"),32)
	cartonboxno = requestCheckVar(request("cartonboxno"),32)
	printername = requestCheckVar(request("printername"),32)

response.write "신매뉴 생성후, 사용중지 매뉴 입니다. 개발팀에 문의 하세요."
response.end

if printername = "" then printername = "TEC_B-FV4"
siteSeq = "10"
if (masteridx = "") then
	masteridx = 0
end if

set ocstoragemaster = new CStorageMaster
	ocstoragemaster.FRectCompanyId = companyid
	ocstoragemaster.FRectMasterIdx = masteridx

if (barcodetype = "offlineorder") then
	barcodetypestring = "오프라인 주문"
	ocstoragemaster.GetOneStorageMaster
else
	barcodetypestring = "오프라인 주문"
	ocstoragemaster.GetOneStorageMaster
end if

if (ocstoragemaster.FOneItem.Fisforeignorder = "Y") then
	barcodetypestring = barcodetypestring + "(해외)"
end if

if C_ADMIN_USER then
elseif (C_IS_SHOP = true) then
	if ((ocstoragemaster.FOneItem.Flocationidfrom <> C_STREETSHOPID) and (ocstoragemaster.FOneItem.Flocationidto <> C_STREETSHOPID)) then
		response.write "<script>alert('"& CTX_The_wrong_approach &"');</script>"
		response.end
	end if
end if

ocstoragemaster.FRectShopId = ocstoragemaster.FOneItem.Flocationidto

IsOneOrderOnly = False

set ocstoragedetail = new CStorageDetail
	ocstoragedetail.FRectCompanyId = companyid
	ocstoragedetail.FRectMasterIdx = masteridx
	ocstoragedetail.FRectIsForeignOrder = ocstoragemaster.FOneItem.Fisforeignorder
	ocstoragedetail.FRectForeignOrderShopid = ocstoragemaster.FOneItem.Fforeignordershopid
	ocstoragedetail.FRectShopId = ocstoragemaster.FOneItem.Flocationidto
	ocstoragedetail.FRectBoxNo = boxno

'상품종류가 3000 가지를 넘기면 문제가 생긴다.
ocstoragedetail.FPageSize = 3000

if (barcodetype = "offlineorder") then

	maxboxno = ocstoragedetail.GetMaxBoxNo

	IsOneOrderOnly = True
	ocstoragedetail.GetStorageDetailList()

elseif (barcodetype = "offlineorderbox") then

	''ocstoragedetail.FRectCartonBoxNo = cartonboxno

	maxboxno = ocstoragedetail.GetMaxBoxNoByBox
	''maxcartonboxno = ocstoragedetail.GetMaxCartonBoxNo(ocstoragemaster.FOneItem.Flocationidto, ocstoragemaster.GetPackingDayList)

	if (boxno <> "") then
		cartonboxno = ocstoragedetail.GetCartonBoxNo(ocstoragemaster.FOneItem.Flocationidto, ocstoragemaster.GetPackingDayList, boxno)
	end if

''rw ocstoragemaster.GetPackingDayList
	if (ocstoragemaster.GetPackingDayList = "") then
		IsOneOrderOnly = True
		ocstoragedetail.GetStorageDetailList
	else
		ocstoragedetail.GetStorageDetailListByBox
	end if

else

	maxboxno = ocstoragedetail.GetMaxBoxNo

	if (ocstoragemaster.GetPackingDayList = "") then
		ocstoragedetail.GetStorageDetailList
	else
		ocstoragedetail.GetStorageDetailListByBox
	end if

end if

divcd = ocstoragemaster.FOneItem.Fdivcd
locationidfrom = ocstoragemaster.FOneItem.Flocationidfrom
locationnamefrom = ocstoragemaster.FOneItem.Flocationnamefrom
locationidto = ocstoragemaster.FOneItem.Flocationidto
locationnameto = ocstoragemaster.FOneItem.Flocationnameto

''' 추가..;;
Dim olocation, currencyunit, currencyChar
set olocation = new CLocation
olocation.FRectCompanyId = companyid
olocation.FRectlocationid = locationidto

if (locationidto <> "") then
	olocation.GetOneLocation

	'useforeigndata = olocation.FOneItem.Fuseforeigndata
	'if (isforeignprint = "") then
	'	isforeignprint = useforeigndata
	'end if
	currencyunit = olocation.FOneItem.Fcurrencyunit
	currencyChar = olocation.FOneItem.FcurrencyChar
end if
Set olocation= Nothing

Function AddSpace(byval str)
	if ((str = "") or (IsNull(str))) then
		AddSpace = "&nbsp;"
	else
		AddSpace = str
	end if
End Function

if isarray(getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")) then
	innerboxweight = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(4,0)
	innerboxidx = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(8,0)
	cartonboxweight = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(9,0)
end if

shopseq = gettenshopidx(ocstoragemaster.FRectShopId)
%>

<% '<script type="text/javascript">drawTTPprintOcxV2('iTTPBar','1,0,0,3');</script> %>
<script type="text/javascript" src="/js/ttpbarcode.js"></script>
<script type="text/javascript" src="/js/barcode.js"></script>
<script type="text/javascript" src="/js/DOSHIBAbarcode.js"></script>
<script type="text/javascript">

//carton box 인덱스 출력
function cartonboxindexprint() {
	var frmmaster = document.frmmaster;

	<% if ocstoragemaster.FRectShopId = "" then %>
		alert('<%= CTX_Please_select %> (<%= CTX_SHOP %>)');
		return;
	<% end if %>

	<% if ocstoragemaster.GetPackingDayList = "" then %>
		alert('<%= CTX_Not_specified %> (<%= CTX_Real_Order_Date %>)');
		return;
	<% end if %>

	<% if cartonboxno = "" or cartonboxno = "0" then %>
		alert('<%= CTX_Please_select %> (<%= CTX_Box %> <%= CTX_number %>');
		return;
	<% end if %>

	<% if cartonboxweight = "" then %>
		alert('<%= CTX_Enter_the_weight %> (CARTONBOX)');
		return;
	<% end if %>

	var paperwidth = frmmaster.paperwidth.value;
	var paperheight = frmmaster.paperheight.value;
	var papermargin = frmmaster.papermargin.value;
	var heightoffset = 0;

    var shopid; var shopname; var packingdate; var cartonboxno; var cartonboxweight; var prdcode; var prdbarcode;
	shopid = '<%=ocstoragemaster.FRectShopId%>';
	shopname = '                <%= locationnameto %>';
	packingdate = '<%=ocstoragemaster.GetPackingDayList%>';
	cartonboxno = '<%=cartonboxno%>';
	cartonboxweight = '<%=cartonboxweight%>';
	prdcode = '                      <%= Format00(2,siteseq) & "-" & Format00(6,shopseq) & "-" & left(ocstoragemaster.GetPackingDayList,4) & mid(ocstoragemaster.GetPackingDayList,6,2) & right(ocstoragemaster.GetPackingDayList,2) & "-" & Format00(3,cartonboxno) %>';
	prdbarcode = '<%= Format00(2,siteseq) & Format00(6,shopseq) & left(ocstoragemaster.GetPackingDayList,4) & mid(ocstoragemaster.GetPackingDayList,6,2) & right(ocstoragemaster.GetPackingDayList,2) & Format00(3,cartonboxno) %>';

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frmmaster.printername.value=='TEC_B-FV4'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("cartonbox 인덱스를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
			TOSHIBA_PAPERWIDTH = 800;
			TOSHIBA_PAPERHEIGHT = 500;
			TOSHIBA_PAPERMARGIN = 3;
			TOSHIBA_HEIGHTOFFSET = heightoffset;
			TOSHIBA_DOMAINNAME = '           CARTON BOX INDEX               ';

			printTOSHIBAcartonboxLabel(shopid, shopname, packingdate, cartonboxno, cartonboxweight, prdcode, prdbarcode);
		}

	}else if (frmmaster.printername.value=='TTP-243_80x50'){
		if (confirm("cartonbox 인덱스를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			TTP_PAPERWIDTH = paperwidth;
			TTP_PAPERHEIGHT = paperheight;
			TTP_PAPERMARGIN = papermargin;
			TTP_HEIGHTOFFSET = heightoffset;
			TTP_DOMAINNAME = '           CARTON BOX INDEX               ';

			printTTPcartonboxLabel(shopid, shopname, packingdate, cartonboxno, cartonboxweight, prdcode, prdbarcode);
		}
	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

//inner box 인덱스 출력
function innerboxindexprint() {
	var frmmaster = document.frmmaster;

	<% if ocstoragemaster.FRectShopId = "" then %>
		alert('<%= CTX_Please_select %> (<%= CTX_SHOP %>)');
		return;
	<% end if %>

	<% if ocstoragemaster.GetPackingDayList = "" then %>
		alert('<%= CTX_Not_specified %> (<%= CTX_Real_Order_Date %>)');
		return;
	<% end if %>

	<% if boxno = "" or boxno = "0" then %>
		alert('<%= CTX_Please_select %> (<%= CTX_Box %> <%= CTX_number %>');
		return;
	<% end if %>

	<% if innerboxweight = "" then %>
		alert('<%= CTX_Enter_the_weight %> (INNERBOX)');
		return;
	<% end if %>

	var paperwidth = frmmaster.paperwidth.value;
	var paperheight = frmmaster.paperheight.value;
	var papermargin = frmmaster.papermargin.value;
	var heightoffset = 0;

    var shopid; var shopname; var packingdate; var innerboxno; var innerboxweight; var prdcode; var prdbarcode;
	shopid = '<%=ocstoragemaster.FRectShopId%>';
	shopname = '                <%= locationnameto %>';
	packingdate = '<%=ocstoragemaster.GetPackingDayList%>';
	innerboxno = '<%=boxno%>';
	innerboxweight = '<%=innerboxweight%>';
	prdcode = '                      <%= Format00(2,siteseq) & "-" & Format00(6,shopseq) & "-" & left(ocstoragemaster.GetPackingDayList,4) & mid(ocstoragemaster.GetPackingDayList,6,2) & right(ocstoragemaster.GetPackingDayList,2) & "-" & Format00(3,boxno) %>';
	prdbarcode = '<%= Format00(2,siteseq) & Format00(6,shopseq) & left(ocstoragemaster.GetPackingDayList,4) & mid(ocstoragemaster.GetPackingDayList,6,2) & right(ocstoragemaster.GetPackingDayList,2) & Format00(3,boxno) %>';

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frmmaster.printername.value=='TEC_B-FV4'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("Innerbox 인덱스를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
			TOSHIBA_PAPERWIDTH = 800;
			TOSHIBA_PAPERHEIGHT = 500;
			TOSHIBA_PAPERMARGIN = 3;
			TOSHIBA_HEIGHTOFFSET = heightoffset;
			TOSHIBA_DOMAINNAME = '                INNER BOX INDEX               ';

			printTOSHIBAinnerboxLabel(shopid, shopname, packingdate, innerboxno, innerboxweight, prdcode, prdbarcode);
		}

	}else if (frmmaster.printername.value=='TTP-243_80x50'){
		if (confirm("Innerbox 인덱스를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			TTP_PAPERWIDTH = paperwidth;
			TTP_PAPERHEIGHT = paperheight;
			TTP_PAPERMARGIN = papermargin;
			TTP_HEIGHTOFFSET = heightoffset;
			TTP_DOMAINNAME = '                INNER BOX INDEX               ';

			printTTPinnerboxLabel(shopid, shopname, packingdate, innerboxno, innerboxweight, prdcode, prdbarcode);
		}
	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

//상품 바코드 출력
function BarcodePrint(barcodetype) {
	var arrbarcode = new Array();
	var chk, itembarcode, makerid, itemname, itemoptionname, customerprice, printno; var ttptype;
	var found = false;
	var isforeignprint; var currencychar; var domainname; var showdomainyn; var showpriceyn; var shopbrandyn; var publicbarcode;

	isforeignprint = document.frmmaster.isforeignprint.value;

	shopbrandyn		= "Y";
	ttptype			= "TTP-243_45x22";
	var paperwidth = frmmaster.paperwidth.value;
	var paperheight = frmmaster.paperheight.value;
	var papermargin = frmmaster.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frmmaster.printpriceyn.value;

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
	if (frmmaster.printername.value=='TEC_B-FV4'){
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
			TOSHIBA_BARCODETYPE = barcodetype;

			printTOSHIBAMultiBarcode(arrbarcode);
		}

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, 0) == true) {
		if (confirm("선택 상품의 바코드를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPMultiBarcode(arrbarcode);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

//인덱스 출력
function IndexBarcodePrint() {
	var arr = new Array();
	var isforeignprint; var printpriceval; var domainname; var showdomainyn; var ttptype;
	var prdname; var itemoptionname; var customerprice; var barcodetype;
	var skipnotinserted; var printno; var shopbrandyn; var currencychar; var showpriceyn;

	isforeignprint = document.frmmaster.isforeignprint.value;
	skipnotinserted = false;

	shopbrandyn		= "Y";
	ttptype			= "TTP-243_80x50";
	barcodetype		= "2";								// T or G or 2(텐바이텐 바코드 or 범용바코드 or 텐텐바코드_범용바코드)
	var paperwidth = frmmaster.paperwidth.value;
	var paperheight = frmmaster.paperheight.value;
	var papermargin = frmmaster.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frmmaster.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	domainname		= "www.10x10.co.kr";
	showdomainyn	= "Y";

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
	if (frmmaster.printername.value=='TEC_B-FV4'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("선택 상품의 바코드를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
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
		if (confirm("선택 상품의 바코드를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPMultiItemLabel(arr);
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

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmmaster" >
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="barcodetype" value="<%= barcodetype %>">
<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>"><%= CTX_SHOP %></td>
	<td width="300"><%= ocstoragemaster.FRectShopId %></td>
	<td width="120" bgcolor="<%= adminColor("gray") %>"><%= CTX_shopname %></td>
	<td>
		<%= locationnameto %>
	</td>
</tr>
<% if (IsOneOrderOnly <> True) then %>
<tr height="30" bgcolor="#FFFFFF">
	<td width="120" bgcolor="<%= adminColor("gray") %>"><%= CTX_Real_Order_Date %></td>
	<td width="300" colspan=3><%= AddSpace(ocstoragemaster.GetPackingDayList) %></td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td width="120" bgcolor="<%= adminColor("gray") %>">Inner박스&nbsp;<%= CTX_number %></td>
	<td width="300">
		<select name="boxno" onChange="frmmaster.submit()">
			<option value="">ALL</option>
			<% If Not IsNULL(maxboxno) then %>
			<% for i=1 to maxboxno %>
				<option value="<%= i %>" <%if (CStr(boxno) = CStr(i)) then %>selected<% end if %>>No. <%= i %></option>
			<% next %>
			<% end if %>
		</select>
	</td>
	<td width="120" bgcolor="<%= adminColor("gray") %>">Inner박스&nbsp;IDX</td>
	<td>
		<% if (boxno <> "") then %>
			10<% if shopseq<>"" then %>-<%= format00(6,shopseq) %><% end if %>-<%= Replace(Replace(ocstoragemaster.GetPackingDayList, "-", ""), "-", "") %>-<%= format00(3,boxno) %>

			<% if printername = "TTP-243_80x50" then %>
				<br><input type="button" value="Inner박스 바코드&nbsp;출력(TTP-243)" onclick="innerboxindexprint()" class="button">
			<% elseif printername = "TEC_B-FV4" then %>
				<br><input type="button" value="Inner박스 바코드&nbsp;출력(TEC B-FV4)" onclick="innerboxindexprint()" class="button">
			<% end if %>
		<% end if %>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td width="120" bgcolor="<%= adminColor("gray") %>">Carton박스 번호</td>
	<td width="300">
		<% If False and Not IsNULL(maxcartonboxno) then %>
		<select name="cartonboxno" onChange="frmmaster.boxno.value = ''; frmmaster.submit()">
			<option value="">ALL</option>
			<% for i=1 to maxcartonboxno %>
			<option value="<%= i %>" <%if (CStr(cartonboxno) = CStr(i)) then %>selected<% end if %>>No. <%= i %></option>
			<% next %>
		</select>
		<% else %>
		<%= cartonboxno %>
		<% end if %>
	</td>
	<td width="120" bgcolor="<%= adminColor("gray") %>">Carton박스&nbsp;IDX</td>
	<td>
		<% if (cartonboxno <> "") then %>
			10<% if shopseq<>"" then %>-<%= format00(6,shopseq) %><% end if %>-<%= Replace(Replace(ocstoragemaster.GetPackingDayList, "-", ""), "-", "") %>-<%= format00(3,cartonboxno) %>

			<% if printername = "TTP-243_80x50" then %>
				<br><input type="button" value="Carton박스 바코드&nbsp;출력(TTP-243)" onclick="cartonboxindexprint()" class="button">
			<% elseif printername = "TEC_B-FV4" then %>
				<br><input type="button" value="Carton박스 바코드&nbsp;출력(TEC B-FV4)" onclick="cartonboxindexprint()" class="button">
			<% end if %>
		<% end if %>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td width="120" bgcolor="<%= adminColor("gray") %>"><%= CTX_relation_order %></td>
	<td width="300" colspan="3">
		<%= AddSpace(ocstoragemaster.GetOrderCodeList) %>
	</td>
</tr>
<% else %>
<tr height="30" bgcolor="#FFFFFF">
	<td width="120" bgcolor="<%= adminColor("gray") %>"><%= CTX_Order_code %></td>
	<td width="300" colspan="3">
		<%= AddSpace(ocstoragemaster.FOneItem.Fordercode) %>
	</td>
</tr>
<% end if %>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>">출력프린트</td>
	<td width="300">
		<select name="printername" onchange="frmmaster.submit();">
			<option value="TTP-243_45x22" <% if printername = "TTP-243_45x22" then response.write " selected" %>>TTP-243 (용지규격&nbsp;45x22)</option>
			<option value="TTP-243_80x50" <% if printername = "TTP-243_80x50" then response.write " selected" %>>TTP-243 (용지규격&nbsp;80x50)</option>
			<option value="TEC_B-FV4" <% if printername = "TEC_B-FV4" then response.write " selected" %>>TEC B-FV4</option>
		</select>
	</td>
	<td width="120" bgcolor="<%= adminColor("gray") %>">바코드&nbsp;용지규격</td>
	<td>
		<% if printername = "TTP-243_45x22" then %>
			<%= CTX_Width %>:<input type="text" name="paperwidth" value="45" size="4" maxlength=9>
			<%= CTX_height %>:<input type="text" name="paperheight" value="22" size="4" maxlength=9>
		<% elseif printername = "TTP-243_80x50" then %>
			<%= CTX_Width %>:<input type="text" name="paperwidth" value="80" size="4" maxlength=9>
			<%= CTX_height %>:<input type="text" name="paperheight" value="50" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4" then %>
			<%= CTX_Width %>:<input type="text" name="paperwidth" value="450" size="4" maxlength=9>
			<%= CTX_height %>:<input type="text" name="paperheight" value="220" size="4" maxlength=9>
		<% end if %>
		<%= CTX_blank %>:<input type="text" name="papermargin" value="3" size="4" maxlength=9>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td width="120" bgcolor="<%= adminColor("gray") %>"><%= CTX_show %>&nbsp;<%= CTX_Description %></td>
	<td width="300">
		<select name="isforeignprint">
			<option value="N"><%= CTX_Domestic %>&nbsp;<%= CTX_Description %></option>
			<option value="Y" <% if (ocstoragemaster.FOneItem.Fisforeignorder = "Y") then %>selected<% end if %>><%= CTX_foreign %>&nbsp;<%= CTX_Description %></option>
		</select>
	</td>
	<td width="120" bgcolor="<%= adminColor("gray") %>"><%= CTX_cost %>&nbsp;<%= CTX_show %></td>
	<td>
		<select name="printpriceyn">
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
	</td>
</tr>
</form>
</table>

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="pagesetfrm" method="get">
<tr>
	<td align="left">
		<a href="http://imgstatic.10x10.co.kr/offshop/sample/print/도시바_TEC B-FV4_상품바코드_셋팅법_v1.docx" target="_blank">TEC B-FV4 셋팅값 다운로드</a>
		<!--
        ※ 바코드 용지 규격
		<% if printername = "TTP-243_45x22" then %>
			<br>넓이:<input type="text" name="paperwidth" value="45" size="4" maxlength=9>
			높이:<input type="text" name="paperheight" value="22" size="4" maxlength=9>
		<% elseif printername = "TTP-243_80x50" then %>
			<br>넓이:<input type="text" name="paperwidth" value="80" size="4" maxlength=9>
			높이:<input type="text" name="paperheight" value="50" size="4" maxlength=9>
		<% end if %>
		여백:<input type="text" name="papermargin" value="3" size="4" maxlength=9>
		-->
	</td>
	<td align="right">
		<% if printername = "TTP-243_45x22" then %>
			<input type="button" class="button" value="상품바코드&nbsp;출력(TTP-243)" onClick="BarcodePrint('T')">
			<input type="button" class="button" value="범용바코드&nbsp;출력(TTP-243)" onClick="BarcodePrint('G')">
		<% elseif printername = "TTP-243_80x50" then %>
			<input type="button" class="button" value="인덱스바코드&nbsp;출력(TTP-243)" onClick="IndexBarcodePrint();">
		<% else %>
			<input type="button" class="button" value="상품바코드&nbsp;출력(TEC B-FV4)" onClick="BarcodePrint('T')">
			<input type="button" class="button" value="범용바코드&nbsp;출력(TEC B-FV4)" onClick="BarcodePrint('G')">
			<input type="button" class="button" value="인덱스바코드&nbsp;출력(TEC B-FV4)" onClick="IndexBarcodePrint();">
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td><%= CTX_Box %><br><%= CTX_number %></td>
	<td><%= CTX_Order_code %></td>
	<td><%= CTX_Image %></td>
	<td>물류바코드<br><font color=blue>[범용 바코드]</font></td>
	<td align="left">
		<%= CTX_Description %><font color=blue>[<%= CTX_Description_Option %>]</font>
		<% if (ocstoragemaster.FOneItem.Fisforeignorder = "Y") then %>
			<br><%= CTX_foreign %>&nbsp;<%= CTX_Description %><font color=blue>[<%= CTX_foreign %>&nbsp;<%= CTX_Description_Option %>]</font>
		<% end if %>
	</td>
	<td>
		<%= CTX_consumer_price %>
		<% if (ocstoragemaster.FOneItem.Fisforeignorder = "Y") then %>
			<br>[<%= CTX_foreign %>&nbsp;<%= CTX_cost %>]
		<% end if %>
	</td>
	<td><%= CTX_quantity %></td>
	<td><%= CTX_Note %></td>
</tr>
<% if ocstoragedetail.FresultCount > 0 then %>
<% for i=0 to ocstoragedetail.FresultCount-1 %>
<form name="frmBuyPrc_<%= i %>">
<% if ocstoragedetail.FItemList(i).Fuseyn = "Y"	then %>
<input type="hidden" name="location_name" value="<%= ocstoragedetail.FItemList(i).Flocationname %>">
<input type="hidden" id="makerid_<%= i %>" name="locationid" value="<%= (ocstoragedetail.FItemList(i).Flocationid) %>">
<input type="hidden" id="itembarcode_<%= i %>" name="prdcode" value="<%= ocstoragedetail.FItemList(i).Fprdcode %>">
<input type="hidden" id="publicbarcode_<%= i %>" name="generalbarcode" value="<%= ocstoragedetail.FItemList(i).Fgeneralbarcode %>">
<input type="hidden" id="itemname_<%= i %>" name="prdname" value="<%= (ocstoragedetail.FItemList(i).Fprdname) %>">
<input type="hidden" id="itemoptionname_<%= i %>" name="itemoptionname" value="<%= (ocstoragedetail.FItemList(i).Fitemoptionname) %>">
<input type="hidden" id="customerprice_<%= i %>" name="customerprice" value="<%= FormatNumber(ocstoragedetail.FItemList(i).Fcustomerprice,0) %>">
<input type="hidden" id="itemname_foreign_<%= i %>" name="lcitemname" value="<%= (ocstoragedetail.FItemList(i).Flcitemname) %>">
<input type="hidden" id="itemoptionname_foreign_<%= i %>" name="lcitemoptionname" value="<%= (ocstoragedetail.FItemList(i).Flcitemoptionname) %>">
<input type="hidden" id="customerprice_foreign_<%= i %>" name="lcprice" value="<%= round(ocstoragedetail.FItemList(i).Flcprice,2) %>">
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" id="chk_<%= i %>" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><%= AddSpace(ocstoragedetail.FItemList(i).Fboxno) %></td>
	<td><%= AddSpace(ocstoragedetail.FItemList(i).Fbaljucode) %></td>
	<td><img src="<%= ocstoragedetail.FItemList(i).Fmainimageurl %>" width="50"></td>
	<td>
		<%= AddSpace(ocstoragedetail.FItemList(i).Fprdcode) %><br><font color=blue>[<%= AddSpace(ocstoragedetail.FItemList(i).Fgeneralbarcode) %>]
	</td>
	<td align="left">
		<%= AddSpace(ocstoragedetail.FItemList(i).Fprdname) %><font color=blue>[<%= AddSpace(ocstoragedetail.FItemList(i).Fitemoptionname) %>]</font>
	<% if (ocstoragemaster.FOneItem.Fisforeignorder = "Y") then %>
		<br><%= AddSpace(ocstoragedetail.FItemList(i).Flcitemname) %><font color=blue>[<%= AddSpace(ocstoragedetail.FItemList(i).Flcitemoptionname) %>]</font>
	<% end if %>
	</td>
	<td align="right">
		\ <%= AddSpace(FormatNumber(ocstoragedetail.FItemList(i).Fcustomerprice,0)) %>
	<% if (ocstoragemaster.FOneItem.Fisforeignorder = "Y") then %>
		<br><%= currencyChar %> <%= AddSpace(ocstoragedetail.FItemList(i).Flcprice) %>
	<% end if %>
	</td>
	<td><input type="text" id="printno_<%= i %>" name="fixedno" value="<%= ocstoragedetail.FItemList(i).Ffixedno %>" size="4" maxlength="4" onFocus="this.select()" class="text"></td>
	<td><!--<input type="checkbox" name="IsSellPricePrint" checked>가격출력--></td>
</tr>
<% end if %>
</form>
<% next %>

<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[<%= CTX_search_returns_no_results %>]</td>
</tr>
<% end if %>
</table>

<%
set ocstoragemaster = nothing
set ocstoragedetail = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
