<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 온라인 바코드 출력
' Hieditor : 2016.12.15 한용민 생성
'/////////////////// 이파일 수정시 밑에 파일도 모두 동일하게 같이 고쳐야 한다. ////////////////////////
' SCM : /common/barcode/inc_barcodeprint_on.asp
' 		/partner/common/barcode/inc_barcodeprint_on.asp
' LOGICS : /v2/common/barcode/inc_barcodeprint_on.asp
'###########################################################
%>
<%
dim divcd ,companyid ,userid, itemname, isfixed, jumunwait, IsForeign_confirmed, IsForeignOrder
dim defaultlocationid ,barcodetypestring, sellyn, usingyn
dim locationidfrom ,locationnamefrom ,locationidto ,locationnameto
dim IsOneOrderOnly ,siteSeq ,innerboxidx ,innerboxweight ,cartonboxweight , shopseq
	listgubun 		= requestCheckVar(request("listgubun"), 32)
	divcd = requestCheckVar(request("divcd"),32)
	'companyid = requestCheckVar(trim(request("companyid")),32)
	companyid = requestCheckVar(session("ssBctID"), 32)
	printername = requestCheckVar(request("printername"),32)
	isforeignprint = requestCheckVar(request("isforeignprint"),1)
	page 			= requestCheckVar(request("page"),32)
	makerid = requestCheckVar(request("makerid"),32)
	itemid      = requestCheckvar(request("itemid"),255)
	itemname    = requestCheckvar(request("itemname"),64)
	prdcode 		= requestCheckVar(request("prdcode"),32)
	generalbarcode 	= requestCheckVar(request("generalbarcode"),32)
	sellyn      = requestCheckvar(request("sellyn"),10)
	usingyn     = requestCheckvar(request("usingyn"),10)
	research 		= requestCheckVar(request("research"),32)
	printpriceyn = requestCheckVar(request("printpriceyn"),1)
	makeriddispyn = requestCheckVar(request("makeriddispyn"),1)
	papername = requestCheckVar(request("papername"),2)
	itemoptionyn = requestCheckVar(request("itemoptionyn"),1)
	titledispyn = requestCheckVar(request("titledispyn"),1)

isdispsql=true
isdispconfirm=true
if papername = "" then papername = "BQ"
jumunwait = false
IsForeignOrder = false		'/업체접수주문
IsForeign_confirmed = false		'/업체접수주문 컨펌완료여부

'/매장일경우
if (C_IS_SHOP) then
	'/가맹점 일경우
	if getoffshopdiv(C_STREETSHOPID) = "3" then
		isdispconfirm=false
	end if
else
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if
if page = "" then page = 1
iPageSize=50
'if sellyn = "" and research <> "on" then sellyn = "Y"
if listgubun = "" then listgubun = "ITEM"
if listgubun="ITEM" then
	if usingyn = "" and research <> "on" then usingyn = "Y"	
end if
if printpriceyn = "" then printpriceyn = "Y"
if printername = "" then printername = "TEC_B-FV4_45x22"
if itemoptionyn = "" then itemoptionyn = "Y"
if titledispyn = "" then titledispyn = "Y"
siteSeq = "10"

if itemid<>"" then
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" then
			arrItemid = arrItemid & Trim(getNumeric(arrTemp(iA))) & ","
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

set oproduct = new CStorageDetail

'/상품리스트
if listgubun = "ITEM" then
	oproduct.FPageSize = iPageSize
	oproduct.FCurrPage = page
	oproduct.FRectMakerid = makerid
	oproduct.FRectItemid       = itemid
	oproduct.FRectItemName     = itemname
	oproduct.FRectPrdCode = prdcode
	oproduct.FRectGeneralBarcode = generalbarcode
	oproduct.FRectisforeignprint = isforeignprint
	oproduct.FRectSellYN       = sellyn
	oproduct.FRectIsUsing      = usingyn

	if makerid<>"" or itemname<>"" or prdcode<>"" or itemid<>"" or generalbarcode<>"" then
		oproduct.GetProductListOnline
	else
		isdispsql=false
	end if
end if

if isforeignprint="" then isforeignprint="N"
if currencyunit="" then
	if isforeignprint="N" then
		currencyunit = "KRW"
	else
		currencyunit = "USD"
	end if
end if
if currencyChar="" then
	if isforeignprint="N" then
		currencyChar = "￦"
	else
		currencyChar = "$"
	end if
end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

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

// 폼텍 바코드 출력
function CssFORMTECBarcodeprint(barcodetype) {
	frmList.barcodetype.value=barcodetype;

	if(!$(frmList.cksel).is(':checked')) {
		alert('선택된 상품이 없습니다.');
		return;
	}

	var browser = navigator.userAgent.toLowerCase();
	if ( -1 != browser.indexOf('chrome') ){
		frmList.target = "FrameCKP" ;
		frmList.action = "/common/barcode/CssBarcodeprint_A4.asp" ;
		frmList.submit() ;
	}else if ( -1 != browser.indexOf('trident') ){	// 익스플로러 출력 기능을 제거 했더니 바로 업체에서 문의옴.. 익스플로러 모드 가 편하다고 출력되게 해달라고..
		try{
			AddArr();
		}catch (e) {
			alert("- 도구 > 인터넷 옵션 > 보안 탭 > 신뢰할 수 있는 사이트 선택\n   1. 사이트 버튼 클릭 > 사이트 추가\n   2. 사용자 지정 수준 클릭 > 스크립팅하기 안전하지 않은 것으로 표시된 ActiveX 컨트롤 (사용)으로 체크\n\n※ 위 설정은 프린트 기능을 사용하기 위함임");
		}
	}else{
		alert("사용하시는 브라우저를 확인해주세요.\n- 크로미움 엔진을 사용한 브라우저(Chrome, Edge, Whale 등)에서 출력하실 수 있습니다.");
	}
}

function AddData(itemid, itemoption, prdname, prdoptionname, socname, itemprice, itemtype, fixedno){
	iaxobject.AddData(itemid, itemoption, prdname, prdoptionname, socname, itemprice, itemtype, fixedno);
}

//AddData(v,'0000','아이템명','옵션명','브랜드',3000,'T','5')
function AddArr(){
	var makeriddisp;
	var printprice; var showpriceyn; var saleyn;

	iaxobject.ClearItem();
	//iaxobject.setTitleVisible(true);

	$("input[name='cksel']:checked").each(function(){
		var vid = $(this).val()-1; // 체크id

		//브랜드표시
		if (frm.makeriddispyn.value != 'N'){
			makeriddisp = makeriddisp = $(frmList.socname).eq(vid).val();
		}else{
			makeriddisp = '';
		}

		//가격표시
		switch (frm.printpriceyn) {
			case 'C':	//할인가표시
				if(frmList.saleyn.value=="Y") {
					//할인가
					printprice = $(frmList.saleprice).eq(vid).val().trim();
				} else {
					//소비자가
					printprice = $(frmList.customerprice).eq(vid).val().trim();
				}
				break;
			case 'R':	//판매가표시
				printprice = $(frmList.sellprice).eq(vid).val().trim();
				break;
			default:
				//소비자가 표시
				printprice = $(frmList.customerprice).eq(vid).val().trim();
				break;
		}

		// 데이터 추가
		if ($(frmList.itemid).eq(vid).val()*1>=1000000){
			AddData($(frmList.itemid).eq(vid).val(),
				$(frmList.itemoption).eq(vid).val(),
				$(frmList.prdname).eq(vid).val(),
				$(frmList.prdoptionname).eq(vid).val(),
				makeriddisp, printprice,
				$(frmList.itemgubun).eq(vid).val()*10,
				$(frmList.fixedno).eq(vid).val());
		}else{
			AddData($(frmList.itemid).eq(vid).val(),
				$(frmList.itemoption).eq(vid).val(),
				$(frmList.prdname).eq(vid).val(),
				$(frmList.prdoptionname).eq(vid).val(),
				makeriddisp, printprice,
				$(frmList.itemgubun).eq(vid).val(),
				$(frmList.fixedno).eq(vid).val());
		}

	});

	iaxobject.ShowFrm();
}

// 상품 바코드 출력
function CssBarcodeprint(barcodetype) {
	var isforeignprint = frm.isforeignprint.value;
	var currencychar="";
	frmList.barcodetype.value=barcodetype;

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	frmList.currencychar.value=currencychar;

	if(!$(frmList.cksel).is(':checked')) {
		alert('선택된 상품이 없습니다.');
		return;
	}

	var browser = navigator.userAgent.toLowerCase();
	if ( -1 != browser.indexOf('chrome') ){
		frmList.target = "FrameCKP" ;
		frmList.action = "/common/barcode/CssBarcodeprint_45x22.asp" ;
		frmList.submit() ;

	}else{
		alert("사용하시는 브라우저를 확인해주세요.\n- 크로미움 엔진을 사용한 브라우저(Chrome, Edge, Whale 등)에서 출력하실 수 있습니다.");
	}
}

// 쥬얼리 바코드 출력
function jewelleryCssBarcodePrint(barcodetype) {
	var isforeignprint = frm.isforeignprint.value;
	var currencychar="";
	frmList.barcodetype.value=barcodetype;

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	frmList.currencychar.value=currencychar;

	if(!$(frmList.cksel).is(':checked')) {
		alert('선택된 상품이 없습니다.');
		return;
	}

	var browser = navigator.userAgent.toLowerCase();
	if ( -1 != browser.indexOf('chrome') ){
		frmList.target = "FrameCKP" ;
		frmList.action = "/common/barcode/CssBarcodeprint_35x15.asp" ;
		frmList.submit() ;

	}else{
		alert("사용하시는 브라우저를 확인해주세요.\n- 크로미움 엔진을 사용한 브라우저(Chrome, Edge, Whale 등)에서 출력하실 수 있습니다.");
	}
}

// 인덱스 바코드 출력
function IndexCssBarcodePrint() {
	var isforeignprint = frm.isforeignprint.value;
	var currencychar="";

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	frmList.currencychar.value=currencychar;

	if(!$(frmList.cksel).is(':checked')) {
		alert('선택된 상품이 없습니다.');
		return;
	}

	var browser = navigator.userAgent.toLowerCase();
	if ( -1 != browser.indexOf('chrome') ){
		frmList.target = "FrameCKP" ;
		frmList.action = "/common/barcode/CssBarcodeprint_80x50.asp" ;
		frmList.submit() ;

	}else{
		alert("사용하시는 브라우저를 확인해주세요.\n- 크로미움 엔진을 사용한 브라우저(Chrome, Edge, Whale 등)에서 출력하실 수 있습니다.");
	}
}

//인덱스 출력
function IndexBarcodePrint() {
	var arr = new Array();
	var isforeignprint; var printpriceval; var domainname; var showdomainyn; var ttptype;
	var found = false;
	var prdname; var itemoptionname; var customerprice; var barcodetype;
	var skipnotinserted; var printno; var shopbrandyn; var currencychar; var showpriceyn;
	var saleprice; var couponprice; var saleyn; var couponyn; var socname, brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode;

	isforeignprint = document.frm.isforeignprint.value;
	skipnotinserted = false;

	shopbrandyn = frm.makeriddispyn.value;
	if (shopbrandyn=="") shopbrandyn="Y";
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
	showdomainyn	= frm.titledispyn.value;

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("cksel" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		if (chk.checked == true) {
			saleprice = document.getElementById("saleprice_" + i).value.trim();
			couponprice = document.getElementById("couponprice_" + i).value.trim();
			saleyn = document.getElementById("saleyn_" + i).value.trim();
			couponyn = document.getElementById("couponyn_" + i).value.trim();

			//해외 상품명
			if (isforeignprint == "Y") {
				itemname = document.getElementById("itemname_foreign_" + i).value.trim();
				itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
				customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();

			//국내 상품명
			} else {
				itemname = document.getElementById("itemname_" + i).value.trim();
				itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();

				//할인가 표시
				if (showpriceyn=='C'){
					//할인
					if (saleyn=='Y'){
						customerprice = saleprice;

					//쿠폰할인
					}else if (couponyn=='Y'){
						customerprice = couponprice;

					//소비자가
					}else{
						customerprice = document.getElementById("customerprice_" + i).value.trim();
					}

				//판매가 표시
				}else if (showpriceyn=='R'){
					customerprice = document.getElementById("sellprice_" + i).value.trim();

				//소비자가 표시
				}else{
					customerprice = document.getElementById("customerprice_" + i).value.trim();
				}
			}

			itembarcode = document.getElementById("itembarcode_" + i).value.trim();
			publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
			itembarcode = itembarcode + "_" + publicbarcode;

			makerid = document.getElementById("makerid_" + i).value.trim();
			socname = document.getElementById("socname_" + i).value.trim();
			//printno = document.getElementById("printno_" + i).value.trim();
			printno = 1;
			brandrackcode = document.getElementById("prtidx_" + i).value.trim();
			itemrackcode = document.getElementById("itemrackcode_" + i).value.trim();
			subitemrackcode = document.getElementById("subitemrackcode_" + i).value.trim();

			if (printno*1 != 0) {
				var v = new BarcodeDataClass_index(itembarcode, socname, itemname, itemoptionname, customerprice, printno, '', '', '', brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode);
				arr.push(v);
			}
		}
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4_80x50'){
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

//상품 바코드 출력
function BarcodePrint(barcodetype) {
	var arrbarcode = new Array();
	var chk, itembarcode, makerid, itemname, itemoptionname, customerprice, printno; var ttptype;
	var found = false;
	var isforeignprint; var currencychar; var domainname; var showdomainyn; var showpriceyn; var shopbrandyn; var publicbarcode;
	var saleprice; var couponprice; var saleyn; var couponyn; var socname; var socname_kor;

	isforeignprint = document.frm.isforeignprint.value;

	shopbrandyn = frm.makeriddispyn.value;
	if (shopbrandyn=="") shopbrandyn="Y";
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
	showdomainyn	= frm.titledispyn.value;

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("cksel" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		saleprice = document.getElementById("saleprice_" + i).value.trim();
		couponprice = document.getElementById("couponprice_" + i).value.trim();
		saleyn = document.getElementById("saleyn_" + i).value.trim();
		couponyn = document.getElementById("couponyn_" + i).value.trim();

		//해외 상품명
		if (isforeignprint == "Y") {
			itemname = document.getElementById("itemname_foreign_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
			customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();

		//국내 상품명
		} else {
			itemname = document.getElementById("itemname_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();

			//할인가 표시
			if (showpriceyn=='C'){
				//할인
				if (saleyn=='Y'){
					customerprice = saleprice;

				//쿠폰할인
				}else if (couponyn=='Y'){
					customerprice = couponprice;

				//소비자가
				}else{
					customerprice = document.getElementById("customerprice_" + i).value.trim();
				}

			//판매가 표시
			}else if (showpriceyn=='R'){
				customerprice = document.getElementById("sellprice_" + i).value.trim();

			//소비자가 표시
			}else{
				customerprice = document.getElementById("customerprice_" + i).value.trim();
			}
		}

		itembarcode = document.getElementById("itembarcode_" + i).value.trim();
		publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
		itembarcode = itembarcode + "_" + publicbarcode;
		makerid = document.getElementById("makerid_" + i).value.trim();
		socname = document.getElementById("socname_" + i).value.trim();
		socname_kor = document.getElementById("socname_kor_" + i).value.trim();
		printno = document.getElementById("printno_" + i).value.trim();

		var v = new BarcodeDataClass(itembarcode, makerid, itemname, itemoptionname, customerprice, printno, '', socname, socname_kor);
		arrbarcode.push(v);
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4_45x22'){
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

//쥬얼리 바코드 출력
function jewellery_BarcodePrint(barcodetype) {
	var arrbarcode = new Array();
	var chk, itembarcode, makerid, itemname, itemoptionname, customerprice, printno; var ttptype;
	var found = false;
	var isforeignprint; var currencychar; var domainname; var showdomainyn; var showpriceyn; var shopbrandyn; var publicbarcode;
	var saleprice; var couponprice; var saleyn; var couponyn; var socname;

	isforeignprint = document.frm.isforeignprint.value;

	shopbrandyn = frm.makeriddispyn.value;
	if (shopbrandyn=="") shopbrandyn="Y";
	ttptype			= "TTP-243_35x15";
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
	showdomainyn	= frm.titledispyn.value;

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("cksel" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		saleprice = document.getElementById("saleprice_" + i).value.trim();
		couponprice = document.getElementById("couponprice_" + i).value.trim();
		saleyn = document.getElementById("saleyn_" + i).value.trim();
		couponyn = document.getElementById("couponyn_" + i).value.trim();

		//해외 상품명
		if (isforeignprint == "Y") {
			itemname = document.getElementById("itemname_foreign_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
			customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();

		//국내 상품명
		} else {
			itemname = document.getElementById("itemname_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();

			//할인가 표시
			if (showpriceyn=='C'){
				//할인
				if (saleyn=='Y'){
					customerprice = saleprice;

				//쿠폰할인
				}else if (couponyn=='Y'){
					customerprice = couponprice;

				//소비자가
				}else{
					customerprice = document.getElementById("customerprice_" + i).value.trim();
				}

			//판매가 표시
			}else if (showpriceyn=='R'){
				customerprice = document.getElementById("sellprice_" + i).value.trim();

			//소비자가 표시
			}else{
				customerprice = document.getElementById("customerprice_" + i).value.trim();
			}
		}

		itembarcode = document.getElementById("itembarcode_" + i).value.trim();
		publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
		itembarcode = itembarcode + "_" + publicbarcode;
		makerid = document.getElementById("makerid_" + i).value.trim();
		socname = document.getElementById("socname_" + i).value.trim();
		printno = document.getElementById("printno_" + i).value.trim();

		var v = new TTPBarcodeDataClass(itembarcode, socname, itemname, itemoptionname, customerprice, printno);
		arrbarcode.push(v);
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4_35x15'){
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

			printTOSHIBAjewelleryMultiBarcode(arrbarcode);
		}

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, 0) == true) {
		if (confirm("선택 상품의 쥬얼리 바코드를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPjewelleryMultiBarcode(arrbarcode);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

String.prototype.trim = function() {
    return this.replace(/^\s+|\s+$/g, "");
};

function reg(page){
//	if(frm.itemid.value!=""){
//		if (!IsDouble(frm.itemid.value)){
//			alert("상품코드는 숫자만 가능합니다.");
//			frm.itemid.focus();
//			return;
//		}
//	}

	if(frm.prdcode.value!=""){
		if ( GetByteLength(frm.prdcode.value) < 10 ){
			alert("물류코드를 정확하게 입력하세요.");
			frm.prdcode.focus();
			return;
		}
	}

	frm.page.value=page;
	frm.action='';
	frm.target='';
	frm.method="get"
	frm.submit();
}

//쇼카드 출력
function paperBarcodePrint(onoffgubun) {
	var chk; var itembarcode; var itembarcodearr=''; var papername='';
	var found = false;
	var papername = frm.papername.value;

//	if (papername==''){
//		alert('인쇄 하실 쇼카드를 선택해 주세요.');
//		return;
//	}

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("cksel" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		itembarcodearr = itembarcodearr + document.getElementById("itembarcode_" + i).value.trim() + ','
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	alert("[필수]인쇄될 쇼카드가 뜨면, 마우스 오른쪽 버튼을 눌러 인쇄를 클릭해 주세요.\n\n상단에 있는 쇼카드출력설명서 대로, 설정후에 인쇄하셔야 간격이 정상적으로 인쇄됩니다.");

	frm.itembarcodearr.value=itembarcodearr;
	frm.action='/common/barcode/paperbarcodeprint_on_off_multi.asp';
	frm.target='_blank';
	frm.method="post"
	frm.submit();
}

function SelectCk(opt){
	$(document.frmList.cksel).prop('checked',opt.checked);
}

function CheckThis(tn){
	var cksel = $("#frmList #cksel"+tn);
	cksel.prop("checked", true);
}

function IndexSudongBarcodePrint(){
	var popwin = window.open('/common/barcode/sudongindexprint.asp?menupos=<%=menupos%>','IndexSudong','width=1024,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<OBJECT
	  id=iaxobject
	  classid="clsid:5D776FEA-8C6B-4C53-8EC3-3585FC040BDB"
	  codebase="https://scm.10x10.co.kr/common/cab/tenbarPrint.cab#version=1,0,0,29"
	  width=0
	  height=0
	  align=center
	  hspace=0
	  vspace=0
>
</OBJECT>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="onoffgubun" value="<%= onoffgubun %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="itembarcodearr" value="">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<% if (C_IS_Maker_Upche) then %>
			* 브랜드 : <%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
		<% else %>
			* 브랜드 : <%	drawSelectBoxDesignerWithName "makerid", makerid %>
		<% end if %>
		&nbsp;
		* 상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32" onKeyPress="if (event.keyCode == 13) 	reg('');">
		&nbsp;
		* 물류코드 :
		<input type="text" class="text" name="prdcode" value="<%= prdcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;
		* 상품코드 : <textarea rows="3" cols="10" name="itemid" id="itemid" ><%=replace(itemid,",",chr(10))%></textarea><!-- onKeyPress="if (event.keyCode == 13) 	reg('');" -->
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 범용바코드 :
		<input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">

		<span style="white-space:nowrap;">* 판매여부 : <% drawSelectBoxSellYN "sellyn", sellyn %></span>
     	&nbsp;
     	<span style="white-space:nowrap;">* 사용여부 : <% drawSelectBoxUsingYN "usingyn", usingyn %></span>
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<br>
<!-- #include virtual="/common/barcode/inc_setting_menu_barcodeprint.asp" -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<br><font color="red">리스트기준 :</font>
		<input type="radio" name="listgubun" value="ITEM" onClick="reg('');" <% if listgubun = "ITEM" then response.write " checked" %>>상품리스트
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->

<br>
<!-- #include virtual="/common/barcode/inc_button_barcodeprint.asp" -->
</form>

<form name="frmList" id="frmList" method="POST" tyle="margin:0px;">
<input type="hidden" name="barcodetype" value="">
<input type="hidden" name="currencychar" value="">
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
	<td>
		<input type="checkbox" name="ckall" onClick="SelectCk(this);" id="ckall">
	</td>
	<td>이미지</td>
	<td>물류바코드<br><font color=blue>[범용 바코드]</font></td>
	<td>브랜드</td>
	<td align="left">
		상품명<font color=blue>[옵션]</font>
		<% if (isforeignprint = "Y") then %>
			<br>해외&nbsp;상품명<font color=blue>[해외&nbsp;옵션]</font>
		<% end if %>
	</td>
	<td>
		소비자가
		<% if (isforeignprint = "Y") then %>
			<br>[해외&nbsp;가격]
		<% end if %>
	</td>
	<td>판매가</td>
	<td>할인가</td>
	<td>수량</td>
	<td>비고</td>
</tr>

<% if not(isdispconfirm) then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center">
			<font color="red"><strong>온라인 상품정보 조회 권한이 없습니다.</strong></font>
		</td>
	</tr>
<% elseif oproduct.FresultCount > 0 then %>
	<% for i=0 to oproduct.FresultCount-1 %>
	<input type="hidden" name="location_name" value="<%= replace(oproduct.FItemList(i).Flocationname,"""","") %>">
	<input type="hidden" id="makerid_<%= i %>" name="locationid" value="<%= oproduct.FItemList(i).Flocationid %>">
	<input type="hidden" id="socname_<%= i %>" name="socname" value="<%= replace(oproduct.FItemList(i).fsocname,"""","") %>">
	<input type="hidden" id="socname_kor_<%= i %>" name="socname_kor" value="<%= replace(oproduct.FItemList(i).fsocname_kor,"""","") %>">
	<input type="hidden" id="itemgubun_<%= i %>" name="itemgubun" value="<%= oproduct.FItemList(i).Fitemgubun %>">
	<input type="hidden" id="itemid_<%= i %>" name="itemid" value="<%= oproduct.FItemList(i).Fitemid %>">
	<input type="hidden" id="itemoption_<%= i %>" name="itemoption" value="<%= oproduct.FItemList(i).Fitemoption %>">
	<input type="hidden" id="itembarcode_<%= i %>" name="prdcode" value="<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>">
	<input type="hidden" id="publicbarcode_<%= i %>" name="generalbarcode" value="<%= replace(oproduct.FItemList(i).Fgeneralbarcode,"""","") %>">
	<input type="hidden" id="itemname_<%= i %>" name="prdname" value="<%= replace(oproduct.FItemList(i).Fprdname,"""","") %>">
	<input type="hidden" id="itemoptionname_<%= i %>" name="prdoptionname" value="<%= replace(oproduct.FItemList(i).Fitemoptionname,"""","") %>">
	<input type="hidden" id="customerprice_<%= i %>" name="customerprice" value="<%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %>">
	<input type="hidden" id="sellprice_<%= i %>" name="sellprice" value="<%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %>">
	<input type="hidden" id="itemname_foreign_<%= i %>" name="prdname_foreign" value="<%= replace(oproduct.FItemList(i).Flcitemname,"""","") %>">
	<input type="hidden" id="itemoptionname_foreign_<%= i %>" name="prdoptionname_foreign" value="<%= replace(oproduct.FItemList(i).Flcitemoptionname,"""","") %>">
	<input type="hidden" id="customerprice_foreign_<%= i %>" name="customerprice_foreign" value="<%= round(oproduct.FItemList(i).Flcprice,2) %>">
	<input type="hidden" id="saleprice_<%= i %>" name="saleprice" value="<%= FormatNumber(oproduct.FItemList(i).Fsaleprice,0) %>">
	<input type="hidden" id="couponprice_<%= i %>" name="couponprice" value="<%= FormatNumber(oproduct.FItemList(i).GetCouponAssignPrice,0) %>">
	<input type="hidden" id="saleyn_<%= i %>" name="saleyn" value="<% if oproduct.FItemList(i).Fsaleyn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).Fsaleprice then %>Y<% else %>N<% end if %>">
	<input type="hidden" id="couponyn_<%= i %>" name="couponyn" value="<% if oproduct.FItemList(i).FitemCouponYn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).GetCouponAssignPrice then %>Y<% else %>N<% end if %>">
	<input type="hidden" id="prtidx_<%= i %>" name="prtidx" value="<%= oproduct.FItemList(i).fprtidx %>">
	<input type="hidden" id="itemrackcode_<%= i %>" name="itemrackcode" value="<%= replace(oproduct.FItemList(i).fitemrackcode,"""","") %>">
	<input type="hidden" id="itemoptionrackcode_<%= i %>" name="itemoptionrackcode" value="">
	<input type="hidden" id="subitemrackcode_<%= i %>" name="subitemrackcode" value="<%= replace(oproduct.FItemList(i).fsubitemrackcode,"""","") %>">

	<tr align="center" bgcolor="#FFFFFF">
		<td><input type="checkbox" id="cksel<%= i %>" name="cksel" value="<%=i+1%>" onClick="AnCheckClick(this);"></td>
		<td><img src="<%= oproduct.FItemList(i).Fmainimageurl %>" width="50"></td>
		<td>
			<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>
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
			<% if (isforeignprint = "Y") then %>
				<br><%= AddSpace(oproduct.FItemList(i).Flcitemname) %>
				<% if oproduct.FItemList(i).Fitemoptionname <> "" then %>
					<font color=blue>[<%= AddSpace(oproduct.FItemList(i).Flcitemoptionname) %>]</font>
				<% end if %>
			<% end if %>
		</td>
		<td align="right">
			<%= currencychar %>&nbsp;<%= AddSpace(FormatNumber(oproduct.FItemList(i).Fcustomerprice,0)) %>

			<% if (isforeignprint = "Y") then %>
				<br><%= currencyChar %> <%= AddSpace(oproduct.FItemList(i).Flcprice) %>
			<% end if %>
		</td>
		<td align="right">
			<%= currencychar %>&nbsp;<%= AddSpace(FormatNumber(oproduct.FItemList(i).Fsellprice,0)) %>
		</td>
		<td align="right">
			<%
			'/할인 처리
			if oproduct.FItemList(i).Fsaleyn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).Fsaleprice then
			%>
				<font color='red'><%= currencychar %>&nbsp;<%= AddSpace(FormatNumber(oproduct.FItemList(i).Fsaleprice,0)) %></font>
			<%
			'/쿠폰 처리
			elseif oproduct.FItemList(i).FitemCouponYn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).GetCouponAssignPrice then
			%>
				<font color='red'><%= currencychar %>&nbsp;<%= AddSpace(FormatNumber(oproduct.FItemList(i).GetCouponAssignPrice,0)) %></font>
			<% end if %>
		</td>
		<td>
			<input type="text" id="printno_<%= i %>" name="fixedno" value="<%= oproduct.FItemList(i).Ffixedno %>" size="4" maxlength="4" onKeyPress="CheckThis(<%= i %>);" onFocus="this.select()" class="text">
		</td>
		<td><!--<input type="checkbox" name="IsSellPricePrint" checked>가격출력--></td>
	</tr>
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
		<td colspan="20" align="center">
			<% if not(isdispsql) then %>
				<font color="red"><strong>검색 조건(브랜드,상품명,물류코드,상품코드,범용바코드)을 입력 하셔야 검색이 됩니다.</strong></font>
			<% else %>
				[검색결과가 없습니다.]
			<% end if %>
		</td>
	</tr>
<% end if %>
</table>
</form>
<%
set oproduct = nothing

Function AddSpace(byval str)
	if ((str = "") or (IsNull(str))) then
		AddSpace = "&nbsp;"
	else
		AddSpace = str
	end if
End Function
%>