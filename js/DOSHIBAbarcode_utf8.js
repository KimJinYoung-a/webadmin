//####################################################
// Description :  도시바 바코드 JS
// History : 2016.11.24 한용민 생성
///////////////////// 이파일 수정시 밑에 파일도 모두 동일하게 같이 고쳐야 한다. ////////////////////////
// SCM : /js/DOSHIBAbarcode.js
//		 /js/DOSHIBAbarcode_utf8.js
// LOGICS : /js/DOSHIBAbarcode.js
//			/js/DOSHIBAbarcode_utf8.js
//####################################################

//////////////////////////////// 기본값 & 모듈 설치 //////////////////////////////
var TOSHIBA_PAPERWIDTH = 450;
var TOSHIBA_PAPERHEIGHT	= 220;
var TOSHIBA_PAPERMARGIN	= 3;
var TOSHIBA_HEIGHTOFFSET = 0;
var TOSHIBA_WIDTHOFFSET	= 0;
var TOSHIBA_isforeignprint = 'N';
var TOSHIBA_currencyChar = '￦';
var TOSHIBA_DOMAINNAME = 'www.10x10.co.kr';
var TOSHIBA_SHOWDOMAINYN = 'Y';
var TOSHIBA_SHOWPRICEYN = 'Y';
var TOSHIBA_SHOPBRANDYN = 'Y';
var TOSHIBA_BARCODETYPE = 'T';
var TOSHIBA_currencyunit = 'KRW';
var TOSHIBA_currencyunit_pos = 'KRW';

var TOSHIBA_brand_str = '품명';
var TOSHIBA_origin_str = '산지';
var TOSHIBA_material_str = '재질';
var TOSHIBA_standard_str = '규격';
var TOSHIBA_manufacturer_str = '제조사';
var TOSHIBA_address_str = '주소';
var TOSHIBA_import_str = '수입상';
var TOSHIBA_telephone_str = '전화';

// 도시바 모듈 설치
DrawReceiptPrintobj_TEC("TEC_DO3","TEC B-FV4");
//////////////////////////////// 기본값 & 모듈 설치 //////////////////////////////

// 해외바코드용
function BarcodeDataClass_foreign(barcode, makerid, itemname, itemoptionname, customerprice, printno, catename
, sourcearea, itemsource, itemsize, manufacturer, m_address1, m_address2, m_telephone, vimport, i_address1, i_address2, i_telephone) {
    this.barcode = barcode;
    this.makerid = makerid;
    this.itemname = itemname;
    this.itemoptionname = itemoptionname;
    this.customerprice = customerprice;
    this.printno = printno;
    this.catename = catename;
    this.sourcearea = sourcearea;
    this.itemsource = itemsource;
    this.itemsize = itemsize;
    this.manufacturer = manufacturer;
    this.m_address1 = m_address1;
    this.m_address2 = m_address2;
    this.m_telephone = m_telephone;
    this.vimport = vimport;
    this.i_address1 = i_address1;
    this.i_address2 = i_address2;
    this.i_telephone = i_telephone;
}

// 국내 바코드용
function BarcodeDataClass(barcode, makerid, itemname, itemoptionname, customerprice, printno, catename
, socname, socname_kor) {
    this.barcode = barcode;
    this.makerid = makerid;
    this.itemname = itemname;
    this.itemoptionname = itemoptionname;
    this.customerprice = customerprice;
    this.printno = printno;
    this.catename = catename;
    this.socname = socname;
    this.socname_kor = socname_kor;
}

// 국내 인덱스 바코드용
function BarcodeDataClass_index(barcode, makerid, itemname, itemoptionname, customerprice, printno, catename
, socname, socname_kor, brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode) {
    this.barcode = barcode;
    this.makerid = makerid;
    this.itemname = itemname;
    this.itemoptionname = itemoptionname;
    this.customerprice = customerprice;
    this.printno = printno;
    this.catename = catename;
    this.socname = socname;
    this.socname_kor = socname_kor;
    this.brandrackcode = brandrackcode;
    this.itemrackcode = itemrackcode;
    this.itemoptionname = itemoptionname;
    this.subitemrackcode = subitemrackcode;
}

// 수기 바코드용
function BarcodeDataClass_udong(msg,itemno,fontName) {
    this.msg = msg;
    this.itemno = itemno;
    this.fontName = fontName;
}

//해외 바코드 출력
function printTOSHIBAforeignBarcode(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var v; var tenBarcode; var pubBarcode;
	var catename; var sourcearea; var itemsource; var itemsize;
	var manufacturer; var m_address1; var m_address2; var m_telephone; var vimport; var i_address1; var i_address2;
	var i_telephone;

	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;

	for (var i = 0; i < arrObject.length; i++) {
		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		printno = v.printno;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		catename = v.catename;
		sourcearea = v.sourcearea;
		itemsource = v.itemsource;
		itemsize = v.itemsize;
		manufacturer = v.manufacturer;
		m_address1 = v.m_address1;
		m_address2 = v.m_address2;
		m_telephone = v.m_telephone;
		vimport = v.vimport;
		i_address1 = v.i_address1;
		i_address2 = v.i_address2;
		i_telephone = v.i_telephone;

		skipnotinserted = true;

		//TEC_DO3.SetPrinterCopies('TEC B-FV4', 2);
		//위에 수량만큼 출력되는게 먹지 안는듯;; 우선 루프 돌려서 뽑는다. 명령어 찾아서 방식을 바꿔야함
		for (var j = 0;j < printno; j++) {
			TEC_DO3.PrinterOpen();

			if ((itemname != "") && (printno*1 > 0)) {
				tenBarcode = '';
				pubBarcode = '';
				var tmpArr = barcode.split("_");
				if (tmpArr.length >= 0) {
					tenBarcode = tmpArr[0];
				}
				if (tmpArr.length >= 1) {
					pubBarcode = tmpArr[1];
				}

				//중국어 짜쯩이... 가끔씩 깨짐. 卡 이글짜도 깨짐.
				//Arial(다국적 폰트.윈도우 기본 폰트)
				//Microsoft JhengHei(번자체.대만) , Microsoft YaHei(간자체.중국)

				//품명. 品名(품명)
				TEC_DO3.PrintText(0,10 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_brand_str+' : ['+makerid+'] '+catename);

				//산지. 産地(산지)
				TEC_DO3.PrintText(0,25 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_origin_str+' : '+sourcearea);

				//재질. 材質(재질)
				TEC_DO3.PrintText(0,41 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_material_str+' : '+itemsource);

				//규격. 規格(규격)
				TEC_DO3.PrintText(0,57 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_standard_str+' : '+itemsize);

				//제조사. 委製商(위제상)
				TEC_DO3.PrintText(0,72 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_manufacturer_str+' : '+manufacturer);

				//전화. 電話(전화)
				TEC_DO3.PrintText(0,87 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_telephone_str+' : '+m_telephone);

				//주소. 地址(지지)
				TEC_DO3.PrintText(0,101 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_address_str+' : '+m_address1);
				TEC_DO3.PrintText(10,115 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, m_address2);

				//수입상. 進口商(진구상)
				TEC_DO3.PrintText(0,131 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_import_str+' : '+vimport);

				//전화. 電話(전화)
				TEC_DO3.PrintText(0,146 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_telephone_str+' : '+i_telephone);

				//주소. 地址(지지)
				TEC_DO3.PrintText(0,161 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_address_str+' : '+i_address1);
				TEC_DO3.PrintText(10,176 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, i_address2);

				TEC_DO3.PrintText(0,177 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 0, 0, '___________________________________');

				//가격
			    if (TOSHIBA_SHOWPRICEYN == "Y" || TOSHIBA_SHOWPRICEYN == "C" || TOSHIBA_SHOWPRICEYN == "R"){
			        TEC_DO3.PrintText(0,197 + TOSHIBA_HEIGHTOFFSET, "Arial", 25*F, 1, 0, TOSHIBA_currencyChar+' '+customerprice);
			    }else if (TOSHIBA_SHOWPRICEYN == "S"){
			    	TEC_DO3.PrintText(0,197 + TOSHIBA_HEIGHTOFFSET, "Arial", 25*F, 1, 0, customerprice);
			    }

				//브랜드
				if (TOSHIBA_SHOPBRANDYN == "Y"){
					TEC_DO3.PrintText(100,197 + TOSHIBA_HEIGHTOFFSET, "Arial", 21*F, 1, 0, makerid);
				}

				//상품명[옵션]
				if (itemoptionname == "") {
					TEC_DO3.PrintText(0,218 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 1, 0, itemname);
					//TEC_DO3.PrintText(0,234 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 1, 0, '[테스트테스트테스트테스트테스트테스트테스트]');
				} else {
					TEC_DO3.PrintText(0,218 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 1, 0, itemname);
					TEC_DO3.PrintText(0,234 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 1, 0, '[' + itemoptionname + ']');
				}

				if (TOSHIBA_BARCODETYPE == "A") {
					//옵션코드에 Z가 들어 있을경우
				    if (InStr(barcode, 'Z') >= 0) {
				    	TEC_DO3.PrintText(0, 254 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 25, 0, 0, tenBarcode);	//'1001739766Z011
				    	TEC_DO3.PrintText(0,281 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 1, 0, getTTPBarcodeString(tenBarcode));

					// 텐바이텐 일반 물류코드
				    } else {
				    	TEC_DO3.PrintText(0, 254 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMEAN128", 25, 0, 0, tenBarcode);
				    	TEC_DO3.PrintText(0,281 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 1, 0, getTTPBarcodeString(tenBarcode));
				    }
			    	TEC_DO3.PrintText(0, 302 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 25, 0, 0, pubBarcode);
			    	TEC_DO3.PrintText(0,329 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 1, 0, pubBarcode);
				}
			}

			TEC_DO3.PrinterClose();
		}
	}
}

//인덱스 바코드 출력
function printTOSHIBAMultiItemLabel(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno, brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode;
	var v;
	var tenBarcode, pubBarcode;

	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;

	for (var i = 0; i < arrObject.length; i++) {
		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		printno = v.printno;
		brandrackcode = v.brandrackcode;
		itemrackcode = v.itemrackcode;
		itemoptionrackcode = v.itemoptionrackcode;
		subitemrackcode = v.subitemrackcode;

		tenBarcode = "";
		pubBarcode = "";

		if (TOSHIBA_BARCODETYPE == "T") {
			tenBarcode = barcode;
		} else if (TOSHIBA_BARCODETYPE == "G") {
			pubBarcode = barcode;
		} else if (TOSHIBA_BARCODETYPE == "2") {
			var tmpArr = barcode.split("_");
			if (tmpArr.length != 2) {
				alert("잘못된 형식입니다.");
				return;
			}
			tenBarcode = tmpArr[0];
			pubBarcode = tmpArr[1];
		}

		skipnotinserted = true;

		//TEC_DO3.SetPrinterCopies('TEC B-FV4', 2);
		//위에 수량만큼 출력되는게 먹지 안는듯;; 우선 루프 돌려서 뽑는다. 명령어 찾아서 방식을 바꿔야함
		for (var j = 0;j < printno; j++) {
			TEC_DO3.PrinterOpen();

			if ((itemname != "") && (printno*1 > 0)) {
//				if (TOSHIBA_SHOWDOMAINYN == "Y") {
//					TEC_DO3.PrintText(190,20 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, TOSHIBA_DOMAINNAME);
//				}

				// 브랜드랙과 보조랙에 값이 있는데 브랜드랙과 보조랙 값이 틀린거
				if (brandrackcode!="" && subitemrackcode!="" && brandrackcode!=subitemrackcode){
					TEC_DO3.PrintText(40,20 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "보조랙 : " + subitemrackcode);

				}else if (brandrackcode!=""){
					TEC_DO3.PrintText(40,20 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "브랜드랙 : " + brandrackcode);
				}

				if (TOSHIBA_SHOPBRANDYN == "Y"){
					TEC_DO3.PrintText(40,55 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, makerid);
				}

				TEC_DO3.PrintText(40,90 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, itemname);
				if (itemoptionname != "") {
					TEC_DO3.PrintText(40,120 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, '[' + itemoptionname + ']');
				}

				TEC_DO3.PrintText(70,160 + TOSHIBA_HEIGHTOFFSET, "Arial", 130*F, 0, 0, getTTPBarcodeItemidString(tenBarcode));

				if ((TOSHIBA_BARCODETYPE == "T") || (TOSHIBA_BARCODETYPE == "2")) {
					//상품구분 + 상품코드
					TEC_DO3.PrintText(310,280 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, left(tenBarcode,2) + "-" + getTTPBarcodeItemidString(tenBarcode));
					//옵션
					TEC_DO3.PrintText(470,275 + TOSHIBA_HEIGHTOFFSET, "Arial", 50*F, 0, 0, "-" + right(tenBarcode,4));

					if (pubBarcode != "") {
						TEC_DO3.PrintText(360,330 + TOSHIBA_HEIGHTOFFSET, "Arial", 35*F, 0, 0, pubBarcode);
					}
				}

				TEC_DO3.PrintText(40, 335 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMEAN128", 30, 0, 0, tenBarcode);
			}

			TEC_DO3.PrinterClose();
		}
	}
}

//상품 바코드 출력
function printTOSHIBAMultiBarcode(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var v; var tenBarcode; var pubBarcode; var socname; var socname_kor;

	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;

	for (var i = 0; i < arrObject.length; i++) {
		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		printno = v.printno;
		socname = v.socname;
		socname_kor = v.socname_kor;

		//표시 상품명(국내,해외)
		if (TOSHIBA_isforeignprint == "N") {
			itemname = v.itemname;
			itemoptionname = v.itemoptionname;
			customerprice = v.customerprice;
		} else {
			itemname = v.itemname;
			itemoptionname = v.itemoptionname;
			customerprice = v.customerprice;
		}

		skipnotinserted = true;

		//TEC_DO3.SetPrinterCopies('TEC B-FV4', 2);
		//위에 수량만큼 출력되는게 먹지 안는듯;; 우선 루프 돌려서 뽑는다. 명령어 찾아서 방식을 바꿔야함
		for (var j = 0;j < printno; j++) {
			TEC_DO3.PrinterOpen();

			if ((itemname != "") && (printno*1 > 0)) {
				tenBarcode = '';
				pubBarcode = '';
				var tmpArr = barcode.split("_");
				if (tmpArr.length >= 0) {
					tenBarcode = tmpArr[0];
				}
				if (tmpArr.length >= 1) {
					pubBarcode = tmpArr[1];
				}

//				if (TOSHIBA_SHOWDOMAINYN == "Y") {
//					if (TOSHIBA_currencyChar!='￥'){
//						TEC_DO3.PrintText(75,0 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, TOSHIBA_DOMAINNAME);
//					}
//				}

				// 브랜드
				if (TOSHIBA_SHOPBRANDYN == "Y"){
					TEC_DO3.PrintText(10,10 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 1, 0, socname_kor + '  ' + socname);
				}

				//상품명[옵션]
				if (itemoptionname == "") {
					TEC_DO3.PrintText(10,35 + TOSHIBA_HEIGHTOFFSET, "Arial", 16*F, 0, 0, itemname);
					//TEC_DO3.PrintText(10,55 + TOSHIBA_HEIGHTOFFSET, "Arial", 16*F, 0, 0, '[테스트테스트테스트테스트]');
				} else {
					TEC_DO3.PrintText(10,35 + TOSHIBA_HEIGHTOFFSET, "Arial", 16*F, 0, 0, itemname);
					TEC_DO3.PrintText(10,55 + TOSHIBA_HEIGHTOFFSET, "Arial", 16*F, 0, 0, '[' + itemoptionname + ']');
				}

			    if (TOSHIBA_SHOWPRICEYN == "Y" || TOSHIBA_SHOWPRICEYN == "C" || TOSHIBA_SHOWPRICEYN == "R"){
			        TEC_DO3.PrintText(220,80 + TOSHIBA_HEIGHTOFFSET, "Arial", 22*F, 1, 0, TOSHIBA_currencyChar+' '+customerprice);
			    }else if (TOSHIBA_SHOWPRICEYN == "S"){
			    	TEC_DO3.PrintText(220,80 + TOSHIBA_HEIGHTOFFSET, "Arial", 22*F, 1, 0, customerprice);
			    }

				TEC_DO3.PrintText(10,85 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 0, 0, '___________________________________');

				//옵션코드에 Z가 들어 있을경우
			    if ((TOSHIBA_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
			    	TEC_DO3.PrintText(10, 110 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 30, 0, 0, tenBarcode);
			    	TEC_DO3.PrintText(10,140 + TOSHIBA_HEIGHTOFFSET, "Arial", 18*F, 0, 0, getTTPBarcodeString(tenBarcode));
			    	
				// 텐바이텐 일반 물류코드
			    } else if (TOSHIBA_BARCODETYPE == "T") {
			    	TEC_DO3.PrintText(10, 110 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMEAN128", 30, 0, 0, tenBarcode);	//'1001739766Z011'
			    	TEC_DO3.PrintText(45,140 + TOSHIBA_HEIGHTOFFSET, "Arial", 18*F, 0, 0, getTTPBarcodeString(tenBarcode));

				// 범용바코드
			    } else {
			    	TEC_DO3.PrintText(10, 110 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 30, 0, 0, pubBarcode);
			    	TEC_DO3.PrintText(10,140 + TOSHIBA_HEIGHTOFFSET, "Arial", 18*F, 0, 0, pubBarcode);
			    }
			}

			TEC_DO3.PrinterClose();
		}
	}
}

//쥬얼리 바코드 출력
function printTOSHIBAjewelleryMultiBarcode(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var v; var tenBarcode; var pubBarcode;

	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;

	for (var i = 0; i < arrObject.length; i++) {
		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		printno = v.printno;

		//표시 상품명(국내,해외)
		if (TOSHIBA_isforeignprint == "N") {
			itemname = v.itemname;
			itemoptionname = v.itemoptionname;
			customerprice = v.customerprice;
		} else {
			itemname = v.itemname;
			itemoptionname = v.itemoptionname;
			customerprice = v.customerprice;
		}

		skipnotinserted = true;

		//TEC_DO3.SetPrinterCopies('TEC B-FV4', 2);
		//위에 수량만큼 출력되는게 먹지 안는듯;; 우선 루프 돌려서 뽑는다. 명령어 찾아서 방식을 바꿔야함
		for (var j = 0;j < printno; j++) {
			TEC_DO3.PrinterOpen();

			if ((itemname != "") && (printno*1 > 0)) {
				tenBarcode = '';
				pubBarcode = '';
				if (TOSHIBA_BARCODETYPE == "T") {
					var tmpArr = barcode.split("_");
					if (tmpArr.length != 2) {
						alert("잘못된 형식입니다.");
						return;
					}
					tenBarcode = tmpArr[0];
					pubBarcode = tmpArr[1];
				} else if (TOSHIBA_BARCODETYPE == "G") {
					var tmpArr = barcode.split("_");
					if (tmpArr.length != 2) {
						alert("잘못된 형식입니다.");
						return;
					}
					tenBarcode = tmpArr[0];
					pubBarcode = tmpArr[1];
				} else if (TOSHIBA_BARCODETYPE == "2") {
					var tmpArr = barcode.split("_");
					if (tmpArr.length != 2) {
						alert("잘못된 형식입니다.");
						return;
					}
					tenBarcode = tmpArr[0];
					pubBarcode = tmpArr[1];
				}

				if (TOSHIBA_SHOPBRANDYN == "Y"){
					TEC_DO3.PrintText(5,5 + TOSHIBA_HEIGHTOFFSET, "10X10", 18*F, 1, 0, makerid);
				}

				//상품명[옵션]
				if (itemoptionname == "") {
					TEC_DO3.PrintText(5,23 + TOSHIBA_HEIGHTOFFSET, "10X10", 13*F, 0, 0, itemname);
				} else {
					TEC_DO3.PrintText(5,23 + TOSHIBA_HEIGHTOFFSET, "10X10", 13*F, 0, 0, itemname);
					TEC_DO3.PrintText(5,36 + TOSHIBA_HEIGHTOFFSET, "10X10", 13*F, 0, 0, '[' + itemoptionname + ']');
				}

			    if (TOSHIBA_SHOWPRICEYN == "Y" || TOSHIBA_SHOWPRICEYN == "C" || TOSHIBA_SHOWPRICEYN == "R"){
			        TEC_DO3.PrintText(165,50 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 1, 0, TOSHIBA_currencyChar+' '+customerprice);
			    }else if (TOSHIBA_SHOWPRICEYN == "S"){
			    	TEC_DO3.PrintText(165,50 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 1, 0, customerprice);
			    }

				TEC_DO3.PrintText(5,51 + TOSHIBA_HEIGHTOFFSET, "10X10", 20*F, 0, 0, '_________________________');

				//옵션코드에 Z가 들어 있을경우
			    if ((TOSHIBA_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
			    	TEC_DO3.PrintText(20, 74 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 20, 0, 0, tenBarcode);
			    	TEC_DO3.PrintText(20,95 + TOSHIBA_HEIGHTOFFSET, "10X10", 15*F, 0, 0, getTTPBarcodeString(tenBarcode));

				// 텐바이텐 일반 물류코드
			    } else if (TOSHIBA_BARCODETYPE == "T") {
			    	TEC_DO3.PrintText(20, 74 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 20, 0, 0, tenBarcode);	//'1001739766Z011'
			    	TEC_DO3.PrintText(20,95 + TOSHIBA_HEIGHTOFFSET, "10X10", 15*F, 0, 0, getTTPBarcodeString(tenBarcode));

				// 범용바코드
			    } else {
			    	TEC_DO3.PrintText(20, 74 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 20, 0, 0, pubBarcode);
			    	TEC_DO3.PrintText(20,95 + TOSHIBA_HEIGHTOFFSET, "10X10", 15*F, 0, 0, pubBarcode);
			    }
			}

			TEC_DO3.PrinterClose();
		}
	}
}

//DAS 바코드 출력
function printTOSHIBADasMultiBarcode(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	var v;
	var dasindex, itembarcode, baljubarcode = "";

	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;

	for (var i = 0; i < arrObject.length; i++) {
		v = arrObject[i];

		if (baljubarcode != v.baljubarcode) {
			baljubarcode = v.baljubarcode;

			TEC_DO3.PrinterOpen();

			TEC_DO3.PrintText(30,0 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, v.titleStr);
			TEC_DO3.PrintText(40,35 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "[상단] 상품 출고지시바코드");
			TEC_DO3.PrintText(60, 70 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMEAN128", 35, 0, 0, v.baljubarcode);
			TEC_DO3.PrintText(70,120 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, v.baljubarcode);

			TEC_DO3.PrinterClose();
		}

		v.dasindexarr = v.dasindexarr.split("|");
		if (v.danpumynarr) {
			v.danpumynarr = v.danpumynarr.split("|");
		}
		for (var j = 0; j < v.dasindexarr.length; j++) {
			dasindex = v.dasindexarr[j].replace(",", "번:");
			itembarcode = v.itembarcode.substring(0, 2) + "-" + v.itembarcode.substring(2, (v.itembarcode.length - 4)) + "-" + v.itembarcode.substring((v.itembarcode.length - 4));

			if (v.dasindexarr.length >= 5) {
				dasindex = "다수:" + v.totalItemNo;
			}

			TEC_DO3.PrinterOpen();

			TEC_DO3.PrintText(30,0 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, v.titleStr);
			TEC_DO3.PrintText(10,55 + TOSHIBA_HEIGHTOFFSET, "Arial", 35*F, 0, 0, "R:"+v.itemrackcode);
			TEC_DO3.PrintText(10,100 + TOSHIBA_HEIGHTOFFSET, "Arial", 35*F, 0, 0, dasindex);
			TEC_DO3.PrintText(120,35 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, v.makerid);
			TEC_DO3.PrintText(120,65 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, itembarcode);
			TEC_DO3.PrintText(120,95 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, v.publicbarcode);

			if (v.dasindexarr.length < 5 && dasindex.split(":").length == 2 && dasindex.split(":")[1] > 1) {
				TEC_DO3.PrintText(200,125 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, dasindex.split(":")[1] + "개");
			} else {
				TEC_DO3.PrintText(200,125 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "--");
			}

			if (v.danpumynarr) {
				if (v.danpumynarr.length < 5 && v.danpumynarr[j].split(",").length == 2 && v.danpumynarr[j].split(",")[1] === "Y") {
					TEC_DO3.PrintText(120,125 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "단품");
				}
			} else {
				TEC_DO3.PrintText(120,125 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "--");
			}

			TEC_DO3.PrinterClose();

			if (v.dasindexarr.length >= 5) {
				break;
			}
		}
	}
}

//innerbox 인덱스 출력
function printTOSHIBAinnerboxLabel(shopid, shopname, packingdate, innerboxno, innerboxweight, prdcode, prdbarcode){
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;
	TEC_DO3.PrinterOpen();

	if (TOSHIBA_SHOWDOMAINYN == "Y") {
		TEC_DO3.PrintText(30,10 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, TOSHIBA_DOMAINNAME);
	}
	TEC_DO3.PrintText(30,55 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, 'SHOPID : ' + shopid);
	TEC_DO3.PrintText(30,90 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, shopname);
	TEC_DO3.PrintText(30,130 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, 'DATE : ' + packingdate);
	TEC_DO3.PrintText(30,170 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, 'INNER BOX NO. : ' + innerboxno);
	TEC_DO3.PrintText(30,210 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, 'INNER BOX WEIGHT : ' + innerboxweight + ' KG');
	TEC_DO3.PrintText(160, 280 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMEAN128", 40, 0, 0, prdbarcode);
	TEC_DO3.PrintText(30,345 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, prdcode);

	TEC_DO3.PrinterClose();
}

//cartonbox 인덱스 출력
function printTOSHIBAcartonboxLabel(shopid, shopname, packingdate, cartonboxno, cartonboxweight, prdcode, prdbarcode){
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;
	TEC_DO3.PrinterOpen();

	if (TOSHIBA_SHOWDOMAINYN == "Y") {
		TEC_DO3.PrintText(30,10 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, TOSHIBA_DOMAINNAME);
	}
	TEC_DO3.PrintText(30,55 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, 'SHOPID : ' + shopid);
	TEC_DO3.PrintText(30,90 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, shopname);
	TEC_DO3.PrintText(30,130 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, 'DATE : ' + packingdate);
	TEC_DO3.PrintText(30,170 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, 'CARTON BOX NO. : ' + cartonboxno);
	TEC_DO3.PrintText(30,210 + TOSHIBA_HEIGHTOFFSET, "Arial", 40*F, 0, 0, 'CARTON BOX WEIGHT : ' + cartonboxweight + ' KG');
	TEC_DO3.PrintText(160, 280 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMEAN128", 40, 0, 0, prdbarcode);
	TEC_DO3.PrintText(30,345 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, prdcode);

	TEC_DO3.PrinterClose();
}

// 수기 인덱스 출력
function printTOSHIBAMultiIndexSudongLabel(msg,itemno,fontName) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	if (itemno=='' || itemno==0) itemno=1;
	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;
	
	var skipnotinserted = false;
	var lines = msg.split("\n");

	for (var j = 0;j < itemno; j++) {
    	TEC_DO3.PrinterOpen();

		for (var i = 0; i < lines.length; i++) {
			//iTTPBar.AXwindowsfont(20, 10 + TTP_HEIGHTOFFSET + (35 * i), 30, 0, 0, 0, fontName, lines[i].replace("\r", ""));
			TEC_DO3.PrintText(20,20 + TOSHIBA_HEIGHTOFFSET + (35 * i), fontName, 30*F, 0, 0, lines[i].replace("\r", ""));	
		}

		TEC_DO3.PrinterClose();
	}
}

// 수기 랙번호 인덱스 출력
function printTOSHIBARackIndexSudongLabel(msg,itemno,fontName, barcodeprintyn) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	if (fontName=='') fontName='10X10';
	if (itemno=='' || itemno==0) itemno=1;
	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;
	
	var skipnotinserted = false;
	var lines = msg.split("\n");
	var linemsg = "";

	for (var i = 0; i < lines.length; i++) {
		for (var j = 0;j < itemno; j++) {
			linemsg = lines[i].replace("\r", "").ltrim().rtrim();
			TEC_DO3.PrinterOpen();

			if (linemsg.length < 5){
				TEC_DO3.PrintText(10,70 + TOSHIBA_HEIGHTOFFSET, fontName, 200*F, 0, 0, linemsg);
				if (barcodeprintyn!=''){
					TEC_DO3.PrintText(120, 300 + TOSHIBA_HEIGHTOFFSET, "TEC-BarFont Code128A", 50, 0, 0, linemsg);
				}
			}else if (linemsg.length < 8){
				TEC_DO3.PrintText(10,70 + TOSHIBA_HEIGHTOFFSET, fontName, 120*F, 0, 0, linemsg);
				if (barcodeprintyn!=''){
					TEC_DO3.PrintText(120, 300 + TOSHIBA_HEIGHTOFFSET, "TEC-BarFont Code128A", 50, 0, 0, linemsg);
				}
			}else{
				TEC_DO3.PrintText(10,70 + TOSHIBA_HEIGHTOFFSET, fontName, 100*F, 0, 0, linemsg);
				if (barcodeprintyn!=''){
					TEC_DO3.PrintText(120, 300 + TOSHIBA_HEIGHTOFFSET, "TEC-BarFont Code128A", 50, 0, 0, linemsg);
				}
			}

			TEC_DO3.PrinterClose();
		}
	}
}

// 랙번호 인덱스 출력
function printTOSHIBAMultiRackIndexLabel(arrObject) {
	var v;
	var msg, itemno, fontName;

	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 드라이버를 설치해 주세요");
		return;
	}

	if (itemno=='' || itemno==0) itemno=1;
	var X = 1;
	var Y = 1;
	var F = 1;

	TEC_DO3.SetPaper(TOSHIBA_PAPERWIDTH,TOSHIBA_PAPERHEIGHT);
	TEC_DO3.OffsetX = TOSHIBA_WIDTHOFFSET;
	TEC_DO3.OffsetY = TOSHIBA_HEIGHTOFFSET;
	
	var skipnotinserted = false;
	var linemsg = "";

	for (var i = 0; i < arrObject.length; i++) {
		v = arrObject[i];

		msg = v.msg;
		itemno = v.itemno;
		fontName = v.fontName;
		if (fontName=='') fontName='10X10';

		for (var j = 0;j < itemno; j++) {
			linemsg = msg.replace("\r", "").ltrim().rtrim();
			TEC_DO3.PrinterOpen();

			if (linemsg.length < 5){
				TEC_DO3.PrintText(10,70 + TOSHIBA_HEIGHTOFFSET, fontName, 200*F, 0, 0, linemsg);
				TEC_DO3.PrintText(120, 300 + TOSHIBA_HEIGHTOFFSET, "TEC-BarFont Code128A", 50, 0, 0, linemsg);
			}else if (linemsg.length < 8){
				TEC_DO3.PrintText(10,70 + TOSHIBA_HEIGHTOFFSET, fontName, 120*F, 0, 0, linemsg);
				TEC_DO3.PrintText(120, 300 + TOSHIBA_HEIGHTOFFSET, "TEC-BarFont Code128A", 50, 0, 0, linemsg);
			}else{
				TEC_DO3.PrintText(10,70 + TOSHIBA_HEIGHTOFFSET, fontName, 100*F, 0, 0, linemsg);
				TEC_DO3.PrintText(120, 300 + TOSHIBA_HEIGHTOFFSET, "TEC-BarFont Code128A", 50, 0, 0, linemsg);
			}

			TEC_DO3.PrinterClose();
		}
	}
}
