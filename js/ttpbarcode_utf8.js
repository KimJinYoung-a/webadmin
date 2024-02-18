//####################################################
// Description :  TTP-243 바코드 JS
// History : 이상구 생성
//			 2016.11.24 한용민 수정
///////////////////// 이파일 수정시 밑에 파일도 모두 동일하게 같이 고쳐야 한다. ////////////////////////
// SCM : /js/ttpbarcode.js
//		 /js/ttpbarcode_utf8.js
// LOGICS : /js/ttpbarcode.js
//			/js/ttpbarcode_utf8.js
//####################################################

/*

<script language="JavaScript" src="/js/ttpbarcode.js"></script>
<script language='javascript'>

// 한상품 바코드 출력
// <input type="button" class="button" value="출력" onClick="BarcodePrint('102140800012', '122kcal', 'roll (pencil case)', 'carrot(orange)', '10,000', 5)">
function BarcodePrint(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
	// /js/barcode.js 참조
	if (initTTPprinter("TTP-243_45x22", "T", "Y", "www.10x10.co.kr", "Y", "￦", "Y", 3, 0) != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.');
		return;
	}

	if (printno*1 < 1) {
		alert("수량이 0 입니다.");
		return;
	}

	printTTPOneBarcode(barcode, makerid, itemname, itemoptionname, customerprice, printno);
}

// 여러상품 바코드 출력
function BarcodePrintSelected() {
	var frmdetail = document.frmdetail;
	var arr = new Array();
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;

	for (var i = 0; i < frmdetail.chk.length; i++) {
		if (frmdetail.chk[i].type == "checkbox") {
			if (frmdetail.chk[i].checked) {
				barcode			= frmdetail.itembarcode[i].value;

				makerid			= frmdetail.makerid[i].value;
				itemname		= frmdetail.itemname[i].value;
				itemoptionname	= frmdetail.itemoptionname[i].value;
				customerprice	= frmdetail.customerprice[i].value;
				printno			= frmdetail.checkitemno[i].value;

				var v = new TTPBarcodeDataClass(barcode, makerid, itemname, itemoptionname, customerprice, printno);
				arr.push(v);
			}
		}
	}

	if (arr.length < 1) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	// /js/barcode.js 참조
	if (initTTPprinter("TTP-243_45x22", "G", "Y", "www.10x10.co.kr", "Y", "￦", "Y", 3, 0) != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[4]');
		return;
	}

	printTTPMultiBarcode(arr);
}

</script>

*/

function checkTTPprinterExist() {
    //alert(iTTPBar);
    if (!iTTPBar.AXopenport(TTP_PRINTERTYPE)) {
        alert("a" + TTP_PRINTERTYPE);
        return false;
    }

    return true;
}

function InStr(str, substr, start) {
	var oStr = new String(str);
	return oStr.indexOf(substr,start);
}

function TTPBarcodeDataClass(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
    this.barcode = barcode;
    this.makerid = makerid;
    this.itemname = itemname;
    this.itemoptionname = itemoptionname;
    this.customerprice = customerprice;
    this.printno = printno;
}

function TTPEquipBarcodeDataClass(equip_code, AccountGubunName, equip_name, buy_date) {
    this.equip_code = equip_code;
    this.AccountGubunName = AccountGubunName;
    this.equip_name = equip_name;
    this.buy_date = buy_date;
}

// ============================================================================
var TTP_INITIALIZED = false			// true or false
var TTP_TTPTYPE						// TTP-243_45x22
var TTP_PRINTERTYPE					// TTP-243
var TTP_BARCODETYPE					// T or G(텐바이텐 바코드 or 범용바코드)
var TTP_SHOWDOMAINYN = 'Y'			// y or n
var TTP_DOMAINNAME					// www.10x10.co.kr
var TTP_SHOWPRICEYN					// y or n
var TTP_CURRENCYCHAR				// ￦(\ 역슬래시 아님) or $ or ￥
var TTP_SHOPBRANDYN					// y or n
var TTP_PAPERWIDTH					// 45
var TTP_PAPERHEIGHT					// 22
var TTP_PAPERMARGIN					// 3
var TTP_HEIGHTOFFSET				// 0
var TTP_isforeignprint = 'N';
var TTP_currencyunit = 'KRW';
var TTP_currencyunit_pos = 'KRW';

var TTP_brand_str = '품명';
var TTP_origin_str = '산지';
var TTP_material_str = '재질';
var TTP_standard_str = '규격';
var TTP_manufacturer_str = '제조사';
var TTP_address_str = '주소';
var TTP_import_str = '수입상';
var TTP_telephone_str = '전화';

// ============================================================================

//해외 바코드 출력
function printTTPMultiforeignBarcode(arrObject) {
	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var v; var tenBarcode; var pubBarcode;
	var catename; var sourcearea; var itemsource; var itemsize;
	var manufacturer; var m_address1; var m_address2; var m_telephone; var vimport; var i_address1; var i_address2;
	var i_telephone;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		printno = v.printno;
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
			iTTPBar.AXwindowsfont(10, 10 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_brand_str+' : ['+makerid+'] '+catename);

			//산지. 産地(산지)
			iTTPBar.AXwindowsfont(10, 25 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_origin_str+' : '+sourcearea);

			//재질. 材質(재질)
			iTTPBar.AXwindowsfont(10, 41 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_material_str+' : '+itemsource);

			//규격. 規格(규격)
			iTTPBar.AXwindowsfont(10, 57 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_standard_str+' : '+itemsize);

			//제조사. 委製商(위제상)
			iTTPBar.AXwindowsfont(10, 72 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_manufacturer_str+' : '+manufacturer);

			//전화. 電話(전화)
			iTTPBar.AXwindowsfont(10, 87 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_telephone_str+' : '+m_telephone);

			//주소. 地址(지지)
			iTTPBar.AXwindowsfont(10, 101 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_address_str+' : '+m_address1);
			iTTPBar.AXwindowsfont(20, 115 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', m_address2);

			//수입상. 進口商(진구상)
			iTTPBar.AXwindowsfont(10, 131 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_import_str+' : '+vimport);

			//전화. 電話(전화)
			iTTPBar.AXwindowsfont(10, 146 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_telephone_str+' : '+i_telephone);

			//주소. 地址(지지)
			iTTPBar.AXwindowsfont(10, 161 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', TTP_address_str+' : '+i_address1);
			iTTPBar.AXwindowsfont(20, 176 + TTP_HEIGHTOFFSET, 15, 0, 0, 0, 'Arial', i_address2);

			iTTPBar.AXwindowsfont(10, 177 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', '____________________________________');

			//가격
		    if (TTP_SHOWPRICEYN == "Y" || TTP_SHOWPRICEYN == "C" || TTP_SHOWPRICEYN == "R"){
		        iTTPBar.AXwindowsfont(10, 197 + TTP_HEIGHTOFFSET, 25, 0, 2, 0, 'Arial', TTP_CURRENCYCHAR + ' ' + customerprice);
		    }else if (TTP_SHOWPRICEYN == "S"){
		    	iTTPBar.AXwindowsfont(10, 197 + TTP_HEIGHTOFFSET, 25, 0, 2, 0, 'Arial', customerprice);
		    }

		    if (TTP_SHOPBRANDYN == "Y"){
				iTTPBar.AXwindowsfont(100, 197 + TTP_HEIGHTOFFSET, 21, 0, 2, 0, 'Arial', makerid);
		    }

			//상품명[옵션]
			if (itemoptionname == "") {
				iTTPBar.AXwindowsfont(10, 218 + TTP_HEIGHTOFFSET, 15, 0, 2, 0, 'Arial', itemname);
				//iTTPBar.AXwindowsfont(10, 234 + TTP_HEIGHTOFFSET, 15, 0, 2, 0, 'Arial', '[테스트테스트테스트테스트테스트테스트테스트]');
			} else {
				iTTPBar.AXwindowsfont(10, 218 + TTP_HEIGHTOFFSET, 15, 0, 2, 0, 'Arial', itemname);
				iTTPBar.AXwindowsfont(10, 234 + TTP_HEIGHTOFFSET, 15, 0, 2, 0, 'Arial', '[' + itemoptionname + ']');
			}

			if (TTP_BARCODETYPE == "A") {
				//옵션코드에 Z가 들어 있을경우
			    if (InStr(barcode, 'Z') >= 0) {
			    	iTTPBar.AXbarcode(30, 254 + TTP_HEIGHTOFFSET,'128','25','0','0','2','4',tenBarcode);		//'1001739766Z011
			    	iTTPBar.AXwindowsfont(30,281 + TTP_HEIGHTOFFSET,20,0,2,0,'Arial', getTTPBarcodeString(tenBarcode) );

				// 텐바이텐 일반 물류코드
			    } else {
			    	iTTPBar.AXbarcode('30', 254 + TTP_HEIGHTOFFSET,'128','25','0','0','2','4',tenBarcode);
			    	iTTPBar.AXwindowsfont(30,281 + TTP_HEIGHTOFFSET,20,0,2,0,'Arial', getTTPBarcodeString(tenBarcode) );
			    }

				// 범용바코드
				iTTPBar.AXbarcode('30', 302 + TTP_HEIGHTOFFSET,'128','25','0','0','2','4',pubBarcode);
		    	iTTPBar.AXwindowsfont(30,329 + TTP_HEIGHTOFFSET,20,0,2,0,'Arial', pubBarcode);
			}

		    // printno 장 프린트
		    iTTPBar.AXprintlabel('1', printno*1);

		    iTTPBar.AXformfeed();
		}

	}

	iTTPBar.AXcloseport();
}

//innerbox 인덱스 출력		//2016.12.14 한용민 생성
function printTTPinnerboxLabel(shopid, shopname, packingdate, innerboxno, innerboxweight, prdcode, prdbarcode){
    if (!iTTPBar.AXopenport('TTP-243')){
        alert("TSC TTP-243 드라이버를 설치해 주세요");
		return;
    }
    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH,TTP_PAPERHEIGHT,'2','10','0',TTP_PAPERMARGIN,'0');

	if (TTP_SHOWDOMAINYN == "Y") {
		iTTPBar.AXwindowsfont(30,0 + TTP_HEIGHTOFFSET,40,0,2,1,'Arial',TTP_DOMAINNAME);
	}
	iTTPBar.AXwindowsfont(30,55 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial','SHOPID : ' + shopid);
	iTTPBar.AXwindowsfont(30,90 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', shopname);
	iTTPBar.AXwindowsfont(30,130 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial','DATE : ' + packingdate);
	iTTPBar.AXwindowsfont(30,170 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial','INNER BOX NO. : ' + innerboxno);
	iTTPBar.AXwindowsfont(30,210 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial','INNER BOX WEIGHT : ' + innerboxweight + ' KG');
	iTTPBar.AXbarcode('160',280 + TTP_HEIGHTOFFSET,'128','40','0','0','2','4',prdbarcode);
	iTTPBar.AXwindowsfont(30,345 + TTP_HEIGHTOFFSET,30,0,0,0,'Arial',prdcode);
	iTTPBar.AXprintlabel('1','1');
	iTTPBar.AXcloseport();
}

//cartonbox 인덱스 출력		//2016.12.14 한용민 생성
function printTTPcartonboxLabel(shopid, shopname, packingdate, cartonboxno, cartonboxweight, prdcode, prdbarcode){
    if (!iTTPBar.AXopenport('TTP-243')){
        alert("TSC TTP-243 드라이버를 설치해 주세요");
		return;
    }

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH,TTP_PAPERHEIGHT,'2','10','0',TTP_PAPERMARGIN,'0');

	if (TTP_SHOWDOMAINYN == "Y") {
		iTTPBar.AXwindowsfont(30,0 + TTP_HEIGHTOFFSET,40,0,2,1,'Arial',TTP_DOMAINNAME);
	}
	iTTPBar.AXwindowsfont(30,55 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial','SHOPID : ' + shopid);
	iTTPBar.AXwindowsfont(30,90 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial',shopname);
	iTTPBar.AXwindowsfont(30,130 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial','DATE : ' + packingdate);
	iTTPBar.AXwindowsfont(30,170 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial','CARTON BOX NO. : ' + cartonboxno);
	iTTPBar.AXwindowsfont(30,210 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial','CARTON BOX WEIGHT : ' + cartonboxweight + ' KG' );
	iTTPBar.AXbarcode('160',280 + TTP_HEIGHTOFFSET,'128','40','0','0','2','4',prdbarcode);
	iTTPBar.AXwindowsfont(30,345 + TTP_HEIGHTOFFSET,30,0,0,0,'Arial',prdcode);

	iTTPBar.AXprintlabel('1','1');
	iTTPBar.AXcloseport();
}

function initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) {
	var s1, s2;

	TTP_INITIALIZED = false;

	if ( (ttptype != "TTP-243_45x22") && (ttptype != "TTP-243_35x15") && (ttptype != "TTP-243_80x50") && (ttptype != "TTP-index243_80x50") && (ttptype != "TTP-243_45x45") ) {
		alert('지원하지 않는 형식입니다.(TTP-243_45x22, TTP-243_35x15, TTP-243_45x45, TTP-243_80x50, TTP-index243_80x50 만 지원)');
		return false;
	}

	s1 = ttptype.split("_");
	s2 = s1[1].split("x");
	TTP_TTPTYPE			= ttptype;
	TTP_PRINTERTYPE		= s1[0];
	TTP_PAPERWIDTH		= s2[0]*1;
	TTP_PAPERHEIGHT		= s2[1]*1;

	TTP_BARCODETYPE		= barcodetype;
	TTP_SHOWDOMAINYN	= showdomainyn;
	TTP_DOMAINNAME		= domainname;
	TTP_SHOWPRICEYN		= showpriceyn;
	TTP_CURRENCYCHAR	= currencychar;
	TTP_SHOPBRANDYN		= shopbrandyn;
	TTP_PAPERMARGIN		= papermargin;
	TTP_HEIGHTOFFSET	= heightoffset;

	if (checkTTPprinterExist() != true) {
        return false;
	}

	TTP_INITIALIZED 	= true;

	return true;
}

//상품 바코드 출력
function printTTPOneBarcode(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
	var skipnotinserted = false;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[1]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
	iTTPBar.AXclearbuffer();

	skipnotinserted = true;

	if ((itemname != "") && printno*1 > 0) {
		if (TTP_SHOWDOMAINYN == "Y") {
			iTTPBar.AXwindowsfont(75, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', TTP_DOMAINNAME);
		}

		//상품명[옵션]
		if (itemoptionname == "") {
			iTTPBar.AXwindowsfont(20, 40 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', itemname);
		} else {
			iTTPBar.AXwindowsfont(20, 40 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', itemname);
			iTTPBar.AXwindowsfont(20, 65 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', '[' + itemoptionname + ']');
		}

	    if (TTP_SHOPBRANDYN == "Y"){
			iTTPBar.AXwindowsfont(20, 90 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', makerid);
	    }

	    if (TTP_SHOWPRICEYN == "Y"){
	        iTTPBar.AXwindowsfont(260, 90 + TTP_HEIGHTOFFSET, 20, 0, 2, 0, 'Arial', TTP_CURRENCYCHAR + ' ' + customerprice);
	    }else if (TTP_SHOWPRICEYN == "S"){
	        iTTPBar.AXwindowsfont(260, 90 + TTP_HEIGHTOFFSET, 20, 0, 2, 0, 'Arial', customerprice);
	    }

	    if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
	    	//옵션코드에 Z가 들어 있을경우
	    	iTTPBar.AXbarcode('30', 110 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',barcode);
	    	iTTPBar.AXwindowsfont(30,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
	    } else if (TTP_BARCODETYPE == "T") {
	    	// 텐바이텐 일반 물류코드
	    	iTTPBar.AXbarcode('30', 110 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',barcode);
	    	iTTPBar.AXwindowsfont(30,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
	    } else {
	    	// 범용바코드
	    	iTTPBar.AXbarcode('30', 110 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',barcode);
			iTTPBar.AXwindowsfont(30,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', barcode );
	    }

	    // printno 장 프린트
	    iTTPBar.AXprintlabel('1', printno*1);

	    iTTPBar.AXformfeed();
	}

	iTTPBar.AXcloseport();
}

function printTTPOneIndexBarcode(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
	var skipnotinserted = false;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[5]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
	iTTPBar.AXclearbuffer();

	skipnotinserted = true;

	if ((itemname != "") && printno*1 > 0) {
		if (TTP_SHOWDOMAINYN == "Y") {
			iTTPBar.AXwindowsfont(50, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 1, 'Arial', TTP_DOMAINNAME);
		}

	    if (TTP_SHOPBRANDYN == "Y"){
			iTTPBar.AXwindowsfont(50, 40 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', makerid);
	    }

		if (itemoptionname == "") {
			iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname);
		} else {
			iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname + " - " + itemoptionname);
		}

	    if (TTP_SHOWPRICEYN == "Y"){
	        iTTPBar.AXwindowsfont(50, 170 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', '소비자가 : ' + TTP_CURRENCYCHAR + ' ' + customerprice);
	    }

	    iTTPBar.AXwindowsfont(180, 220 + TTP_HEIGHTOFFSET, 110, 0, 2, 0, 'Arial', getTTPBarcodeItemidString(barcode) );

	    if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
	    	//옵션코드에 Z가 들어 있을경우
	    	iTTPBar.AXwindowsfont(370,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
	    } else if (TTP_BARCODETYPE == "T") {
	    	// 텐바이텐 일반 물류코드
	    	iTTPBar.AXwindowsfont(340,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
	    } else {
	    	// 범용바코드
	    	// iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',barcode);
	    }

	    iTTPBar.AXbarcode('50', 335 + TTP_HEIGHTOFFSET, '128', '30', '0', '0', '2', '4',barcode);

	    // printno 장 프린트
	    iTTPBar.AXprintlabel('1', printno*1);

	    iTTPBar.AXformfeed();
	}

	iTTPBar.AXcloseport();
}

//상품 바코드 출력
function printTTPMultiBarcode(arrObject) {
	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var v; var tenBarcode; var pubBarcode; var socname; var socname_kor;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		printno = v.printno;
		socname = v.socname;
		socname_kor = v.socname_kor;

		skipnotinserted = true;
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

//			if (TTP_SHOWDOMAINYN == "Y") {
//				iTTPBar.AXwindowsfont(75, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', TTP_DOMAINNAME);
//			}

		    if (TTP_SHOPBRANDYN == "Y"){
				iTTPBar.AXwindowsfont(20, 10 + TTP_HEIGHTOFFSET, 20, 0, 2, 0, 'Arial', socname_kor + '  ' + socname);
		    }

			//상품명[옵션]
			if (itemoptionname == "") {
				iTTPBar.AXwindowsfont(20, 35 + TTP_HEIGHTOFFSET, 16, 0, 0, 0, 'Arial', itemname);
				//iTTPBar.AXwindowsfont(20, 55 + TTP_HEIGHTOFFSET, 16, 0, 0, 0, 'Arial', '[테스트테스트테스트테스트]');
			} else {
				iTTPBar.AXwindowsfont(20, 35 + TTP_HEIGHTOFFSET, 16, 0, 0, 0, 'Arial', itemname);
				iTTPBar.AXwindowsfont(20, 55 + TTP_HEIGHTOFFSET, 16, 0, 0, 0, 'Arial', '[' + itemoptionname + ']');
			}

		    if (TTP_SHOWPRICEYN == "Y" || TTP_SHOWPRICEYN == "C" || TTP_SHOWPRICEYN == "R"){
		        iTTPBar.AXwindowsfont(230, 80 + TTP_HEIGHTOFFSET, 22, 0, 2, 0, 'Arial', TTP_CURRENCYCHAR + ' ' + customerprice);
		    }else if (TTP_SHOWPRICEYN == "S"){
		    	iTTPBar.AXwindowsfont(230, 80 + TTP_HEIGHTOFFSET, 22, 0, 2, 0, 'Arial', customerprice);
		    }

			iTTPBar.AXwindowsfont(20, 85 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', '___________________________________');

		    //옵션코드에 Z가 들어 있을경우
		    if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
				iTTPBar.AXbarcode('30', 110 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',tenBarcode);
		    	iTTPBar.AXwindowsfont(30,140 + TTP_HEIGHTOFFSET,18,0,0,0,'Arial', getTTPBarcodeString(tenBarcode) );

		    // 텐바이텐 일반 물류코드
		    } else if (TTP_BARCODETYPE == "T") {
				iTTPBar.AXbarcode('30', 110 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',tenBarcode);	//'1001739766Z011'
		    	iTTPBar.AXwindowsfont(30,140 + TTP_HEIGHTOFFSET,18,0,0,0,'Arial', getTTPBarcodeString(tenBarcode) );

			// 범용바코드
		    } else {
		    	iTTPBar.AXbarcode('30', 110 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',pubBarcode);
		    	iTTPBar.AXwindowsfont(30,140 + TTP_HEIGHTOFFSET,18,0,0,0,'Arial', pubBarcode);  //20121219 추가		// 8809436242402
		    }

		    // printno 장 프린트
		    iTTPBar.AXprintlabel('1', printno*1);

		    iTTPBar.AXformfeed();
		}

	}

	iTTPBar.AXcloseport();
}

//쥬얼리 바코드 출력
function printTTPjewelleryMultiBarcode(arrObject) {
	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var v; var tenBarcode; var pubBarcode;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		printno = v.printno;

		skipnotinserted = true;
		if ((itemname != "") && (printno*1 > 0)) {
			tenBarcode = '';
			pubBarcode = '';
			if (TTP_BARCODETYPE == "T") {
				var tmpArr = barcode.split("_");
				if (tmpArr.length != 2) {
					alert("잘못된 형식입니다.");
					return;
				}
				tenBarcode = tmpArr[0];
				pubBarcode = tmpArr[1];
			} else if (TTP_BARCODETYPE == "G") {
				var tmpArr = barcode.split("_");
				if (tmpArr.length != 2) {
					alert("잘못된 형식입니다.");
					return;
				}
				tenBarcode = tmpArr[0];
				pubBarcode = tmpArr[1];
			} else if (TTP_BARCODETYPE == "2") {
				var tmpArr = barcode.split("_");
				if (tmpArr.length != 2) {
					alert("잘못된 형식입니다.");
					return;
				}
				tenBarcode = tmpArr[0];
				pubBarcode = tmpArr[1];
			}

		    if (TTP_SHOPBRANDYN == "Y"){
				iTTPBar.AXwindowsfont(5, 5 + TTP_HEIGHTOFFSET, 18, 0, 2, 0, '10X10', makerid);
		    }

			//상품명[옵션]
			if (itemoptionname == "") {
				iTTPBar.AXwindowsfont(5, 23 + TTP_HEIGHTOFFSET, 13, 0, 0, 0, '10X10', itemname);
			} else {
				iTTPBar.AXwindowsfont(5, 23 + TTP_HEIGHTOFFSET, 13, 0, 0, 0, '10X10', itemname);
				iTTPBar.AXwindowsfont(5, 36 + TTP_HEIGHTOFFSET, 13, 0, 0, 0, '10X10', '[' + itemoptionname + ']');
			}

		    if (TTP_SHOWPRICEYN == "Y" || TTP_SHOWPRICEYN == "C" || TTP_SHOWPRICEYN == "R"){
		        iTTPBar.AXwindowsfont(165, 50 + TTP_HEIGHTOFFSET, 20, 0, 2, 0, 'Arial', TTP_CURRENCYCHAR + ' ' + customerprice);
		    }else if (TTP_SHOWPRICEYN == "S"){
		    	iTTPBar.AXwindowsfont(165, 50 + TTP_HEIGHTOFFSET, 20, 0, 2, 0, 'Arial', customerprice);
		    }

		    iTTPBar.AXwindowsfont(5, 51 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, '10X10', '_________________________');

		    if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
		    	//옵션코드에 Z가 들어 있을경우
		    	iTTPBar.AXbarcode('30', 74 + TTP_HEIGHTOFFSET,'128','20','0','0','1','4',tenBarcode);
		    	iTTPBar.AXwindowsfont(30,95 + TTP_HEIGHTOFFSET,15,0,0,0,'10X10', getTTPBarcodeString(tenBarcode) );
		    } else if (TTP_BARCODETYPE == "T") {
		    	// 텐바이텐 일반 물류코드
		    	iTTPBar.AXbarcode('30', 74 + TTP_HEIGHTOFFSET,'128','20','0','0','1','4',tenBarcode);	//'1001739766Z011'
		    	iTTPBar.AXwindowsfont(30,95 + TTP_HEIGHTOFFSET,15,0,0,0,'10X10', getTTPBarcodeString(tenBarcode) );
		    } else {
		    	// 범용바코드
		    	iTTPBar.AXbarcode('30', 74 + TTP_HEIGHTOFFSET,'128','20','0','0','1','4',pubBarcode);
		    	iTTPBar.AXwindowsfont(30,95 + TTP_HEIGHTOFFSET,15,0,0,0,'10X10', pubBarcode);		// 8809436242402
		    }

		    // printno 장 프린트
		    iTTPBar.AXprintlabel('1', printno*1);

		    iTTPBar.AXformfeed();
		}

	}

	iTTPBar.AXcloseport();
}

function printTTPOneItemLabel(barcode, makerid, makername, itemid, itemname, itemoptionname, customerprice, printno) {
	var skipnotinserted = false;
	var tenBarcode, pubBarcode;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[5]');
		return;
	}


	//alert("111");
	// return;

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
	iTTPBar.AXclearbuffer();

	skipnotinserted = true;

	if ((itemname != "") && printno*1 > 0) {
		if (TTP_SHOWDOMAINYN == "Y") {
			iTTPBar.AXwindowsfont(50, 20 + TTP_HEIGHTOFFSET, 30, 0, 2, 1, 'Arial', ("        " + TTP_DOMAINNAME + "        "));
		}

	    if (TTP_SHOPBRANDYN == "Y"){
			if (makername != "") {
				iTTPBar.AXwindowsfont(50, 55 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', (makerid + "(" + makername + ")"));
			} else {
				iTTPBar.AXwindowsfont(50, 55 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', (makerid));
			}

	    }

		iTTPBar.AXwindowsfont(50, 90 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname);
		if (itemoptionname != "") {
			iTTPBar.AXwindowsfont(50, 140 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemoptionname);
		}

		tenBarcode = "";
		pubBarcode = "";

		if (TTP_BARCODETYPE == "T") {
			tenBarcode = barcode;
		} else if (TTP_BARCODETYPE == "G") {
			pubBarcode = barcode;
		} else if (TTP_BARCODETYPE == "2") {
			var tmpArr = barcode.split("_");
			if (tmpArr.length != 2) {
				alert("잘못된 형식입니다.");
				return;
			}
			tenBarcode = tmpArr[0];
			pubBarcode = tmpArr[1];
		}

		iTTPBar.AXwindowsfont(110, 180 + TTP_HEIGHTOFFSET, 130, 0, 2, 0, 'Arial', getTTPBarcodeItemidString(tenBarcode) );

		if ((TTP_BARCODETYPE == "T") || (TTP_BARCODETYPE == "2")) {
			if (InStr(tenBarcode, 'Z') >= 0) {
				iTTPBar.AXwindowsfont(370,310 + TTP_HEIGHTOFFSET,30,0,2,0,'Arial', tenBarcode);
			} else {
				iTTPBar.AXwindowsfont(340,310 + TTP_HEIGHTOFFSET,30,0,2,0,'Arial', tenBarcode);
			}
			if (pubBarcode != "") {
				iTTPBar.AXwindowsfont(340,340 + TTP_HEIGHTOFFSET,30,0,2,0,'Arial', pubBarcode);
			}
		}
		/*
	    if (((TTP_BARCODETYPE == "T") || (TTP_BARCODETYPE == "2")) && (InStr(tenBarcode, 'Z') >= 0)) {
	    	//옵션코드에 Z가 들어 있을경우
	    	iTTPBar.AXwindowsfont(370,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', tenBarcode);
	    } else if (TTP_BARCODETYPE == "T") {
	    	// 텐바이텐 일반 물류코드
	    	iTTPBar.AXwindowsfont(340,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', tenBarcode);
	    } else if (TTP_BARCODETYPE == "2") {
	    	// 텐바이텐 일반 물류코드
	    	iTTPBar.AXwindowsfont(340,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
	    } else {
	    	// 범용바코드
	    	// iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',barcode);
	    }
		*/

	    iTTPBar.AXbarcode('50', 335 + TTP_HEIGHTOFFSET, '128', '30', '0', '0', '2', '4',tenBarcode);

	    // printno 장 프린트
	    iTTPBar.AXprintlabel('1', printno*1);

	    //iTTPBar.AXformfeed();
	}

	iTTPBar.AXcloseport();
}

// 선택상품 인덱스 출력
function printTTPMultiItemLabel(arrObject) {
	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno, brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode;
	var v;
	var tenBarcode, pubBarcode;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

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

		if (TTP_BARCODETYPE == "T") {
			tenBarcode = barcode;
		} else if (TTP_BARCODETYPE == "G") {
			pubBarcode = barcode;
		} else if (TTP_BARCODETYPE == "2") {
			var tmpArr = barcode.split("_");
			if (tmpArr.length != 2) {
				alert("잘못된 형식입니다.");
				return;
			}
			tenBarcode = tmpArr[0];
			pubBarcode = tmpArr[1];
		}

		skipnotinserted = true;
		if ((itemname != "") && (printno*1 > 0)) {
//			if (TTP_SHOWDOMAINYN == "Y") {
//				iTTPBar.AXwindowsfont(50, 20 + TTP_HEIGHTOFFSET, 30, 0, 2, 1, 'Arial', ("        " + TTP_DOMAINNAME + "        "));
//			}

			// 브랜드랙과 보조랙에 값이 있는데 브랜드랙과 보조랙 값이 틀린거
			if (brandrackcode!="" && subitemrackcode!="" && brandrackcode!=subitemrackcode){
				iTTPBar.AXwindowsfont(40, 20 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', "보조랙 : " + subitemrackcode);
			}else if (brandrackcode!=""){
				iTTPBar.AXwindowsfont(40, 20 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', "브랜드랙 : " + brandrackcode);
			}

			if (TTP_SHOPBRANDYN == "Y"){
				iTTPBar.AXwindowsfont(40, 55 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', (makerid));
			}

			iTTPBar.AXwindowsfont(40, 90 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname);
			if (itemoptionname != "") {
				iTTPBar.AXwindowsfont(40, 120 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemoptionname);
			}

			iTTPBar.AXwindowsfont(70, 160 + TTP_HEIGHTOFFSET, 130, 0, 2, 0, 'Arial', getTTPBarcodeItemidString(tenBarcode) );

			if ((TTP_BARCODETYPE == "T") || (TTP_BARCODETYPE == "2")) {
				if (InStr(tenBarcode, 'Z') >= 0) {
					//상품구분 + 상품코드
					iTTPBar.AXwindowsfont(310,280 + TTP_HEIGHTOFFSET,30,0,2,0,'Arial', left(tenBarcode,2) + "-" + getTTPBarcodeItemidString(tenBarcode));
					//옵션
					iTTPBar.AXwindowsfont(470,275 + TTP_HEIGHTOFFSET,50,0,2,0,'Arial', "-" + right(tenBarcode,4));
				} else {
					//상품구분 + 상품코드
					iTTPBar.AXwindowsfont(310,280 + TTP_HEIGHTOFFSET,30,0,2,0,'Arial', left(tenBarcode,2) + "-" + getTTPBarcodeItemidString(tenBarcode));
					//옵션
					iTTPBar.AXwindowsfont(470,275 + TTP_HEIGHTOFFSET,50,0,2,0,'Arial', "-" + right(tenBarcode,4));
				}
				if (pubBarcode != "") {
					iTTPBar.AXwindowsfont(360,330 + TTP_HEIGHTOFFSET,35,0,2,0,'Arial', pubBarcode);
				}
			}

			iTTPBar.AXbarcode('40', 335 + TTP_HEIGHTOFFSET, '128', '30', '0', '0', '2', '4',tenBarcode);

			// printno 장 프린트
			iTTPBar.AXprintlabel('1', printno*1);

			// 띄우지 않고 출력한다.
			// iTTPBar.AXformfeed();
		}

	}

	iTTPBar.AXcloseport();
}

function printTTPMultiIndexBarcode(arrObject) {
	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var v;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		printno = v.printno;

		skipnotinserted = true;
		if ((itemname != "") && (printno*1 > 0)) {
			if (TTP_SHOWDOMAINYN == "Y") {
				iTTPBar.AXwindowsfont(50, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 1, 'Arial', TTP_DOMAINNAME);
			}

			if (TTP_SHOPBRANDYN == "Y"){
				iTTPBar.AXwindowsfont(50, 40 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', makerid);
			}

			if (itemoptionname == "") {
				iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname);
			} else {
				iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname + " - " + itemoptionname);
			}

			if (TTP_SHOWPRICEYN == "Y"){
				iTTPBar.AXwindowsfont(50, 170 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', '소비자가 : ' + TTP_CURRENCYCHAR + ' ' + customerprice);
			}

			iTTPBar.AXwindowsfont(180, 220 + TTP_HEIGHTOFFSET, 110, 0, 2, 0, 'Arial', getTTPBarcodeItemidString(barcode) );

			if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
				//옵션코드에 Z가 들어 있을경우
				iTTPBar.AXwindowsfont(370,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
			} else if (TTP_BARCODETYPE == "T") {
				// 텐바이텐 일반 물류코드
				iTTPBar.AXwindowsfont(340,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
			} else {
				// 범용바코드
				// iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',barcode);
			}

			iTTPBar.AXbarcode('50', 335 + TTP_HEIGHTOFFSET, '128', '30', '0', '0', '2', '4',barcode);

			// printno 장 프린트
			iTTPBar.AXprintlabel('1', printno*1);

			// 띄우지 않고 출력한다.
			// iTTPBar.AXformfeed();
		}

	}

	iTTPBar.AXcloseport();
}

function printTTPMultiEquipBarcode(arrObject) {
	var skipnotinserted = false;
	var equip_code, AccountGubunName, equip_name, buy_date, barcode;
	var tmp;
	var v;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		equip_code = v.equip_code;
		AccountGubunName = v.AccountGubunName;
		equip_name = v.equip_name;
		buy_date = v.buy_date;
		tmp = equip_code.split("-");
		barcode = tmp[1] + tmp[1];

		// ① ② ③ ④ ⑤ ⑥
		iTTPBar.AXwindowsfont(50, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', "① 장비코드 :");
		iTTPBar.AXwindowsfont(50, 40 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', "② 자산구분:");
		iTTPBar.AXwindowsfont(50, 80 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', "③ 장비구분:");
		iTTPBar.AXwindowsfont(50, 120 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', "④ 취득일자:");
		iTTPBar.AXwindowsfont(50, 160 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', "⑤ 문    의:");

		iTTPBar.AXwindowsfont(200, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', equip_code);
		iTTPBar.AXwindowsfont(200, 40 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', AccountGubunName);
		iTTPBar.AXwindowsfont(200, 80 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', (AccountGubunName + "_"+ equip_name).substring(0, 25));
		iTTPBar.AXwindowsfont(200, 120 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', buy_date);
		iTTPBar.AXwindowsfont(200, 160 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', "운영기획");

		iTTPBar.AXbarcode('300', 280 + TTP_HEIGHTOFFSET, '128', '30', '0', '0', '2', '4', barcode);
		iTTPBar.AXwindowsfont(150, 320 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', "이 물품은 텐바이텐의 소중한 재산입니다.");

		// printno 장 프린트
		iTTPBar.AXprintlabel('1', 1);

		// 띄우지 않고 출력한다.
		// iTTPBar.AXformfeed();

	}

	iTTPBar.AXcloseport();
}

function printTTPMultiEquipSmallBarcode(arrObject) {
	var skipnotinserted = false;
	var equip_code, AccountGubunName, equip_name, buy_date, barcode;
	var tmp;
	var v;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		equip_code = v.equip_code;
		AccountGubunName = v.AccountGubunName;
		equip_name = v.equip_name;
		buy_date = v.buy_date;
		tmp = equip_code.split("-");
		barcode = tmp[1] + tmp[1];

		iTTPBar.AXwindowsfont(60, 0 + TTP_HEIGHTOFFSET, 25, 0, 2, 0, 'Arial', equip_code);
		iTTPBar.AXwindowsfont(25, 55 + TTP_HEIGHTOFFSET, 25, 0, 2, 0, 'Arial', (equip_name).substring(0, 25));

		iTTPBar.AXbarcode('50', 90 + TTP_HEIGHTOFFSET, '128', '25', '0', '0', '2', '4', barcode);
		iTTPBar.AXwindowsfont(55, 130 + TTP_HEIGHTOFFSET, 15, 0, 2, 0, 'Arial', "이 물품은 텐바이텐의 소중한 재산입니다.");

		// printno 장 프린트
		iTTPBar.AXprintlabel('1', 1);

		// 띄우지 않고 출력한다.
		// iTTPBar.AXformfeed();

	}

	iTTPBar.AXcloseport();
}

// 불량상품용 인덱스바코드
function printTTPOneIndexBarcodeForBadItem(barcode, makerid, itemname, itemoptionname, regdate, printno) {
	var skipnotinserted = false;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[5]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
	iTTPBar.AXclearbuffer();

	skipnotinserted = true;

	if ((itemname != "") && printno*1 > 0) {
		if (TTP_SHOWDOMAINYN == "Y") {
			iTTPBar.AXwindowsfont(50, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 1, 'Arial', TTP_DOMAINNAME);
		}

		iTTPBar.AXwindowsfont(50, 40 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', "등록일 : " + regdate);

	    if (TTP_SHOPBRANDYN == "Y"){
			iTTPBar.AXwindowsfont(50, 80 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', makerid);
	    }

		if (itemoptionname == "") {
			iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname);
		} else {
			iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname + " - " + itemoptionname);
		}

		iTTPBar.AXwindowsfont(50, 170 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', '불량사유 : ');

	    // iTTPBar.AXwindowsfont(180, 220 + TTP_HEIGHTOFFSET, 110, 0, 2, 0, 'Arial', getTTPBarcodeItemidString(barcode) );

	    if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
	    	//옵션코드에 Z가 들어 있을경우
	    	iTTPBar.AXwindowsfont(370,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
	    } else if (TTP_BARCODETYPE == "T") {
	    	// 텐바이텐 일반 물류코드
	    	iTTPBar.AXwindowsfont(340,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
	    } else {
	    	// 범용바코드
	    	// iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'128','30','0','0','2','4',barcode);
	    }

	    iTTPBar.AXbarcode('50', 335 + TTP_HEIGHTOFFSET, '128', '30', '0', '0', '2', '4',barcode);

	    // printno 장 프린트
	    iTTPBar.AXprintlabel('1', printno*1);

	    // iTTPBar.AXformfeed();
	}

	iTTPBar.AXcloseport();
}

function printTTPOneIndexBarcodeForEventItem(eventCode, eventName01, eventName02, eventStartdate, eventEnddate, eventGiftCode, eventGiftKind, eventGift01, eventGift02, printno) {

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[5]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
	iTTPBar.AXclearbuffer();

	if (printno*1 > 0) {
		iTTPBar.AXwindowsfont(400, 30 + TTP_HEIGHTOFFSET, 80, 0, 0, 0,'Arial Bold', eventCode);

		iTTPBar.AXwindowsfont(50, 30 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', "[시작일]");
		iTTPBar.AXwindowsfont(170, 30 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', eventStartdate);
		iTTPBar.AXwindowsfont(50, 70 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', "[종료일]");
		iTTPBar.AXwindowsfont(170, 70 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', eventEnddate);

		iTTPBar.AXwindowsfont(50, 110 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', "[이벤트명]");
		iTTPBar.AXwindowsfont(170, 110 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', eventName01);
		iTTPBar.AXwindowsfont(170, 150 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', eventName02);

		iTTPBar.AXwindowsfont(50, 190 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', "-----------------------------------------------------------------------------");

		iTTPBar.AXwindowsfont(400, 230 + TTP_HEIGHTOFFSET, 80, 0, 0, 0,'Arial Bold', eventGiftCode);

		iTTPBar.AXwindowsfont(50, 270 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', "[사은품]");
		iTTPBar.AXwindowsfont(170, 270 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', eventGiftKind);			// 한줄에 한글 24자까지 넘으면 아래로
		iTTPBar.AXwindowsfont(170, 310 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', eventGift01);
		iTTPBar.AXwindowsfont(170, 350 + TTP_HEIGHTOFFSET, 25, 0, 0, 0,'Arial', eventGift02);

	    // printno 장 프린트
	    iTTPBar.AXprintlabel('1', printno*1);

	    // iTTPBar.AXformfeed();
	}

	iTTPBar.AXcloseport();
}

// 10010000000000 => 10-01000000-0000
// 104444440000 => 10-444444-0000
function getTTPBarcodeString(barcode) {
	var itemgubun, itemid, itemoption;

	itemgubun 	= barcode.substring(0, 2);
	itemid 		= barcode.substring(2, (barcode.length - 4));
	itemoption	= barcode.substring((barcode.length - 4), barcode.length);

	return itemgubun + "-" + itemid + "-" + itemoption;
}

// =============================================================================
// 100만 이상
// 10010000000000 => 01000000
// 1000000 => 01000000
// =============================================================================
// 100만 미만
// 100444440000 => 044444
// 444444 => 0444444
// =============================================================================
function getTTPBarcodeItemidString(barcode) {
	var itemgubun, itemid, itemoption;

	if (barcode.length >= 12) {
		itemid 		= barcode.substring(2, (barcode.length - 4));
	} else {
		if ((barcode*1) >= 1000000) {
			itemid = (100000000 + barcode*1) + "";
			itemid = itemid.substring((itemid.length - 8), itemid.length);
		} else {
			itemid = (1000000 + barcode*1) + "";
			itemid = itemid.substring((itemid.length - 6), itemid.length);
		}
	}

	return itemid;
}

function printTTPInnerBoxBarcode(baljudate, baljuid, boxno, innerboxweight, innerboxbarcode, innerboxbarcodeforshow) {
	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[5]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand("DIRECTION 1");
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, "2", "10", "0", TTP_PAPERMARGIN, "0");

	iTTPBar.AXwindowsfont(30,0 + TTP_HEIGHTOFFSET,40,0,2,1,"Arial","                INNER BOX INDEX               ");
	iTTPBar.AXwindowsfont(30,55 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","SHOPID : " + baljuid);
	iTTPBar.AXwindowsfont(30,90 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","                " + TTP_DOMAINNAME);
	iTTPBar.AXwindowsfont(30,130 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","DATE : " + baljudate);
	iTTPBar.AXwindowsfont(30,170 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","INNER BOX NO. : " + boxno);
	iTTPBar.AXwindowsfont(30,210 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","INNER BOX WEIGHT : " + innerboxweight + " KG");
	iTTPBar.AXbarcode("160",280 + TTP_HEIGHTOFFSET,"128","40","0","0","2","4", innerboxbarcode);
	iTTPBar.AXwindowsfont(30,345 + TTP_HEIGHTOFFSET,30,0,0,0,"Arial","                      " + innerboxbarcodeforshow);

	iTTPBar.AXprintlabel("1","1");
	iTTPBar.AXcloseport();
}

// 수기 인덱스 출력
function printTTPMultiIndexSudongLabel(msg,itemno,fontName) {
	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

	if (itemno=='' || itemno==0) itemno=1;
    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	var skipnotinserted = false;
	var lines = msg.split("\n");

	// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
	iTTPBar.AXclearbuffer();

	for (var i = 0; i < lines.length; i++) {
		iTTPBar.AXwindowsfont(20, 10 + TTP_HEIGHTOFFSET + (35 * i), 30, 0, 0, 0, fontName, lines[i].replace("\r", ""));
	}

	// printno 장 프린트
	iTTPBar.AXprintlabel('1', itemno*1);

	// 띄우지 않고 출력한다.
	// iTTPBar.AXformfeed();
	iTTPBar.AXcloseport();
}

// 수기 랙번호 인덱스 출력
function printTTPRackIndexSudongLabel(msg,itemno,fontName,barcodeprintyn) {
	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

	if (fontName=='') fontName='10X10';
	if (itemno=='' || itemno==0) itemno=1;
    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	var skipnotinserted = false;
	var lines = msg.split("\n");
	var linemsg = "";

	for (var i = 0; i < lines.length; i++) {
		linemsg = lines[i].replace("\r", "").ltrim().rtrim();

		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		if (linemsg.length < 5){
			iTTPBar.AXwindowsfont(10, 70 + TTP_HEIGHTOFFSET, 190, 0, 0, 0, fontName, linemsg);
			if (barcodeprintyn!=''){
				iTTPBar.AXbarcode('90', 300 + TTP_HEIGHTOFFSET, '128', '50', '0', '0', '4', '4',linemsg);
			}
		}else if (linemsg.length < 8){
			iTTPBar.AXwindowsfont(10, 70 + TTP_HEIGHTOFFSET, 110, 0, 0, 0, fontName, linemsg);
			if (barcodeprintyn!=''){
				iTTPBar.AXbarcode('70', 300 + TTP_HEIGHTOFFSET, '128', '50', '0', '0', '4', '4',linemsg);
			}
		}else{
			iTTPBar.AXwindowsfont(10, 70 + TTP_HEIGHTOFFSET, 95, 0, 0, 0, fontName, linemsg);
			if (barcodeprintyn!=''){
				iTTPBar.AXbarcode('30', 300 + TTP_HEIGHTOFFSET, '128', '50', '0', '0', '4', '4',linemsg);
			}
		}

		// printno 장 프린트
		iTTPBar.AXprintlabel('1', itemno*1);
	}

	// 띄우지 않고 출력한다.
	// iTTPBar.AXformfeed();
	iTTPBar.AXcloseport();
}

// 랙번호 인덱스 출력
function printTTPMultiRackIndexLabel(arrObject) {
	var v;
	var msg, itemno, fontName;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	var skipnotinserted = false;
	var linemsg = "";

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		msg = v.msg;
		itemno = v.itemno;
		fontName = v.fontName;
		if (fontName=='') fontName='10X10';

		linemsg = msg.replace("\r", "").ltrim().rtrim();

		if (linemsg.length < 5){
			iTTPBar.AXwindowsfont(10, 70 + TTP_HEIGHTOFFSET, 190, 0, 0, 0, fontName, linemsg);
			iTTPBar.AXbarcode('90', 300 + TTP_HEIGHTOFFSET, '128', '50', '0', '0', '4', '4',linemsg);
		}else if (linemsg.length < 8){
			iTTPBar.AXwindowsfont(10, 70 + TTP_HEIGHTOFFSET, 115, 0, 0, 0, fontName, linemsg);
			iTTPBar.AXbarcode('70', 300 + TTP_HEIGHTOFFSET, '128', '50', '0', '0', '4', '4',linemsg);
		}else{
			iTTPBar.AXwindowsfont(10, 70 + TTP_HEIGHTOFFSET, 95, 0, 0, 0, fontName, linemsg);
			iTTPBar.AXbarcode('30', 300 + TTP_HEIGHTOFFSET, '128', '50', '0', '0', '4', '4',linemsg);
		}

		// printno 장 프린트
		iTTPBar.AXprintlabel('1', itemno*1);
	}

	// 띄우지 않고 출력한다.
	// iTTPBar.AXformfeed();
	iTTPBar.AXcloseport();
}

function drawTTPprintOcxV2__(iname, iversion) {
    var iObjStr = "";
    iObjStr = "<OBJECT"
    iObjStr = iObjStr + "      name='" + iname + "'";
    iObjStr = iObjStr + "	  classid='clsid:4B4DE9A2-A9B5-403B-8AFF-4967823E3BB2'";
    iObjStr = iObjStr + "	  codebase='http://webadmin.10x10.co.kr/common/cab/TenTTPBar.cab#version=" + iversion + "'";
    iObjStr = iObjStr + "	  width=0";
    iObjStr = iObjStr + "	  height=0";
    iObjStr = iObjStr + "	  align=center";
    iObjStr = iObjStr + "	  hspace=0";
    iObjStr = iObjStr + "	  vspace=0";
    iObjStr = iObjStr + ">";
    iObjStr = iObjStr + "</OBJECT>";

    document.write(iObjStr);
}

// TTP 모듈 설치
drawTTPprintOcxV2__('iTTPBar','1,0,0,3');
