//####################################################
// Description :  ���ù� ���ڵ� JS
// History : 2016.11.24 �ѿ�� ����
///////////////////// ������ ������ �ؿ� ���ϵ� ��� �����ϰ� ���� ���ľ� �Ѵ�. ////////////////////////
// SCM : /js/DOSHIBAbarcode.js
//		 /js/DOSHIBAbarcode_utf8.js
// LOGICS : /js/DOSHIBAbarcode.js
//			/js/DOSHIBAbarcode_utf8.js
//####################################################

//////////////////////////////// �⺻�� & ��� ��ġ //////////////////////////////
var TOSHIBA_PAPERWIDTH = 450;
var TOSHIBA_PAPERHEIGHT	= 220;
var TOSHIBA_PAPERMARGIN	= 3;
var TOSHIBA_HEIGHTOFFSET = 0;
var TOSHIBA_WIDTHOFFSET	= 0;
var TOSHIBA_isforeignprint = 'N';
var TOSHIBA_currencyChar = '��';
var TOSHIBA_DOMAINNAME = 'www.10x10.co.kr';
var TOSHIBA_SHOWDOMAINYN = 'Y';
var TOSHIBA_SHOWPRICEYN = 'Y';
var TOSHIBA_SHOPBRANDYN = 'Y';
var TOSHIBA_BARCODETYPE = 'T';
var TOSHIBA_currencyunit = 'KRW';
var TOSHIBA_currencyunit_pos = 'KRW';

var TOSHIBA_brand_str = 'ǰ��';
var TOSHIBA_origin_str = '����';
var TOSHIBA_material_str = '����';
var TOSHIBA_standard_str = '�԰�';
var TOSHIBA_manufacturer_str = '������';
var TOSHIBA_address_str = '�ּ�';
var TOSHIBA_import_str = '���Ի�';
var TOSHIBA_telephone_str = '��ȭ';

// ���ù� ��� ��ġ
DrawReceiptPrintobj_TEC("TEC_DO3","TEC B-FV4");
//////////////////////////////// �⺻�� & ��� ��ġ //////////////////////////////

// �ؿܹ��ڵ��
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

// ���� ���ڵ��
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

// ���� �ε��� ���ڵ��
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

// ���� ���ڵ��
function BarcodeDataClass_udong(msg,itemno,fontName) {
    this.msg = msg;
    this.itemno = itemno;
    this.fontName = fontName;
}

//�ؿ� ���ڵ� ���
function printTOSHIBAforeignBarcode(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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
		//���� ������ŭ ��µǴ°� ���� �ȴµ�;; �켱 ���� ������ �̴´�. ��ɾ� ã�Ƽ� ����� �ٲ����
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

				//�߱��� ¥����... ������ ����. ? �̱�¥�� ����.
				//Arial(�ٱ��� ��Ʈ.������ �⺻ ��Ʈ)
				//Microsoft JhengHei(����ü.�븸) , Microsoft YaHei(����ü.�߱�)

				//ǰ��. ��٣(ǰ��)
				TEC_DO3.PrintText(0,10 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_brand_str+' : ['+makerid+'] '+catename);

				//����. ߧ�(����)
				TEC_DO3.PrintText(0,25 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_origin_str+' : '+sourcearea);

				//����. ���(����)
				TEC_DO3.PrintText(0,41 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_material_str+' : '+itemsource);

				//�԰�. Ю̫(�԰�)
				TEC_DO3.PrintText(0,57 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_standard_str+' : '+itemsize);

				//������. �����(������)
				TEC_DO3.PrintText(0,72 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_manufacturer_str+' : '+manufacturer);

				//��ȭ. ���(��ȭ)
				TEC_DO3.PrintText(0,87 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_telephone_str+' : '+m_telephone);

				//�ּ�. ��(����)
				TEC_DO3.PrintText(0,101 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_address_str+' : '+m_address1);
				TEC_DO3.PrintText(10,115 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, m_address2);

				//���Ի�. ��Ϣ��(������)
				TEC_DO3.PrintText(0,131 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_import_str+' : '+vimport);

				//��ȭ. ���(��ȭ)
				TEC_DO3.PrintText(0,146 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_telephone_str+' : '+i_telephone);

				//�ּ�. ��(����)
				TEC_DO3.PrintText(0,161 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, TOSHIBA_address_str+' : '+i_address1);
				TEC_DO3.PrintText(10,176 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 0, 0, i_address2);

				TEC_DO3.PrintText(0,177 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 0, 0, '___________________________________');

				//����
			    if (TOSHIBA_SHOWPRICEYN == "Y" || TOSHIBA_SHOWPRICEYN == "C" || TOSHIBA_SHOWPRICEYN == "R"){
			        TEC_DO3.PrintText(0,197 + TOSHIBA_HEIGHTOFFSET, "Arial", 25*F, 1, 0, TOSHIBA_currencyChar+' '+customerprice);
			    }else if (TOSHIBA_SHOWPRICEYN == "S"){
			    	TEC_DO3.PrintText(0,197 + TOSHIBA_HEIGHTOFFSET, "Arial", 25*F, 1, 0, customerprice);
			    }

				//�귣��
				if (TOSHIBA_SHOPBRANDYN == "Y"){
					TEC_DO3.PrintText(100,197 + TOSHIBA_HEIGHTOFFSET, "Arial", 21*F, 1, 0, makerid);
				}

				//��ǰ��[�ɼ�]
				if (itemoptionname == "") {
					TEC_DO3.PrintText(0,218 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 1, 0, itemname);
					//TEC_DO3.PrintText(0,234 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 1, 0, '[�׽�Ʈ�׽�Ʈ�׽�Ʈ�׽�Ʈ�׽�Ʈ�׽�Ʈ�׽�Ʈ]');
				} else {
					TEC_DO3.PrintText(0,218 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 1, 0, itemname);
					TEC_DO3.PrintText(0,234 + TOSHIBA_HEIGHTOFFSET, "Arial", 15*F, 1, 0, '[' + itemoptionname + ']');
				}

				if (TOSHIBA_BARCODETYPE == "A") {
					//�ɼ��ڵ忡 Z�� ��� �������
				    if (InStr(barcode, 'Z') >= 0) {
				    	TEC_DO3.PrintText(0, 254 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 25, 0, 0, tenBarcode);	//'1001739766Z011
				    	TEC_DO3.PrintText(0,281 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 1, 0, getTTPBarcodeString(tenBarcode));

					// �ٹ����� �Ϲ� �����ڵ�
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

//�ε��� ���ڵ� ���
function printTOSHIBAMultiItemLabel(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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
				alert("�߸��� �����Դϴ�.");
				return;
			}
			tenBarcode = tmpArr[0];
			pubBarcode = tmpArr[1];
		}

		skipnotinserted = true;

		//TEC_DO3.SetPrinterCopies('TEC B-FV4', 2);
		//���� ������ŭ ��µǴ°� ���� �ȴµ�;; �켱 ���� ������ �̴´�. ��ɾ� ã�Ƽ� ����� �ٲ����
		for (var j = 0;j < printno; j++) {
			TEC_DO3.PrinterOpen();

			if ((itemname != "") && (printno*1 > 0)) {
//				if (TOSHIBA_SHOWDOMAINYN == "Y") {
//					TEC_DO3.PrintText(190,20 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, TOSHIBA_DOMAINNAME);
//				}

				// �귣�巢�� �������� ���� �ִµ� �귣�巢�� ������ ���� Ʋ����
				if (brandrackcode!="" && subitemrackcode!="" && brandrackcode!=subitemrackcode){
					TEC_DO3.PrintText(40,20 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "������ : " + subitemrackcode);

				}else if (brandrackcode!=""){
					TEC_DO3.PrintText(40,20 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "�귣�巢 : " + brandrackcode);
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
					//��ǰ���� + ��ǰ�ڵ�
					TEC_DO3.PrintText(310,280 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, left(tenBarcode,2) + "-" + getTTPBarcodeItemidString(tenBarcode));
					//�ɼ�
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

//��ǰ ���ڵ� ���
function printTOSHIBAMultiBarcode(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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

		//ǥ�� ��ǰ��(����,�ؿ�)
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
		//���� ������ŭ ��µǴ°� ���� �ȴµ�;; �켱 ���� ������ �̴´�. ��ɾ� ã�Ƽ� ����� �ٲ����
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
//					if (TOSHIBA_currencyChar!='��'){
//						TEC_DO3.PrintText(75,0 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, TOSHIBA_DOMAINNAME);
//					}
//				}

				// �귣��
				if (TOSHIBA_SHOPBRANDYN == "Y"){
					TEC_DO3.PrintText(10,10 + TOSHIBA_HEIGHTOFFSET, "Arial", 20*F, 1, 0, socname_kor + '  ' + socname);
				}

				//��ǰ��[�ɼ�]
				if (itemoptionname == "") {
					TEC_DO3.PrintText(10,35 + TOSHIBA_HEIGHTOFFSET, "Arial", 16*F, 0, 0, itemname);
					//TEC_DO3.PrintText(10,55 + TOSHIBA_HEIGHTOFFSET, "Arial", 16*F, 0, 0, '[�׽�Ʈ�׽�Ʈ�׽�Ʈ�׽�Ʈ]');
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

				//�ɼ��ڵ忡 Z�� ��� �������
			    if ((TOSHIBA_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
			    	TEC_DO3.PrintText(10, 110 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 30, 0, 0, tenBarcode);
			    	TEC_DO3.PrintText(10,140 + TOSHIBA_HEIGHTOFFSET, "Arial", 18*F, 0, 0, getTTPBarcodeString(tenBarcode));
			    	
				// �ٹ����� �Ϲ� �����ڵ�
			    } else if (TOSHIBA_BARCODETYPE == "T") {
			    	TEC_DO3.PrintText(10, 110 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMEAN128", 30, 0, 0, tenBarcode);	//'1001739766Z011'
			    	TEC_DO3.PrintText(45,140 + TOSHIBA_HEIGHTOFFSET, "Arial", 18*F, 0, 0, getTTPBarcodeString(tenBarcode));

				// ������ڵ�
			    } else {
			    	TEC_DO3.PrintText(10, 110 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 30, 0, 0, pubBarcode);
			    	TEC_DO3.PrintText(10,140 + TOSHIBA_HEIGHTOFFSET, "Arial", 18*F, 0, 0, pubBarcode);
			    }
			}

			TEC_DO3.PrinterClose();
		}
	}
}

//��� ���ڵ� ���
function printTOSHIBAjewelleryMultiBarcode(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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

		//ǥ�� ��ǰ��(����,�ؿ�)
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
		//���� ������ŭ ��µǴ°� ���� �ȴµ�;; �켱 ���� ������ �̴´�. ��ɾ� ã�Ƽ� ����� �ٲ����
		for (var j = 0;j < printno; j++) {
			TEC_DO3.PrinterOpen();

			if ((itemname != "") && (printno*1 > 0)) {
				tenBarcode = '';
				pubBarcode = '';
				if (TOSHIBA_BARCODETYPE == "T") {
					var tmpArr = barcode.split("_");
					if (tmpArr.length != 2) {
						alert("�߸��� �����Դϴ�.");
						return;
					}
					tenBarcode = tmpArr[0];
					pubBarcode = tmpArr[1];
				} else if (TOSHIBA_BARCODETYPE == "G") {
					var tmpArr = barcode.split("_");
					if (tmpArr.length != 2) {
						alert("�߸��� �����Դϴ�.");
						return;
					}
					tenBarcode = tmpArr[0];
					pubBarcode = tmpArr[1];
				} else if (TOSHIBA_BARCODETYPE == "2") {
					var tmpArr = barcode.split("_");
					if (tmpArr.length != 2) {
						alert("�߸��� �����Դϴ�.");
						return;
					}
					tenBarcode = tmpArr[0];
					pubBarcode = tmpArr[1];
				}

				if (TOSHIBA_SHOPBRANDYN == "Y"){
					TEC_DO3.PrintText(5,5 + TOSHIBA_HEIGHTOFFSET, "10X10", 18*F, 1, 0, makerid);
				}

				//��ǰ��[�ɼ�]
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

				//�ɼ��ڵ忡 Z�� ��� �������
			    if ((TOSHIBA_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
			    	TEC_DO3.PrintText(20, 74 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 20, 0, 0, tenBarcode);
			    	TEC_DO3.PrintText(20,95 + TOSHIBA_HEIGHTOFFSET, "10X10", 15*F, 0, 0, getTTPBarcodeString(tenBarcode));

				// �ٹ����� �Ϲ� �����ڵ�
			    } else if (TOSHIBA_BARCODETYPE == "T") {
			    	TEC_DO3.PrintText(20, 74 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 20, 0, 0, tenBarcode);	//'1001739766Z011'
			    	TEC_DO3.PrintText(20,95 + TOSHIBA_HEIGHTOFFSET, "10X10", 15*F, 0, 0, getTTPBarcodeString(tenBarcode));

				// ������ڵ�
			    } else {
			    	TEC_DO3.PrintText(20, 74 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMSMALLCODE128", 20, 0, 0, pubBarcode);
			    	TEC_DO3.PrintText(20,95 + TOSHIBA_HEIGHTOFFSET, "10X10", 15*F, 0, 0, pubBarcode);
			    }
			}

			TEC_DO3.PrinterClose();
		}
	}
}

//DAS ���ڵ� ���
function printTOSHIBADasMultiBarcode(arrObject) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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
			TEC_DO3.PrintText(40,35 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "[���] ��ǰ ������ù��ڵ�");
			TEC_DO3.PrintText(60, 70 + TOSHIBA_HEIGHTOFFSET, "TEC-ITEMEAN128", 35, 0, 0, v.baljubarcode);
			TEC_DO3.PrintText(70,120 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, v.baljubarcode);

			TEC_DO3.PrinterClose();
		}

		v.dasindexarr = v.dasindexarr.split("|");
		if (v.danpumynarr) {
			v.danpumynarr = v.danpumynarr.split("|");
		}
		for (var j = 0; j < v.dasindexarr.length; j++) {
			dasindex = v.dasindexarr[j].replace(",", "��:");
			itembarcode = v.itembarcode.substring(0, 2) + "-" + v.itembarcode.substring(2, (v.itembarcode.length - 4)) + "-" + v.itembarcode.substring((v.itembarcode.length - 4));

			if (v.dasindexarr.length >= 5) {
				dasindex = "�ټ�:" + v.totalItemNo;
			}

			TEC_DO3.PrinterOpen();

			TEC_DO3.PrintText(30,0 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, v.titleStr);
			TEC_DO3.PrintText(10,55 + TOSHIBA_HEIGHTOFFSET, "Arial", 35*F, 0, 0, "R:"+v.itemrackcode);
			TEC_DO3.PrintText(10,100 + TOSHIBA_HEIGHTOFFSET, "Arial", 35*F, 0, 0, dasindex);
			TEC_DO3.PrintText(120,35 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, v.makerid);
			TEC_DO3.PrintText(120,65 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, itembarcode);
			TEC_DO3.PrintText(120,95 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, v.publicbarcode);

			if (v.dasindexarr.length < 5 && dasindex.split(":").length == 2 && dasindex.split(":")[1] > 1) {
				TEC_DO3.PrintText(200,125 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, dasindex.split(":")[1] + "��");
			} else {
				TEC_DO3.PrintText(200,125 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "--");
			}

			if (v.danpumynarr) {
				if (v.danpumynarr.length < 5 && v.danpumynarr[j].split(",").length == 2 && v.danpumynarr[j].split(",")[1] === "Y") {
					TEC_DO3.PrintText(120,125 + TOSHIBA_HEIGHTOFFSET, "Arial", 30*F, 0, 0, "��ǰ");
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

//innerbox �ε��� ���
function printTOSHIBAinnerboxLabel(shopid, shopname, packingdate, innerboxno, innerboxweight, prdcode, prdbarcode){
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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

//cartonbox �ε��� ���
function printTOSHIBAcartonboxLabel(shopid, shopname, packingdate, cartonboxno, cartonboxweight, prdcode, prdbarcode){
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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

// ���� �ε��� ���
function printTOSHIBAMultiIndexSudongLabel(msg,itemno,fontName) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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

// ���� ����ȣ �ε��� ���
function printTOSHIBARackIndexSudongLabel(msg,itemno,fontName, barcodeprintyn) {
	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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

// ����ȣ �ε��� ���
function printTOSHIBAMultiRackIndexLabel(arrObject) {
	var v;
	var msg, itemno, fontName;

	if (TEC_DO3.IsDriver != 1){
		alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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
