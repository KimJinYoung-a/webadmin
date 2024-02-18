function DrawBarcode(elementid) {
        var objstring = "";
        var e;

        objstring = '<object id="' + elementid + '" ';
        objstring = objstring + ' classid="clsid:A4F3A486-2537-478C-B023-F8CCC41BF29D" ';
        objstring = objstring + ' codebase="http://partner.10x10.co.kr/cab/tenbarShow.cab#version=1,0,0,3" ';
        objstring = objstring + ' width=113 ';
        objstring = objstring + ' height=15 ';
        objstring = objstring + ' align=top ';
        objstring = objstring + ' hspace=0 ';
        objstring = objstring + ' vspace=0 ';
        objstring = objstring + '></object> ';

        document.write(objstring);
}

function DrawBarcodeWithParam(elementid, iwidth, iheight) {
        var objstring = "";
        var e;

        objstring = '<object id="' + elementid + '" ';
        objstring = objstring + ' classid="clsid:A4F3A486-2537-478C-B023-F8CCC41BF29D" ';
        objstring = objstring + ' codebase="http://partner.10x10.co.kr/cab/tenbarShow.cab#version=1,0,0,3" ';
        objstring = objstring + ' width=' + iwidth;
        objstring = objstring + ' height=' + iheight;
        objstring = objstring + ' align=top ';
        objstring = objstring + ' hspace=0 ';
        objstring = objstring + ' vspace=0 ';
        objstring = objstring + '></object> ';

        document.write(objstring);
}

function DrawReceiptPrintobj(elementid){
		var objstring = "";
        var e;

        objstring = '<OBJECT name="' + elementid + '" ';
	  	objstring = objstring + ' classid="clsid:F280BDF9-0B2F-4713-A1DB-206ACF79804C" ';
	  	objstring = objstring + ' codebase="/TenBaljuPrintProj1.cab#version=1,0,0,0" ';
	  	objstring = objstring + ' width=1 ';
	  	objstring = objstring + ' height=1 ';
	  	objstring = objstring + ' align=center ';
	  	objstring = objstring + ' hspace=0 ';
	  	objstring = objstring + ' vspace=0 ';
		objstring = objstring + ' ></OBJECT> ';

		document.write(objstring);
}


function DrawReceiptPrintobj_TEC(elementid,printname){
        var objstring = "";
        var e;
        objstring = '<OBJECT name="' + elementid + '" ';
        objstring = objstring + ' classid="clsid:E76C9051-A8C4-458E-9F60-3C14DB9EECF9" ';
        objstring = objstring + ' codebase="http://billyman/Tec_dol.cab#version=1,5,0,0" ';
        objstring = objstring + ' width=0 ';
        objstring = objstring + ' height=0 ';
        objstring = objstring + ' align=center ';
        objstring = objstring + ' hspace=0 ';
        objstring = objstring + ' vspace=0 ';
        objstring = objstring + ' > ';
        objstring = objstring + ' <PARAM Name="PrinterName" Value="' + printname + '"> ';
        objstring = objstring + ' </OBJECT>';

        document.write(objstring);
}
