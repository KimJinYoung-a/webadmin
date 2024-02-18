function jsPopCalendar(fName,sName){
	var fd = eval("document."+fName+"."+sName);

    //calendar.js 포함 되어있으면
    if (typeof calPopup == "function"){
        var compname = 'document.' + fName + '.' + sName;
        calPopup(fd,'calendarPopup',20+80,0, compname,'');
    }else{
    	var winCal;
    	winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
    	winCal.focus();
    }
}

////////////////////////// 바코드 프린트 출력 팝업 //////////////////////////		'//2013.03.04 한용민 생성
//listgubun구분 : 상품리스트=ITEM , 주문리스트=JUMUN , 패킹리스트=PACKING , 업체주문리스트=UPCHEJUMUN
//idx구분 : 주문idx
//prdcode구분 : 물류코드

//온라인, 오프라인 선택 안하는 곳에서 사용
function printbarcode_on_off_multi(){
	var popwin = window.open('/common/barcode/barcodeprint_on_off_multi.asp','printbarcode_on_off_multi','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//온라인, 오프라인 선택 안하는 곳에서 사용 업체 전용
function printbarcode_on_off_multi_upche(){
	var popwin = window.open('/partner/common/barcode/barcodeprint_on_off_multi_pop.asp','printbarcode_on_off_multi','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//온라인 출력
function printbarcode_on(listgubun, tmp1, tmp2, tmp3, prdcode, makerid) {
	var popwin = window.open('/common/barcode/barcodeprint_on_off_multi.asp?onoffgubun=ONLINE&listgubun='+listgubun +'&prdcode='+prdcode+'&makerid='+makerid,'printbarcode_on','width=1280 height=960 scrollbars=yes resizable=yes');
	popwin.focus();
}

//오프라인 출력
function printbarcode_off(listgubun, upchemasteridx, prdcode, makerid, shopid, baljucode, jumunmasteridx, boxno, tmp1){
	var popwin = window.open('/common/barcode/barcodeprint_on_off_multi.asp?onoffgubun=OFFLINE&listgubun='+listgubun+'&ipchul='+upchemasteridx+'&prdcode='+prdcode+'&makerid='+makerid+'&shopid='+shopid+'&baljucode='+baljucode+'&masteridx=' + jumunmasteridx + '&boxno=' + boxno,'printbarcode_off','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//오프라인 출력 업체 전용
function printbarcode_off_upche(listgubun, upchemasteridx, prdcode, makerid, shopid, baljucode, jumunmasteridx, boxno, tmp1){
	var popwin = window.open('/partner/common/barcode/barcodeprint_on_off_multi_pop.asp?onoffgubun=OFFLINE&listgubun='+listgubun+'&ipchul='+upchemasteridx+'&prdcode='+prdcode+'&makerid='+makerid+'&shopid='+shopid+'&baljucode='+baljucode+'&masteridx=' + jumunmasteridx + '&boxno=' + boxno,'printbarcode_off_upche','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}
////////////////////////// 바코드 프린트 출력 팝업 //////////////////////////

function ViewOfflineOrderSheet(masteridx) {

	var popwin = window.open("/common/popOrderSheetView.asp?masteridx=" + masteridx + "&ordersheettype=offlineorder","ViewOfflineOrderSheet","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ExcelOfflineOrderSheet(masteridx) {

	var popwin = window.open("/common/popOrderSheetExcel.asp?masteridx=" + masteridx + "&ordersheettype=offlineorder","ExcelOfflineOrderSheet","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

//범용버전 		//20120831 한용민 생성
function ExcelOfflineOrderSheetpublic(masteridx) {

	var ExcelOfflineOrderSheetpublic = window.open("/common/popOrderpublicSheetExcel.asp?masteridx=" + masteridx + "&ordersheettype=offlineorder","ExcelOfflineOrderSheetpublic","width=800 height=600 scrollbars=yes resizable=yes");
	ExcelOfflineOrderSheetpublic.focus();
}

function fnTrim(orgStr){
    return orgStr.replace(/(^\s*)|(\s*$)/gi, "");
}

function phone_format(obj) {
	var tmp;

	tmp = obj.value;

	tmp = tmp.replace(/\-/g, "");

	if (isNaN(tmp) == true) {
		alert("전화번호에는 숫자이외에 입력할 수 없습니다.");
		obj.value = "";
		obj.focus();
		return;
	}

	if (tmp.length <= 4) {
		obj.value = tmp;
	} else if (tmp.length <= 8) {
		obj.value = tmp.replace(/([0-9]+)([0-9]{4})/,"$1-$2");
	} else {
		obj.value = tmp.replace(/(^02.{0}|^01.{1}|^070|[0-9]{3})([0-9]+)([0-9]{4})/,"$1-$2-$3");
	}
}

function fnCheckAll(bool, comp){
    var frm = comp.form;

    if (!comp.length){
        comp.checked = bool;
        AnCheckClick(comp);
    }else{
        for (var i=0;i<comp.length;i++){
            comp[i].checked = bool;
            AnCheckClick(comp[i]);
        }
    }
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

//상품 배송소요일 LIST
function popItemAvgDlvList(itemid){
    //return;
    var popwin = window.open("/admin/datamart/baesong/iframe_baesong_term_list.asp?itemid=" + itemid,"popItemAvgDlvList","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

//상품 배송소요일 graph
function popItemAvgDlvGraph(itemid){
    //return;
    var popwin = window.open("/admin/datamart/baesong/iframe_baesong_term_graph.asp?itemid=" + itemid,"popItemAvgDlvGraph","width=900 height=500 scrollbars=yes resizable=yes");
	popwin.focus();
}

//브랜드 ID 검색 팝업창
function jsSearchBrandID(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsSearchBrandIDwithUserDIV(frmName,compName, userdiv) {
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal + "&userdiv=" + userdiv,"jsSearchBrandIDwithUserDIV","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsSearchMeachulID(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popMeachulIDSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"jsSearchMeachulID","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsSearchBrandID2(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popBrandSearch2.asp?frmName=" + frmName + "&compName=" + compName + "&rect=","popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsSearchBrandIDchgMargin(frmName,compName,mgnfName,evtjs){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popBrandSearch_chgMargin.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal + "&mgnfName=" + mgnfName + "&evtjs=" + evtjs,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function PopBrandInfoEdit(makerid){
    var popwin = window.open("/admin/member/popbrandinfoonly.asp?designer=" + makerid,"popbrandinfoonly","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheInfoEdit(groupid){
    var popwin = window.open("/admin/member/popupcheinfoonly.asp?groupid=" + groupid,"popupcheinfoonly","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheBrandInfoEdit(makerid){
    var popwin = window.open("/admin/member/popupchebrandinfo.asp?designer=" + makerid,"popupchebrandinfoedit","width=660 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopBrandAdminUsingChange(v){
	var popwin = window.open("/admin/member/popbrandadminusing.asp?designer=" + v,"popbrandadminusing","width=1200 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheSelect(frmname, params) {
	var url = '/admin/member/popupcheselect.asp?frmname=' + frmname;
	if (params != undefined) {
		url = url + '&' + params;
	}

	var popwin = window.open(url,"popupcheselect","width=800 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheSelect_shop(frmname,shopyn){
	var PopUpcheSelect_shop = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&shopyn="+shopyn,"PopUpcheSelect_shop","width=800 height=580 scrollbars=yes resizable=yes");
	PopUpcheSelect_shop.focus();
}

function TnPopItemStocknew(itemgubun, itemid, itemoption){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun="+ itemgubun +"&itemid=" + itemid + "&itemoption=" + itemoption,"popitemstocklist","width=1024 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}

function TnPopItemStock(itemid,itemoption){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid + "&itemoption=" + itemoption,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopItemIpChulList(fromdate,todate,itemgubun,itemid,itemoption,ipchulflag){
	var popwin = window.open('/common/pop_stock_ipgo.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag, 'pop_stock_ipgo', 'width=800,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function PopItemIpChulListOffLine(fromdate,todate,itemgubun,itemid,itemoption,ipchulflag, shopid){
	var popwin = window.open('/common/pop_ipgolist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid, 'pop_stock_ipgo_off', 'width=1000,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function calendarOpen4(objTarget,caption,defaultval){
	 //var objTarget = document.getElementById(targetName);
	 output = window.showModalDialog("/lib/calendar.asp?caption=" + caption + "&defaultval=" + defaultval , null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 if(output!=''){
	  	objTarget.value=output;
	  	return true;
	 }else{
		return false;
	 }
}


function calendarOpen3(objTarget,caption,defaultval){
	 //var objTarget = document.getElementById(targetName);
	 output = window.showModalDialog("/lib/calendar.asp?caption=" + caption + "&defaultval=" + defaultval , null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 if(output!=''){
	  	objTarget.value=output;
	 }else{

	 }
}

function calendarOpen2(objTarget){
	 //var objTarget = document.getElementById(targetName);
	 output = window.showModalDialog("/lib/calendar.html" , null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 if(output!=''){
	  	objTarget.value=output;
	  	return true;
	 }else{
	 	return false;
	 }
}

function calendarOpen(objTarget){
/*
	 output = window.showModalDialog("/lib/calendar.html" , null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 if(output!=''){
	  	objTarget.value=output;
	 }else{

	 }
*/
    //calendar.js 포함 되어있으면
    if (typeof calPopup == "function"){
        var compname = 'document.' + objTarget.form.name + '.' + objTarget.name;
        calPopup(objTarget,'calendarPopup',50,50, compname,'');
    }else{
        var fName = objTarget.form.name;
        var sName = objTarget.name;
    	var winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
    	winCal.focus();
    }
}

function AnModiNormal(frm){
	var ret = confirm('수정하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function AnDelNormal(frm){
	var ret = confirm('삭제 하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function AnPopItemList(){
	window.open('/module/popitemlistbydesigner.asp', 'itemlist', 'width=600,height=660,location=no,menubar=no,resizable=no,scrollbars=no,status=no,toolbar=no');
}

function CheckNDolimitsell(frm){
	if (!IsDigit(frm.limitno.value)){
		alert('한정수량은 숫자만 가능합니다.');
		frm.limitno.focus();
		return false;
	}

	if (!IsDigit(frm.limitsold.value)){
		alert('판매된 수량은 숫자만 가능합니다.');
		frm.limitsold.focus();
		return false;
	}

	if (frm.baedalcd.value.length<1){
		alert('배송구분을 선택하세요..');
		frm.baedalcd.focus();
		return false;
	}

	var ret = confirm('저장하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function CheckNDobuyprice(frm){
	if (!IsDigit(frm.sellcash.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.sellcash.focus();
		return false;
	}

	if (!IsDigit(frm.sellvat.value)){
		alert('판매Vat는 숫자만 가능합니다.');
		frm.sellvat.focus();
		return false;
	}

	if (!IsDigit(frm.buycash.value)){
		alert('구매가는 숫자만 가능합니다.');
		frm.buycash.focus();
		return false;
	}

	if (!IsDigit(frm.buyvat.value)){
		alert('구매Vat는 숫자만 가능합니다.');
		frm.buyvat.focus();
		return false;
	}

	if (!IsDigit(frm.marginrate.value)){
		alert('마진율은 숫자만 가능합니다.');
		frm.marginrate.focus();
		return false;
	}

	if ((frm.marginrate.value<0)&&(frm.marginrate.value>100)){
		alert('마진율은 0~100만 가능합니다.');
		frm.marginrate.focus();
		return false;
	}
	var ret = confirm('저장하시겠습니까?');

	return ret;
}

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}

function hL(E){
	while (E.tagName!="TR")
	{
		E=E.parentElement;
	}

	E.className = "H";
}

function dL(E){
	while (E.tagName!="TR"){
		E=E.parentElement;
	}

	E.className = "";
}

function IsInteger(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("-0123456789".indexOf(v.charAt(j)) < 0) {
			return false;
		}

		//if ((v.charAt(j) * 0 == 0) == false){
		//	return false;
		//}
	}
	return true;
}

function IsDigit(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("0123456789".indexOf(v.charAt(j)) < 0) {
			return false;
		}

		//if ((v.charAt(j) * 0 == 0) == false){
		//	return false;
		//}
	}
	return true;
}

function IsDouble(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("0123456789.".indexOf(v.charAt(j)) < 0) {
			return false;
		}
	}
	return true;
}

function IsNumbers(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("-0123456789.,".indexOf(v.charAt(j)) < 0) {
			return false;
		}
	}
	return true;
}

function AnCheckNBalju(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 주문이 없습니다.');
		return;
	}

	var ret = confirm('선택 주문을 새 발주서로 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderserial.value = upfrm.orderserial.value + "|" + frm.orderserial.value;
					upfrm.sitename.value = upfrm.sitename.value + "|" + frm.sitename.value;
				}
			}
		}
		upfrm.submit();
	}
}

// 오프라인발주서작성
function AnCheckNBaljuOffLine(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 주문이 없습니다.');
		return;
	}

	var ret = confirm('선택 주문을 새 발주서로 저장하시겠습니까?');
	if (ret){
		upfrm.orderidx.value = "";
		upfrm.baljucode.value = "";
		upfrm.baljuid.value = "";

		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderidx.value = upfrm.orderidx.value + "|" + frm.orderidx.value;
					upfrm.baljucode.value = upfrm.baljucode.value + "|" + frm.baljucode.value;
					upfrm.baljuid.value = upfrm.baljuid.value + "|" + frm.baljuid.value;
				}
			}
		}
		upfrm.submit();
		// alert(upfrm.baljuid.value);
	}
}

function AnAllCalcu(frm){
	var frmtarget;

	if (!IsDouble(frm.mgall.value)){
		alert('마진율은 숫자만 가능합니다.');
		frm.mgall.focus();
		return;
	}

	if ((frm.mgall.value<0)&&(frm.mgall.value>100)){
		alert('마진율은 0~100만 가능합니다.');
		frm.mgall.focus();
		return;
	}

	//if ((!frm.rdall[0].checked)&&(!frm.rdall[1].checked)){
	//	alert('부가세 포함여부를 선택하세요.');
	//	return;
	//}

	for (var i=0;i<document.forms.length;i++){
		frmtarget = document.forms[i];
		if (frmtarget.name.substr(0,9)=="frmBuyPrc") {
			if (!frmtarget.cksel.checked) continue;
			frmtarget.marginrate.value = frm.mgall.value;
			//if (frm.rdall[0].checked){
			//	frmtarget.vtinclude[0].checked=true;
				AnAutoCalcu(frmtarget,true);
			//}else{
			//	frmtarget.vtinclude[1].checked=true;
			//	AnAutoCalcu(frmtarget,false);
			//}
		}
	}
}

function AnAutoCalcu(frm,bool){
	var buf;
	if (!IsDigit(frm.sellcash.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDouble(frm.marginrate.value)){
		alert('마진율은 숫자만 가능합니다.');
		frm.marginrate.focus();
		return;
	}

	if ((frm.marginrate.value<0)&&(frm.marginrate.value>100)){
		alert('마진율은 0~100만 가능합니다.');
		frm.marginrate.focus();
		return;
	}

	if (bool){
		frm.sellvat.value = parseInt(Math.round(frm.sellcash.value/11));
		buf = parseInt(Math.round(frm.sellcash.value*(1-frm.marginrate.value/100.0)));
		frm.buycash.value = buf;
		frm.buyvat.value = parseInt(Math.round(buf/11));
		frm.tmpbuycash.value = parseInt(Math.round(buf-frm.buyvat.value));
		//frm.buyvat.value = Math.floor(buf/11);
		//frm.tmpbuycash.value = Math.floor(buf-frm.buyvat.value);
		//frm.buycash.value = Math.floor(frm.buyvat.value * 1 + frm.tmpbuycash.value * 1);

	}else{
		frm.sellvat.value = parseInt(Math.round(frm.sellcash.value/11));
		frm.tmpbuycash.value = parseInt(Math.round(frm.sellcash.value*(1-frm.marginrate.value/100)));
		frm.buycash.value = parseInt(Math.round(frm.tmpbuycash.value*1.1));
		frm.buyvat.value = parseInt(Math.round(frm.tmpbuycash.value*0.1));
		//frm.tmpbuycash.value = Math.floor(frm.sellcash.value*(1-frm.marginrate.value/100));
		//frm.buyvat.value = Math.floor(frm.buycash.value/10);
		//frm.buycash.value = Math.floor(frm.buyvat.value * 1  + frm.tmpbuycash.value * 1 );
	}
}

function AnDesignerSearchSaveFrame(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 업체가 없습니다.');
		return;
	}

	var ret = confirm('선택 업체를 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.designerid.value = upfrm.designerid.value + "|" + frm.designerid.value;

					if (frm.isusing[0].checked){
						upfrm.isusing.value = upfrm.isusing.value + "|" + "Y";
					}else{
						upfrm.isusing.value = upfrm.isusing.value + "|" + "N";
					}

					if (frm.isextusing[0].checked){
						upfrm.isextusing.value = upfrm.isextusing.value + "|" + "Y";
					}else{
						upfrm.isextusing.value = upfrm.isextusing.value + "|" + "N";
					}

					if (frm.isb2b[0].checked){
						upfrm.isb2b.value = upfrm.isb2b.value + "|" + "Y";
					}else{
						upfrm.isb2b.value = upfrm.isb2b.value + "|" + "N";
					}
				}
			}
		}
		frm.submit();
	}
}



function CheckNDoNormal(frm){
	var ret = confirm('저장하시겠습니까?');
	if (ret){
		return true;//frm.submit();
	}
	return false;
}

function CheckNDoitemviewset(frm){
	var ret = confirm('저장하시겠습니까?');
	if (ret){
		frm.submit();
	}
}


function AnItemlimitsellSaveFrame(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}

	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;
					upfrm.baedalcd.value = upfrm.baedalcd.value + "|" + frm.baedalcd.value;
					upfrm.limitno.value = upfrm.limitno.value + "|" + frm.limitno.value;
					upfrm.limitsold.value = upfrm.limitsold.value + "|" + frm.limitsold.value;

					if (frm.limityn[0].checked){
						upfrm.limityn.value = upfrm.limityn.value + "|" + "Y";
					}else{
						upfrm.limityn.value = upfrm.limityn.value + "|" + "N";
					}
				}
			}
		}
		frm.submit();
	}

}

function AnBuyPriceSaveFrame(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}

	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;
					upfrm.sellcash.value = upfrm.sellcash.value + "|" + frm.sellcash.value;
					upfrm.sellvat.value = upfrm.sellvat.value + "|" + frm.sellvat.value;
					upfrm.buycash.value = upfrm.buycash.value + "|" + frm.buycash.value;
					upfrm.buyvat.value = upfrm.buyvat.value + "|" + frm.buyvat.value;
					upfrm.marginrate.value = upfrm.marginrate.value + "|" + frm.marginrate.value;

					if (frm.vtinclude[0].checked){
						upfrm.vtinclude.value = upfrm.vtinclude.value + "|" + "Y";
					}else{
						upfrm.vtinclude.value = upfrm.vtinclude.value + "|" + "N";
					}
				}
			}
		}
		frm.submit();
	}
}

function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}


function AnSelectAll(frm,bool){
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
			e.checked = bool;
		}
	}
}

function AnBaljuDetail(frm){
	if (frm.baljumaster.value.length<1){
		alert('발주서를 선택하세요');
		return;
	}
	frm.submit();
}

function AnBaljuSearch(frm){
	frm.baljumaster.value="";
	frm.submit();
}

function AnDeliverRead(iid,frm){
	frm.id.value = iid;
	frm.action = 'bct_admin_deliver_read.asp';
	frm.submit();
}


function AnCheckDelivery(frm){
	var v = confirm('변경하시겠습니까?');
	if (v) {
		frm.mode.value='check';
		frm.submit();
	}
}

function AnDeleteDelivery(frm){
	var v = confirm('삭제하시겠습니까?');
	if (v) {
		frm.mode.value='del';
		frm.submit();
	}
}

function AnWriteDeliveryCom(frm){
	if (frm.tx_com.value.length<1) {
		alert('내용을 입력하세요.');
		frm.tx_com.focus();
		return;
	}

	if (frm.writer.value=='') {
		alert('글쓴이를 선택하세요.');
		return;
	}

	frm.submit();
}

function AnDeliveryWrite(frm){
	if (frm.sitename.value=='') {
		alert('사이트를 선택하세요.');
		return;
	}

	if (frm.buyname.value.length<1) {
		alert('고객명을 입력하세요.');
		frm.buyname.focus();
		return;
	}

	if (frm.orderserial.value.length<1) {
		alert('주문번호를 입력하세요.');
		frm.orderserial.focus();
		return;
	}

	if (frm.writer.value=='') {
		alert('글쓴이를 선택하세요.');
		return;
	}

	if (frm.title.value.length<1) {
		alert('타이틀을 입력하세요.');
		frm.title.focus();
		return;
	}

	if (frm.txmemo.value.length<1) {
		alert('메모를 입력하세요.');
		frm.txmemo.focus();
		return;
	}
	frm.submit();
}

function TnUnderConstruction(){
	alert('UnderConstruction..');
}

function SvNoticeConfirm(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret==true) {
		frm.submit();
	}
}


function checkdate3(form){

        var year=form.yyyy2.value;
        var month=form.mm2.value;
        var cal;

                var lastdate=new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

					if((0==year%4&&0!=year%100)||0==year%400){
					  lastdate[1]=29;
					}

 		if(form.dd2.value >= (lastdate[month-1]+1)){
		  alert(month+"월은" + lastdate[month-1] + "일 까지 입니다.\n\n"+lastdate[month-1]+"일 이전날짜로 검색해주세요!");
		}

}

function CheckDateValid(yyyy, mm, dd) {
	var year = yyyy;
    var month = mm;
	var day = dd;

    var lastdate = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

	if(((0 == (year % 4)) && (0 != (year % 100))) || (0 == (year % 400))) {
		lastdate[1] = 29;
	}

 	if (day >= (lastdate[month-1]+1)) {
		alert("에러!!\n\n" + month + " 월은 " + lastdate[month-1] + " 일 까지 입니다.");
		return false;
	}

	return true;
}

function ResetValidDate(yyyy, mm, dd, svalue) {
	var year = yyyy;
    var month = mm;
	var day = dd;

    var lastdate = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

	if(((0 == (year % 4)) && (0 != (year % 100))) || (0 == (year % 400))) {
		lastdate[1] = 29;
	}

 	if (day >= (lastdate[month-1]+1)) {
		eval("document.all."+svalue).value = lastdate[month-1];
	}

}

function AnItemviewsetSaveAll(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}

	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemidlist.value = upfrm.itemidlist.value + "|" + frm.itemid.value;
					upfrm.tingpointlist.value = upfrm.tingpointlist.value + "|" + frm.tingpoint.value;
					upfrm.tingpoint_blist.value = upfrm.tingpoint_blist.value + "|" + frm.tingpoint_b.value;
					upfrm.userclasslist.value = upfrm.userclasslist.value + "|" + frm.userclass.value;
					upfrm.limitdivlist.value = upfrm.limitdivlist.value + "|" + frm.limitdiv.value;
					upfrm.limitealist.value = upfrm.limitealist.value + "|" + frm.limitea.value;
					upfrm.selectitemlist.value = upfrm.selectitemlist.value + "|" + frm.selectitem.value;
				}
			}
		}
		frm.submit();
	}
}


function GetByteLength(val){
 	var real_byte = val.length;
 	for (var ii=0; ii<val.length; ii++) {
  		var temp = val.substr(ii,1).charCodeAt(0);
  		if (temp > 127) { real_byte++; }
 	}

   return real_byte;
}

//-----------2009 추가 ------------
//입력 키 값이 숫자인지 체크
function onlyNumberInput()
{
	var code = window.event.keyCode;
	if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) {
		window.event.returnValue = true;
		return;
	}
	window.event.returnValue = false;
}
function jsTrim(str){
    return str.replace(/^\s\s*/, '').replace(/\s\s*$/, '');
}

// 숫자만 입력받기 IE, FF
function onlyNumber(obj,evt) {
	var evCode = (window.netscape) ? evt.which : event.keyCode ;
	if (isNumeric(evCode)==false)
	{
		alert("숫자만 입력이 가능합니다.");
		if (!window.netscape)
			event.returnValue=false;
		else
			obj.value = obj.value.substr(0,obj.value.length-1);
	}
}

// 숫자와 구분자만 입력받기 IE, FF
function onlyNumberDot(obj,evt) {
	var evCode = (window.netscape) ? evt.which : event.keyCode ;
	if (isNumericDot(evCode)==false)
	{
		alert("숫자만 입력이 가능합니다.");
		if (!window.netscape)
			event.returnValue=false;
		else
			obj.value = obj.value.substr(0,obj.value.length-1);
	}
}

// 숫자인지 체크
function isNumeric( value )
{
	if (value == 189 || value == 109 || value == 8 || value == 9 || value == 13 || (value >= 37 && value <= 40) || value == 46 || value == 27 || value == 116 || (value >= 48 && value <= 57) || (value >= 96 && value <= 105))
		return true;
	else
		return false;
}

// 숫자나 구분자인지 체크
function isNumericDot( value ) {
	if (value == 189 || value == 109 || value == 8 || value == 9 || value == 13 || (value >= 37 && value <= 40) || value == 46 || (value >= 48 && value <= 57) ||  value == 110 || value == 190 || value == 192 || value == 27 || value == 116 || (value >= 96 && value <= 105) || value ==17 || value ==16 || value ==186 || value ==188)
		return true;
	else
		return false;
}

// 폼 필수 필드 유효성 체크
function validField(obj, msg, len)
{
	if (obj.length > 1)
	{
		if (obj[0].type == "radio" || obj[0].type == "checkbox")
		{
			var chk = 0;
			for (var i = 0; i < obj.length; i++)
				if (obj[i].checked)
					chk++;

			if (chk==0)
			{
				if (obj[0].type == "checkbox")
					alert("" + msg + " 하나 이상 체크해주세요.");
				else
					alert("" + msg + " 체크해주세요.");

				obj[0].focus();
				return false;
			}
		}
		else if (obj.type == "select-one")
		{
			if(jsTrim(obj.value) == "")
			{
				alert("" + msg + " 선택해주세요.");
				obj.focus();
				return false;
			}
		}
	}
	else if (obj.type == "radio" || obj.type == "checkbox")
	{
		if (obj.checked==false)
		{
			alert("" + msg + " 체크해주세요.");
			obj.focus();
			return false;
		}
	}
	else
	{
		if(jsTrim(obj.value) == "")
		{
			alert("" + msg + " 입력해주세요.");
			obj.focus();
			return false;
		}
		if (len)
		{
			if (returnByteCount(obj.value) > len)
			{
				alert("" + msg + " 한글기준 "+parseInt(len/2)+"자, 영문기준 "+len+"자 이내로 해주세요.");
				obj.focus();
				return false;
			}
		}
	}

	return true;
}

// 자리수 형식 체크
function validFieldLeng(obj, msg, len)
{
	if(obj.value.length < len)	// 작으면 false 같거나 크면 상관없음 maxlength로 제한하기 때문에 클 수는 없음
	{
		alert(msg + " 자리수가 맞지 않습니다.\n"+len+" 자리 또는 그 이상으로 입력해주세요.");
		obj.focus();
		return false;
	}
	else
		return true;
}

// 바이트수를 리턴하는 함수
function returnByteCount(val)
{
	var len = val.length;            //현재 value값의 글자 수
	var cnt = 0;                    //한글일 경우 2 그외에는 1바이트 수 저장
	var chr = "";                 //현재 한/영 체크할 letter를 저장

	for (i=0; i<len; i++)
	{
		chr = val.charAt(i);

		// 체크문자가 한글일 경우 2byte 그 외의 경우 1byte 증가
		if (escape(chr).length > 4)
		   cnt += 2;
		else
		   cnt++;
	}
	return cnt;
}

// 폼 필드값 리턴, 조건에 따라 개수가 달라질때
function getFieldValue(obj)
{
	var ret = "";
	if (obj.length > 1)
	{
		if (obj[0].type == "radio" || obj[0].type == "checkbox")
		{
			for (var i = 0; i < obj.length; i++)
				if (obj[i].checked)
					if (ret=="")
						ret = obj[i].value;
					else
						ret += "," + obj[i].value;
		}
		else if (obj.type == "select-one")
		{
			ret = obj.value;
		}
	}
	else
	{
		ret = obj.value;
	}

	return ret;
}

// 팝업창 자동리사이즈, Width를 지정하면 지정한대로
function popupResize(innerWidth)
{
	var strAgent = navigator.userAgent.toLowerCase();
	var strVersion = strAgent.substr(strAgent.indexOf("msie")+5,1);
    var IE	= strAgent.indexOf("MSIE") ?	true : false;
	if (IE)
	{
		var addHeight = (strAgent >=  7) ? 70 : 55;	// 7 이상은 URL창크기만큼 추가

		var innerBody = document.body;
		var innerHeight = innerBody.scrollHeight + (innerBody.offsetHeight - innerBody.clientHeight);
		if (!innerWidth)
			var innerWidth = innerBody.scrollWidth + (innerBody.offsetWidth - innerBody.clientWidth);

		innerWidth += 10;
		innerHeight += addHeight;
		window.resizeTo(innerWidth,innerHeight);
	}
	else					// FF
	{
		var Dwidth = parseInt(document.body.scrollWidth);
		var Dheight = parseInt(document.body.scrollHeight);
		var divEl = document.createElement("div");
		divEl.style.position = "absolute";
		divEl.style.left = "0px";
		divEl.style.top = "0px";
		divEl.style.width = "100%";
		divEl.style.height = "100%";
	    document.body.appendChild(divEl);
	    window.resizeBy(Dwidth-divEl.offsetWidth, Dheight-divEl.offsetHeight);
		document.body.removeChild(divEl);
	}
}


// iframe 길이 자동
function resizeIfr(obj, minHeight) {
	minHeight = minHeight || 10;

	try {
		var getHeightByElement = function(body) {
			var last = body.lastChild;
			try {
				while (last && last.nodeType != 1 || !last.offsetTop) last = last.previousSibling;
				return last.offsetTop+last.offsetHeight;
			} catch(e) {
				return 0;
			}

		}

		var doc = obj.contentDocument || obj.contentWindow.document;
		if (doc.location.href == 'about:blank') {
			obj.style.height = minHeight+'px';
			return;
		}

		//var h = Math.max(doc.body.scrollHeight,getHeightByElement(doc.body));
		//var h = doc.body.scrollHeight;
		if (/MSIE/.test(navigator.userAgent)) {
			var h = doc.body.scrollHeight;
		} else {
			var s = doc.body.appendChild(document.createElement('DIV'))
			s.style.clear = 'both';

			var h = s.offsetTop;
			s.parentNode.removeChild(s);
		}

		//if (/MSIE/.test(navigator.userAgent)) h += doc.body.offsetHeight - doc.body.clientHeight;
		if (h < minHeight) h = minHeight;

		obj.style.height = h + 'px';
		if (typeof resizeIfr.check == 'undefined') resizeIfr.check = 0;
		if (typeof obj._check == 'undefined') obj._check = 0;

//		if (obj._check < 5) {
//			obj._check++;
			setTimeout(function(){ resizeIfr(obj,minHeight) }, 200); // check 5 times for IE bug
//		} else {
			//obj._check = 0;
//		}
	} catch (e) {
		//alert(e);
	}

}


//동영상 SerialNumber 검색 팝업창
function jsSearchVideoSn(frmName,compName,vDiv){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/sitemaster/popVideoSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal+ "&vDiv=" + vDiv,"popVideoSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}



function plusComma(num){
	if (num < 0) { num *= -1; var minus = true}
	else var minus = false

	var dotPos = (num+"").split(".")
	var dotU = dotPos[0]
	var dotD = dotPos[1]
	var commaFlag = dotU.length%3

	if(commaFlag) {
		var out = dotU.substring(0, commaFlag)
		if (dotU.length > 3) out += ","
	}
	else var out = ""

	for (var i=commaFlag; i < dotU.length; i+=3) {
		out += dotU.substring(i, i+3)
		if( i < dotU.length-3) out += ","
	}

	if(minus) out = "-" + out
	if(dotD) return out + "." + dotD
	else return out
}

//바코드 OCX
function drawTTPprintOcx(iname, iversion){
    var iObjStr = "";
    iObjStr = "<OBJECT"
    iObjStr = iObjStr + "      name='" + iname + "'";
    iObjStr = iObjStr + "	  classid='clsid:E0CFD990-7055-4BFD-927E-BF5553AF0F54'";
    iObjStr = iObjStr + "	  codebase='http://webadmin.10x10.co.kr/common/cab/tenTTPBarPrint.cab#version=" + iversion + "'";
    iObjStr = iObjStr + "	  width=0";
    iObjStr = iObjStr + "	  height=0";
    iObjStr = iObjStr + "	  align=center";
    iObjStr = iObjStr + "	  hspace=0";
    iObjStr = iObjStr + "	  vspace=0";
    iObjStr = iObjStr + ">";
    iObjStr = iObjStr + "</OBJECT>";

    document.write(iObjStr);
}

function drawTTPprintOcxV2(iname, iversion){
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

function drawSrp350PlotOcx(iname, iversion){
    var iObjStr = "";
    iObjStr = "<OBJECT"
    iObjStr = iObjStr + "     id='" + iname + "' name='" + iname + "'";
    iObjStr = iObjStr + "	  classid='clsid:5DC34DA8-9C0F-4A43-B772-78090A204600'";
    iObjStr = iObjStr + "	  codebase='http://webadmin.10x10.co.kr/common/cab/TnSRPPlot.cab#version=" + iversion + "'";
    iObjStr = iObjStr + "	  width=0";
    iObjStr = iObjStr + "	  height=0";
    iObjStr = iObjStr + "	  align=center";
    iObjStr = iObjStr + "	  hspace=0";
    iObjStr = iObjStr + "	  vspace=0";
    iObjStr = iObjStr + ">";
    iObjStr = iObjStr + "</OBJECT>";

    document.write(iObjStr);
}

//우편번호 입력
function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".zipcode").value = post1 + "-" + post2;

    eval(frmname + ".zipaddr").value = addr;
    eval(frmname + ".useraddr").value = dong;
}

// 날짜크기 비교
function dateChkComp(dt1,dt2)
{
	//구분자로 나누어 배열로 변환
	v0=dt1.split("-");
	v1=dt2.split("-");

	//일자에 해당하는 타임스탬프로 변환
	v0=new Date(v0[0],v0[1],v0[2]).valueOf();
	v1=new Date(v1[0],v1[1],v1[2]).valueOf();

	//일차이를 구한뒤 하루에 해당하는 값으로 곱하여, 초단위를 일단위로 변환
	cha=(v1-v0)/(1000*60*60*24);

	if(cha>0)
		return true;
	else
		return false;
}

//매장직원 매장 지정
function shopreg(empno){
	var shopreg = window.open('/common/offshop/member/shopuser_reg.asp?empno='+empno,'shopreg','width=800,height=400,scrollbars=yes,resizable=yes');
	shopreg.focus();
}

//asp instr 함수 생성	//2011.02.11 한용민 생성
function instr(strsearch,charsearchfor){
	for ( var i =0; i < strsearch.length; i++){
		if (charsearchfor == mid(strsearch , i ,1) ){
			return i;
		}
	}
	return -1;
}

//asp mid 함수생성	//2011.02.11 한용민 생성
function mid(str ,start,len){
	if (start < 0 || len < 0) return '';
	var iend;
	var ilen = str.length ;

	if( start + len > ilen){
		iend = ilen;
	} else {
		iend = start + len;
	}

	return str.substring(start,iend);
}

//문자열의 공백여부 체크 2011.03.16 정윤정 생성
function jsChkBlank(str)
{
    if (str == "" || str.split(" ").join("") == ""){
        return true;
	}
    else{
        return false;
	}
}

//상품 미리보기
function jsGoPreItem(wwwURL, itemid){
	 window.open('about:blank').location.href = wwwURL+"/shopping/category_prd.asp?itemid="+itemid;
}

// 업체 어드민 우편번호 찾기 		//2016.07.04 한용민 생성
// 사용안함. 사용하지 말것.
function TnFindZipNewdesigner(frmname, strMode){
    var TnFindZipNewdesigner = window.open('/designer/lib/searchzip_new.asp?target=' + frmname + '&strMode='+strMode, 'TnFindZipNewdesigner', 'width=580,height=690,left=400,top=200,scrollbars=yes,resizable=yes');
    TnFindZipNewdesigner.focus();
}

//2016.07.04 한용민 생성
function TnFindZipNew(frmname, strMode){
	// 국토부 통신버전 http://www.juso.go.kr/addrlink/addrLinkApiJsonp.do
    var TnFindZipNew = window.open('/lib/searchzip_new.asp?target=' + frmname + '&strMode='+strMode, 'TnFindZipNew', 'width=700,height=768,left=400,top=200,scrollbars=yes,resizable=yes');
    TnFindZipNew.focus();
}

//업체 어드민 우편번호 찾기 		//2016.11.21 정윤정 생성
function TnFindZipNewPartner(frmname, strMode){
    var TnFindZipNewPartner = window.open('/partner/lib/searchzip_new.asp?target=' + frmname + '&strMode='+strMode, 'TnFindZipNewPartner', 'width=580,height=690,left=400,top=200,scrollbars=yes,resizable=yes');
    TnFindZipNewPartner.focus();
}

//2019.07.30 한용민 생성
function FnFindZipNew(frmname, strMode){
	// 다음카카오 통신버전 https://ssl.daumcdn.net/dmaps/map_js_init/postcode.v2.js
	var FnFindZipNew = window.open('/lib/searchzip_ka.asp?target=' + frmname + '&strMode='+strMode, 'FnFindZipNew', 'width=700,height=768,left=400,top=200,scrollbars=yes,resizable=yes');
    FnFindZipNew.focus();
}

//2019.07.31 한용민 생성
function FnFindZipNewPartner(frmname, strMode){
	// 다음카카오 통신버전 https://ssl.daumcdn.net/dmaps/map_js_init/postcode.v2.js
	var FnFindZipNew = window.open('/partner/lib/searchzip_ka.asp?target=' + frmname + '&strMode='+strMode, 'FnFindZipNew', 'width=700,height=768,left=400,top=200,scrollbars=yes,resizable=yes');
    FnFindZipNew.focus();
}

//left 함수 구현	//2016.12.09 한용민 생성
function left(str, n){
if (n <= 0)
    return "";
else if (n > String(str).length)
    return str;
else
    return String(str).substring(0,n);
}

//Right 함수 구현	//2016.12.09 한용민 생성
function right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}

// 패스워드 복잡도 검사		//2017.09.25 한용민 수정
function fnChkComplexPassword(pwd) {
    var aAlpha = /[a-z]|[A-Z]/;
    var aNumber = /[0-9]/;
    var aSpecial = /[!|@|#|$|%|^|&|*|(|)|-|_]/;
    var sRst = true;

    if(pwd.length < 8){
        sRst=false;
        return sRst;
    }

    var numAlpha = 0;
    var numNums = 0;
    var numSpecials = 0;
    for(var i=0; i<pwd.length; i++){
        if(aAlpha.test(pwd.substr(i,1)))
            numAlpha++;
        else if(aNumber.test(pwd.substr(i,1)))
            numNums++;
        else if(aSpecial.test(pwd.substr(i,1)))
            numSpecials++;
    }

	// 3가지 조합
    if( (numAlpha>0 && numNums>0 && numSpecials>0) ) {
    	if (pwd.length >= 8){
    		sRst=true;
    	}else{
    		sRst=false;
    	}

	// 2가지 조합
    } else if((numAlpha>0 && numNums>0)||(numAlpha>0 && numSpecials>0)||(numNums>0 && numSpecials>0)) {
    	if (pwd.length >= 10){
    		sRst=true;
    	}else{
    		sRst=false;
    	}

    } else {
    	sRst=false;
    }
    return sRst;
}

//trim 함수 구현	//2017.12.11 한용민 생성
String.prototype.ltrim = function() { return this.replace(/^\s+/,""); }
String.prototype.rtrim = function() { return this.replace(/\s+$/,""); }

////////////////////////// 전안법 안전인증(업체쪽 상품등록쪽은 저장단에도 체크 로직이 있음) /////////////// 2018.03.23 한용민 생성
// 안전인증 유효성 검사. false 로 하면 scm 전체에서 유효성 체크 안함.
var _isSafetyCheck = false;

// 안전인증 등록 체크
function jsSafetyCheck(orgsafetyYn,temp2){
	// 유효성 검사 사용여부
	if ( !_isSafetyCheck ) return false;

	// 상품고시 품목
	var infoDiv = itemreg.infoDiv.value;
	if (infoDiv=='' || infoDiv==undefined) return false;

	// 우선 선택 가능한 값은 모두 안보이게 리셋
	for (var i=0; i < itemreg.safetyYn.length; i++){
		itemreg.safetyYn[i].checked=false;
		itemreg.safetyYn[i].disabled=true;
	}
	$('select[name="safetyDiv"] option').prop("disabled", true);

	// 상품고시에 딸려 있는 각각의 안전인증 법적 기본값을 가져온다.
	// 법적 안전인증대상여부
	var SafetyTargetYN = $('select[name="infoDiv"] option:selected').attr('SafetyTargetYN');
	// 법적 안전인증여부
	var SafetyCertYN = $('select[name="infoDiv"] option:selected').attr('SafetyCertYN');
	// 법적 안전확인여부
	var SafetyConfirmYN = $('select[name="infoDiv"] option:selected').attr('SafetyConfirmYN');
	// 법적 공급자적합성여부
	var SafetySupplyYN = $('select[name="infoDiv"] option:selected').attr('SafetySupplyYN');
	// 법적 안전기준준수여부
	var SafetyComply = $('select[name="infoDiv"] option:selected').attr('SafetyComply');

	// 안전인증대상구분
	var safetyDiv = itemreg.safetyDiv.value
	// 안전인증대상인증번호
	var safetyNum = itemreg.safetyNum.value

	// 법적 안전인증대상여부 비대상
	if ( SafetyTargetYN=='N'){
		// 법적 안전기준준수여부 대상
		if ( SafetyComply=='Y'){
			itemreg.safetyYn[3].disabled=false;
			itemreg.safetyYn[3].checked=true;
		}else{
			itemreg.safetyYn[1].disabled=false;
			itemreg.safetyYn[1].checked=true;
		}

	// 법적 안전인증대상여부 대상 or 비대상
	}else if ( SafetyTargetYN=='Y' || SafetyTargetYN=='S' ){
		itemreg.safetyYn[0].disabled=false;
		itemreg.safetyYn[1].disabled=false;
		itemreg.safetyYn[2].disabled=false;

		// 법적 안전기준준수여부 대상
		if ( SafetyComply=='Y'){
			itemreg.safetyYn[3].disabled=false;
		}

		// 법적 안전인증여부
		if (SafetyCertYN=='Y'){
			$('select[name="safetyDiv"] option[value="10"]').prop("disabled", false)
			$('select[name="safetyDiv"] option[value="40"]').prop("disabled", false)
			$('select[name="safetyDiv"] option[value="70"]').prop("disabled", false)
		}

		// 법적 안전확인여부
		if (SafetyConfirmYN=='Y'){
			$('select[name="safetyDiv"] option[value="20"]').prop("disabled", false)
			$('select[name="safetyDiv"] option[value="50"]').prop("disabled", false)
			$('select[name="safetyDiv"] option[value="80"]').prop("disabled", false)
		}

		// 법적 공급자적합성여부
		if (SafetySupplyYN=='Y'){
			$('select[name="safetyDiv"] option[value="30"]').prop("disabled", false)
			$('select[name="safetyDiv"] option[value="60"]').prop("disabled", false)
			$('select[name="safetyDiv"] option[value="90"]').prop("disabled", false)
		}

		// 이미 이전에 입력한 저장값이 있을경우.
		if (orgsafetyYn=='Y'){
			itemreg.safetyYn[0].checked=true;
		}else if (orgsafetyYn=='N'){
			itemreg.safetyYn[1].checked=true;
		}else if (orgsafetyYn=='I'){
			itemreg.safetyYn[2].checked=true;
		}else if (orgsafetyYn=='S'){
			itemreg.safetyYn[3].checked=true;

		// 없으면 기본값 대상으로 선택
		}else{
			itemreg.safetyYn[0].checked=true;
		}
	}

	// 상품구분
	var itemdiv='';
	for (var i=0; i < itemreg.itemdiv.length; i++){
		if (itemreg.itemdiv[i].checked){
			itemdiv = itemreg.itemdiv[i].value;
		}
	}
	// 티켓상품일경우 대상아님으로 체크
	if (itemdiv=='08'){
		itemreg.safetyYn[1].disabled=false;
		itemreg.safetyYn[1].checked=true;
	}

	chgSafetyYn(document.itemreg);
}

// 안전인증정보 선택
function chgSafetyYn(frm) {
	if(frm.safetyYn[0].checked) {
		frm.safetyDiv.disabled=false;
		frm.safetyNum.disabled=false;
		$("#safetybtn").show();
		$("#safetyYnI").hide();
		$("#safetyDivList").show();
	} else if(frm.safetyYn[2].checked) {
		frm.safetyDiv.disabled=true;
		frm.safetyNum.disabled=true;
		$("#safetybtn").hide();
		$("#safetyYnI").show();
		$("#safetyDivList").hide();
	} else {
		jsAlertCatecodeSafety();
		frm.safetyDiv.disabled=true;
		frm.safetyNum.disabled=true;
		$("#safetybtn").hide();
		$("#safetyYnI").hide();
		$("#safetyDivList").hide();
	}
}

//안전인증 추가 버튼 액션
function jsSafetyAuth(){
	//var cnum = $("#safetyNum").val();
	var cnum = itemreg.safetyNum.value.ltrim().rtrim();
	var sdiv = itemreg.safetyDiv.value.ltrim().rtrim();
	var listbody = "";
	var safetyvalue = "";
	var safetynum = "";

	if(typeof itemreg.catecode == "undefined"){
		alert("카테고리를 선택해 주세요.");
		return;
	}

	if($("#safetyDiv").val() == ""){
		alert("안전인증구분을 선택해 주세요.");
		return;
	}

	var isExist = $("#real_safetydiv").attr("value").indexOf($("#safetyDiv").val()) > -1;
	if(isExist){
		alert("이미 선택된 안전인증구분 입니다.");
		return;
	}
//	var isExistsafetynum = $("#real_safetynum").attr("value").indexOf(cnum) > -1;
//	if(isExistsafetynum){
//		alert("이미 선택된 안전인증번호 입니다.");
//		return;
//	}

	if($("#safetyDiv").val() == "30" || $("#safetyDiv").val() == "60" || $("#safetyDiv").val() == "90"){
		$("#issafetyauth").val("ok");

		safetyvalue = $("#real_safetydiv").val();
		if(safetyvalue == ""){
			$("#real_safetydiv").val($("#safetyDiv").val());
		}else{
			$("#real_safetydiv").val(safetyvalue + "," + $("#safetyDiv").val())
		}

		safetynum = $("#real_safetynum").val();
		if(safetynum == ""){
			$("#real_safetynum").val("x");
		}else{
			$("#real_safetynum").val(safetynum + "," + "x");
		}


		listbody = $("#safetyDivList").html();
		$("#safetyDivList").html(listbody + "<p id='l"+$("#safetyDiv").val()+"'>- " + $("#safetyDiv option:selected").text() + "(인증번호 없음) <input type='button' value='삭제' onClick='jsSafetyDivListDel("+$("#safetyDiv").val()+");' class='button'><p>");
	}else{

		var msgg = jsCallAPIsafety(cnum,"x",sdiv);

		if(msgg == "적합" || msgg == "변경" || msgg == "개선명령" || msgg == "청문실시"){
			$("#issafetyauth").val("ok");

			safetyvalue = $("#real_safetydiv").val();
			if(safetyvalue == ""){
				$("#real_safetydiv").val($("#safetyDiv").val());
			}else{
				$("#real_safetydiv").val(safetyvalue + "," + $("#safetyDiv").val())
			}

			safetynum = $("#real_safetynum").val();
			if(safetynum == ""){
				$("#real_safetynum").val(cnum);
			}else{
				$("#real_safetynum").val(safetynum + "," + cnum);
			}


			listbody = $("#safetyDivList").html();
			$("#safetyDivList").html(listbody + "<p id='l"+$("#safetyDiv").val()+"'>- " + $("#safetyDiv option:selected").text() + "("+cnum+") <input type='button' value='삭제' onClick='jsSafetyDivListDel("+$("#safetyDiv").val()+");' class='button'><p>");
		}else{
			alert("인증번호의 현재 상태 : " + msgg);
			return;
		}
	}
	jsSafetyDefault();
}

//안전인증 선택 및 입력된 값 디폴트셋팅
function jsSafetyDefault(){
	$("#safetyDiv").val("").attr("selected","selected");
	//$("#safetyNum").val("");
	itemreg.safetyNum.value = "";
}

//안전인증 필수 품목 확인 새창
function jsSafetyPopup(){
	window.open("http://www.safetykorea.kr/policy/targetsSafetyCert","jsSafetyPopup","width=1200, height=1000, scrollbars=yes, resizable=yes");
}

///////////////////////////////////////////////////////////////////

// 주민등록번호 확인
function jsChkSocialNum1(varSno){
	var sno = varSno;
	var IDAdd = "234567892345";
	var iDot=0;
	
	//숫자확인 
	if(!IsDouble(sno)){
	  return false;	
  }	
	//숫자가 13자리 인지 확인 
	if(sno.length != 13){
	  return false;	
   }	
	if (sno.substring(2,3) > 1) return false;
	if (sno.substring(4,5) > 3) return false;
	if (sno.substring(0,2) == '00' && (sno.substring(6,7) != 0 || sno.substring(6,7) != 9 || sno.substring(6,7) != 3 || sno.substring(6,7) !=4)) return false;
	if (sno.substring(0,2) != '00' && (sno.substring(6,7) > 4 || sno.substring(6,7) == 0)) return false;	
  
	for(var i=0; i < 13; i ++)
	  iDot = iDot + sno.substr(i, 1) * IDAdd.substr(i,1);
	
	iDot = 11 - (iDot % 11);
	
	if(iDot == 10){
	  iDot = 0;
	} else if (iDot == 11){
	  iDot = 1;
	}
		  
	if(sno.substr(12,1) == iDot){
	  return true;
	} else {
	  return false;
	} 
}

// AES256 암호화 추가	2022.10.31 한용민
function AES_Encode(key,plain_text){
	GibberishAES.size(256);	
	return GibberishAES.aesEncrypt(plain_text, key);
}
// AES256 복호화 추가	2022.10.31 한용민
function AES_Decode(key,base64_text){
	GibberishAES.size(256);	
	return GibberishAES.aesDecrypt(base64_text, key);
}

// 클립보드로 테스트 카피
function copyStringToClipboard (string) {
	function handler (event){
		event.clipboardData.setData('text/plain', string);
		event.preventDefault();
		document.removeEventListener('copy', handler, true);
	}

	document.addEventListener('copy', handler, true);
	document.execCommand('copy');
}