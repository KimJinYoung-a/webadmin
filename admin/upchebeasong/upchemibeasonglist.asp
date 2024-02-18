<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 미출고리스트_업배
' Hieditor : 이상구 생성
'			 2023.02.07 한용민 수정(검색조건 수기처리 되어 있는 부분 공통함수로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%

Sub DrawItemReipgoDateBox(byval yyyy1,mm1,dd1, yyyy2,mm2,dd2)
	dim buf,i

	buf = "<select class='select' name='item_yyyy1'>"
    buf = buf + "<option value='" + CStr(yyyy1) +"' selected>" + CStr(yyyy1) + "</option>"
    for i=2002 to Year(now) + 2
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='item_mm1' >"
    buf = buf + "<option value='" + CStr(mm1) + "' selected>" + CStr(mm1) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='item_dd1'>"
    buf = buf + "<option value='" + CStr(dd1) +"' selected>" + CStr(dd1) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    buf = buf + " ~ "

    response.write buf

	buf = "<select class='select' name='item_yyyy2'>"
    buf = buf + "<option value='" + CStr(yyyy2) +"' selected>" + CStr(yyyy2) + "</option>"
    for i=2002 to Year(now) + 2
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='item_mm2' >"
    buf = buf + "<option value='" + CStr(mm2) + "' selected>" + CStr(mm2) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='item_dd2'>"
    buf = buf + "<option value='" + CStr(dd2) +"' selected>" + CStr(dd2) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    response.write buf

end Sub

Sub DrawChulgoDateBox(byval yyyy1,mm1,dd1, yyyy2,mm2,dd2)
	dim buf,i

	buf = "<select class='select' name='chulgo_yyyy1' onChange='SetRegBeasongDayType(1);'>"
    buf = buf + "<option value='" + CStr(yyyy1) +"' selected>" + CStr(yyyy1) + "</option>"
    for i=2002 to Year(now) + 2
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='chulgo_mm1' onChange='SetRegBeasongDayType(1);'>"
    buf = buf + "<option value='" + CStr(mm1) + "' selected>" + CStr(mm1) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='chulgo_dd1' onChange='SetRegBeasongDayType(1);'>"
    buf = buf + "<option value='" + CStr(dd1) +"' selected>" + CStr(dd1) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    buf = buf + " ~ "

    response.write buf

	buf = "<select class='select' name='chulgo_yyyy2' onChange='SetRegBeasongDayType(1);'>"
    buf = buf + "<option value='" + CStr(yyyy2) +"' selected>" + CStr(yyyy2) + "</option>"
    for i=2002 to Year(now) + 2
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='chulgo_mm2' onChange='SetRegBeasongDayType(1);'>"
    buf = buf + "<option value='" + CStr(mm2) + "' selected>" + CStr(mm2) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='chulgo_dd2' onChange='SetRegBeasongDayType(1);'>"
    buf = buf + "<option value='" + CStr(dd2) +"' selected>" + CStr(dd2) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    response.write buf

end Sub

dim chulgoone_yyyy1, chulgoone_mm1 , chulgoone_dd1
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim item_yyyy1,item_yyyy2,item_mm1,item_mm2,item_dd1,item_dd2
dim chulgo_yyyy1,chulgo_yyyy2,chulgo_mm1,chulgo_mm2,chulgo_dd1,chulgo_dd2
dim nowdate,searchnextdate
dim makerid,dateback, cdl
dim cknodate,page, detailstate
dim MisendReason, MisendState, dplusOver, dplusLower, vSiteName, sortby
dim itemid
dim Dtype
dim excludeall, exinmaychulgoday, exinneedchulgoday, exstockout
dim cdm, cds, dispCate

nowdate = Left(CStr(now()),10)
makerid = requestCheckVar(request("makerid"),32)

yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1     = requestCheckVar(request("mm1"),2)
dd1     = requestCheckVar(request("dd1"),2)
yyyy2   = requestCheckVar(request("yyyy2"),4)
mm2     = requestCheckVar(request("mm2"),2)
dd2     = requestCheckVar(request("dd2"),2)

detailstate   = requestCheckVar(request("detailstate"),9)
cdl         = requestCheckVar(request("cdl"),3)
cknodate    = requestCheckVar(request("cknodate"),16)
page        = requestCheckVar(request("page"),9)
MisendReason = requestCheckVar(request("MisendReason"),2)
MisendState  = requestCheckVar(request("MisendState"),2)
dplusOver   = requestCheckVar(request("dplusOver"),10)
dplusLower   = requestCheckVar(request("dplusLower"),10)
itemid      = requestCheckVar(request("itemid"),10)
Dtype       = requestCheckVar(request("Dtype"),10)
vSiteName	= requestCheckVar(request("sitename"),10)
sortby	= requestCheckVar(request("sortby"),32)

excludeall	= requestCheckVar(request("excludeall"),32)
exinmaychulgoday	= requestCheckVar(request("exinmaychulgoday"),32)
exinneedchulgoday	= requestCheckVar(request("exinneedchulgoday"),32)
exstockout	= requestCheckVar(request("exstockout"),32)

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)

if (excludeall = "Y") then
	exinmaychulgoday = "Y"
	exinneedchulgoday = "Y"
end if

Dim popupFlag	: popupFlag = req("popupFlag","")	' 팝업
Dim currState	: currState = req("currState","")	' 상태

if (page="") then page=1

if (Dtype="") then Dtype = "date"
if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

    dateback = DateSerial(yyyy1,mm2-2, dd2+1)

    yyyy1 = Left(dateback,4)
    mm1   = Mid(dateback,6,2)
    dd1   = Mid(dateback,9,2)

	If popupFlag = "" Then
	    MisendState = "0"
	    Dtype = "topN"
	End If
end if

if yyyy2<>"" and not(isnull(yyyy2)) and mm2<>"" and not(isnull(mm2)) and dd2<>"" and not(isnull(dd2)) then
searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)
end if


'// ===========================================================================
dim ojumun

set ojumun = new CBaljuMaster

if cknodate="" then
	ojumun.FRectRegStart = LEft(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
	ojumun.FRectRegEnd = searchnextdate
end if


ojumun.FRectDesignerID = makerid
ojumun.FPageSize = 50
ojumun.FCurrPage = page
ojumun.FRectCDL  = cdl
If currState = "" Then
	ojumun.FRectDetailState = "NOT7" ''"UP2NOT7"
Else
	ojumun.FRectDetailState = currState
End If
ojumun.FRectMisendReason = MisendReason
ojumun.FRectMisendState  = MisendState
ojumun.FRectdplusOver = dplusOver
ojumun.FRectdplusLower = dplusLower
ojumun.FRectItemID = itemid
ojumun.FRectSiteName = vSiteName
ojumun.FRectSortBy = sortby

ojumun.FRectExInMayChulgoDay = exinmaychulgoday
ojumun.FRectExInNeedChulgoDay = exinneedchulgoday
ojumun.FRectExStockOut = exstockout

''ojumun.FRectDispCDL     = cdl
''ojumun.FRectDispCDM     = cdm
''ojumun.FRectDispCDS     = cds
ojumun.FRectDispCate	= dispCate

if (Dtype = "topN") then
    ojumun.FPageSize = 100					'// 300개 -> 100개
    ojumun.getUpcheMichulgoList(true)
else
    ojumun.getUpcheMichulgoList(false)
end if



'// ===========================================================================
dim OCSBrandMemo

set OCSBrandMemo = new CCSBrandMemo

OCSBrandMemo.FRectMakerid = makerid

if (makerid <> "") then

	OCSBrandMemo.GetBrandMemo

end if

if (OCSBrandMemo.Fbeasongneedday = "") or (IsNull(OCSBrandMemo.Fbeasongneedday)) then
	OCSBrandMemo.Fbeasongneedday = 0
	OCSBrandMemo.Fbeasong_comment = "해당 브랜드 배송담당자 연락처 등"
end if

'// ===========================================================================
dim OCSItemMemo

set OCSItemMemo = new CCSItemMemo

OCSItemMemo.FRectItemId = itemid

if (itemid <> "") then

	OCSItemMemo.GetItemidMemo

end if

if (OCSItemMemo.Fbeasongneedday = "") or (IsNull(OCSItemMemo.Fbeasongneedday)) then
	OCSItemMemo.Fbeasongneedday = 0
	OCSItemMemo.Fbeasong_comment = "일시적인 출고지연이면 언제부터 정상출고 가능한지, 통상적인 출고소요일 등"

	OCSItemMemo.Fmaketoorderyn = "N"
	OCSItemMemo.Fstockshortyn = "N"
	OCSItemMemo.Freipgostartday = Left(now, 10)
	OCSItemMemo.Freipgoendday = Left(now, 10)
end if

item_yyyy1 = Left(OCSItemMemo.Freipgostartday, 4)
item_yyyy2 = Left(OCSItemMemo.Freipgoendday, 4)
item_mm1 = Right(Left(OCSItemMemo.Freipgostartday, 7), 2)
item_mm2 = Right(Left(OCSItemMemo.Freipgoendday, 7), 2)
item_dd1 = Right(OCSItemMemo.Freipgostartday, 2)
item_dd2 = Right(OCSItemMemo.Freipgoendday, 2)

dim reipgotype
if (item_yyyy1 = item_yyyy2) and (item_mm1 = item_mm2) and (item_dd1 = item_dd2) then
	reipgotype = "1"
else
	reipgotype = "N"
end if

if (OCSItemMemo.Freipgoendday <= Left(now, 10)) then
	OCSItemMemo.Fstockshortyn = "N"
end if


if (chulgo_yyyy1 = "") then
	chulgo_yyyy1 = Left(nowdate, 4)
	chulgo_yyyy2 = Left(nowdate, 4)
	chulgo_mm1 = Right(Left(nowdate, 7), 2)
	chulgo_mm2 = Right(Left(nowdate, 7), 2)
	chulgo_dd1 = Right(nowdate, 2)
	chulgo_dd2 = Right(nowdate, 2)
end if

if (chulgoone_yyyy1 = "") then
	chulgoone_yyyy1 = Left(nowdate, 4)
	chulgoone_mm1 = Right(Left(nowdate, 7), 2)
	chulgoone_dd1 = Right(nowdate, 2)
end if


dim ix,iy
dim IsDisableThis
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function chkSubmit(){
    var frm = document.frm;

    if ((frm.itemid.value.length>0)&&(!IsDigit(frm.itemid.value))){
        alert('상품번호는 숫자로 입력하세요.');
        frm.itemid.focus();
        return;
    }

    frm.yyyy1.disabled=false;
    frm.yyyy2.disabled=false;
    frm.mm1.disabled=false;
    frm.mm2.disabled=false;
    frm.dd1.disabled=false;
    frm.dd2.disabled=false;

    frm.submit();
}

function changecontent(){
    //nothing
}

function misendmaster(v){
	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + v,"misendmaster","width=1200 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = '_blank';
    frm.action="/admin/ordermaster/viewordermaster.asp"
	frm.submit();

}

function ViewItem(itemid){
window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"sample");
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function chkComp(comp){
    comp.form.yyyy1.disabled=(comp.value=="topN");
    comp.form.yyyy2.disabled=(comp.value=="topN");
    comp.form.mm1.disabled=(comp.value=="topN");
    comp.form.mm2.disabled=(comp.value=="topN");
    comp.form.dd1.disabled=(comp.value=="topN");
    comp.form.dd2.disabled=(comp.value=="topN");
}

function chkComp2(frm){
    var chk = frm.excludeall.checked;

    frm.exinmaychulgoday.disabled = chk;
    frm.exinneedchulgoday.disabled = chk;
}

function searchByMakerId(frm, makerid) {
	frm.makerid.value = makerid;
	frm.itemid.value = "";

	chkSubmit();
}

function searchByItemId(frm, itemid) {
	frm.itemid.value = itemid;

	chkSubmit();
}

function jsShowHideObject(id) {
	if (document.getElementById) {
		obj = document.getElementById(id);

		if (obj.style.display == "none") {
			obj.style.display = "";
		} else {
			obj.style.display = "none";
		}
	}
}

function jsShowHideItemInfo(frm) {
	var obj;

	obj = document.getElementById("maketoorder");
	if (getCheckedValue(frm.maketoorderyn) == "Y") {
		obj.style.display = "";
	} else {
		obj.style.display = "none";
	}

	obj = document.getElementById("stockshort");
	if (getCheckedValue(frm.stockshortyn) == "Y") {
		obj.style.display = "";
	} else {
		obj.style.display = "none";
	}

	if (getCheckedValue(frm.reipgotype) == "1") {
		frm.item_yyyy2.disabled = true;
		frm.item_mm2.disabled = true;
		frm.item_dd2.disabled = true;
	} else {
		frm.item_yyyy2.disabled = false;
		frm.item_mm2.disabled = false;
		frm.item_dd2.disabled = false;
	}
}

function jsUpcheBrandBeasongMemo(makerid, tableobj)
{

	if (makerid == "") {
		alert("먼저 브랜드로 검색하세요.");
		return;
	}

	jsShowHideObject(tableobj);
}

function jsUpcheItemBeasongMemo(itemid, tableobj)
{

	if (itemid == "") {
		alert("먼저 상품코드로 검색하세요.");
		return;
	}

	jsShowHideObject(tableobj);
}

function jsMultiMichulgoReason(makerid, itemid, tableobj)
{

	if ((makerid == "") && (itemid == "")) {
		alert("브랜드 또는 상품코드로 검색하세요.");
		return;
	}

	jsShowHideObject(tableobj);
}

function submitSaveBrandMemo(frm)
{

	if (frm.makerid.value == "") {
		alert("먼저 브랜드로 검색하세요.");
		return;
	}

	if (frm.beasongneedday.value == "") {
		alert("평균출고소요일을 입력하세요.");
		return;
	}

	if (frm.beasongneedday.value*0 != 0) {
		alert("평균출고소요일은 숫자만 입력가능합니다.");
		return;
	}

	if (confirm("저장하시겠습니까?") == true) {
		frm.submit();
	}
}

function submitSaveItemMemo(frm)
{

	if (frm.itemid.value == "") {
		alert("먼저 상품코드로 검색하세요.");
		return;
	}

	if (frm.beasongneedday.value == "") {
		alert("평균출고소요일을 입력하세요.");
		return;
	}

	if (frm.beasongneedday.value*0 != 0) {
		alert("평균출고소요일은 숫자만 입력가능합니다.");
		return;
	}

	if (confirm("저장하시겠습니까?") == true) {
		frm.submit();
	}
}

var IsButtonPressed = false;
function multiMisendInput(frm) {

	if (IsButtonPressed == true) {
		alert("먼저 검색하기 버튼을 누르세요. ");
		return;
	}

	if (CheckSelected() != true) {
		alert("선택된 주문이 없습니다.");
		return;
	}

	if (frm.regMisendReason.value == "") {
		alert("미출고 사유를 선택하세요.");
		frm.regMisendReason.focus();
		return;
	}

	if ((frm.ckSendSMS.checked != true) && (frm.ckSendEmail.checked != true)) {
		alert("SMS 와 메일발송 둘중 하나는 체크해야 합니다.");
		return;
	}

	if (frm.regbeasongdaytype[2].checked == true) {
		if (frm.regbeasongneedday.value == "") {
			alert("배송 소요일을 입력하세요.");
			frm.regbeasongneedday.focus();
			return;
		}

		if (frm.regbeasongneedday.value*0 != 0) {
			alert("배송 소요일은 숫자만 입력가능합니다.");
			frm.regbeasongneedday.focus();
			return;
		}
	} else if (frm.regbeasongdaytype[0].checked == true) {
		if (frm.chulgooneday.value.length != 10) {
			alert("출고예정일을 입력하세요.");
			frm.chulgooneday.focus();
			return;
		}

		frm.chulgoone_yyyy1.value = frm.chulgooneday.value.substr(0, 4);
		frm.chulgoone_mm1.value = frm.chulgooneday.value.substr(5, 2);
		frm.chulgoone_dd1.value = frm.chulgooneday.value.substr(8, 2);

		if ((frm.chulgoone_yyyy1.value*0 != 0) || (frm.chulgoone_mm1.value*0 != 0) || (frm.chulgoone_dd1.value*0 != 0)) {
			alert("잘못된 출고예정일입니다.");
			frm.chulgooneday.focus();
			return;
		}

		var nowDate = new Date();
		var date1 = new Date();

		date1.setFullYear((frm.chulgoone_yyyy1.value * 1), (frm.chulgoone_mm1.value * 1 - 1), (frm.chulgoone_dd1.value * 1));

		if (nowDate > date1) {
			alert("잘못된 출고예정일입니다.");
			frm.chulgooneday.focus();
			return;
		}
	} else if (frm.regbeasongdaytype[1].checked == true) {
		var nowDate = new Date();
		var date1 = new Date();
		var date2 = new Date();

		date1.setFullYear((frm.chulgo_yyyy1.value * 1), (frm.chulgo_mm1.value * 1 - 1), (frm.chulgo_dd1.value * 1));
		date2.setFullYear((frm.chulgo_yyyy2.value * 1), (frm.chulgo_mm2.value * 1 - 1), (frm.chulgo_dd2.value * 1));

		if (nowDate > date1) {
			alert("잘못된 출고예정일입니다.");
			frm.chulgo_yyyy1.focus();
			return;
		}

		if (nowDate > date2) {
			alert("잘못된 출고예정일입니다.");
			frm.chulgo_yyyy2.focus();
			return;
		}

		if (date1 > date2) {
			alert("잘못된 출고예정일입니다.");
			frm.chulgo_yyyy1.focus();
			return;
		}
	}

	if (frm.sendsmsmsg.value == "") {
		alert("SMS발송문구를 입력하세요.");
		frm.sendsmsmsg.focus();
		return;
	}

	if (frm.sendmailmsg.value == "") {
		alert("MAIL발송문구를 입력하세요.");
		frm.sendmailmsg.focus();
		return;
	}

	if (confirm("일괄저장하시겠습니까?") == true) {

		IsButtonPressed = true;

		for (var i=0;i<document.forms.length;i++){
			f = document.forms[i];
			if (f.name.substr(0,9)=="frmBuyPrc") {
				if (f.detailidx.checked) {
					frm.arrdetailidx.value = frm.arrdetailidx.value + "," + f.detailidx.value;
					frm.arrbaljudate.value = frm.arrbaljudate.value + "," + f.baljudate.value;
				}
			}
		}

		frm.submit();
	}
}

function getCheckedValue(radioObj) {
	if(!radioObj)
		return "";
	var radioLength = radioObj.length;
	if(radioLength == undefined)
		if(radioObj.checked)
			return radioObj.value;
		else
			return "";
	for(var i = 0; i < radioLength; i++) {
		if(radioObj[i].checked) {
			return radioObj[i].value;
		}
	}
	return "";
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.detailidx.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function jsSetMisendReason(frm) {
	jsSetSMSMailText(frm);
}

function jsSetSMSMailText(frm) {
	jsSetSMSText(frm);
	jsSetMailText(frm);
}

function jsSetSMSText(frm) {
	var smsText;

	smsText = "";

	switch (frm.regMisendReason.value) {
		case "03" : {
			// 출고지연
			smsText = "<%= GetMichulgoSMSString("03") %>";

			break;
		}

		case "02" : {
			// 주문제작
			smsText = "<%= GetMichulgoSMSString("02") %>";

			break;
		}

		case "08" : {
			// 수입
			smsText = "<%= GetMichulgoSMSString("08") %>";

			break;
		}

		case "09" : {
			// 가구배송
			smsText = "<%= GetMichulgoSMSString("09") %>";

			break;
		}

		case "04" : {
			// 예약배송
			smsText = "<%= GetMichulgoSMSString("04") %>";

			break;
		}

		case "10" : {
			// 업체휴가
			smsText = "<%= GetMichulgoSMSString("10") %>";

			break;
		}

		case "07" : {
			// 고객지정배송
			smsText = "<%= GetMichulgoSMSString("07") %>";

			break;
		}

		default : {
			//
		}
	}

	frm.sendsmsmsg.value = smsText;
}

function jsSetMailText(frm) {
	var mailText;

	mailText = "";

	switch (frm.regMisendReason.value) {
		case "03" : {
			// 출고지연
			mailText = "<%= GetMichulgoMailString("03") %>";

			break;
		}

		case "02" : {
			// 주문제작
			mailText = "<%= GetMichulgoMailString("02") %>";

			break;
		}

		case "08" : {
			// 수입
			mailText = "<%= GetMichulgoMailString("08") %>";

			break;
		}

		case "09" : {
			// 가구배송
			mailText = "<%= GetMichulgoMailString("09") %>";

			break;
		}

		case "04" : {
			// 예약배송
			mailText = "<%= GetMichulgoMailString("04") %>";

			break;
		}

		case "10" : {
			// 업체휴가
			mailText = "<%= GetMichulgoMailString("10") %>";

			break;
		}

		case "07" : {
			// 고객지정배송
			mailText = "<%= GetMichulgoMailString("07") %>";

			break;
		}

		default : {
			//
		}
	}

	frm.sendmailmsg.value = mailText;
}

function SetRegBeasongDayType(idx) {
    var frm = frmMisendInput;

    frm.regbeasongdaytype[idx].checked = true;

     jsSetSMSMailText(frm);
}

function CheckAll(chk) {
	for (var i = 0; ; i++) {
		var v = document.getElementById("detailidx_" + i);
		if (v == undefined) {
			return;
		}

		if (v.disabled != true) {
			v.checked = chk.checked;
		}
	}
}

function getOnLoad(){
    if (document.frm.Dtype[0].checked){
        chkComp(document.frm.Dtype[0]);
    }else{
        chkComp(document.frm.Dtype[1]);
    }

    // chkComp2(frm);


	<% if (OCSItemMemo.Fmaketoorderyn = "Y") or (OCSItemMemo.Fstockshortyn = "Y") then %>
		jsUpcheItemBeasongMemo('<%= itemid %>', 'itemmemo');
	<% end if %>

    jsShowHideItemInfo(frmItemMemo);
}

window.onload=getOnLoad;

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			&nbsp;
			Site : <% call drawSelectBoxXSiteOrderInputPartnerCS("sitename", vSiteName) %>
			&nbsp;
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">
			&nbsp;

			<input type="radio" name="Dtype" value="topN" <%= cHKIIF(Dtype="topN","checked","") %> onClick="chkComp(this);" >TOP <%= CHKIIF(Dtype = "topN",ojumun.FPageSize,100) %>개(최근2달)
			&nbsp;<input type="radio" name="Dtype" value="date" <%= cHKIIF(Dtype="date","checked","") %>  onClick="chkComp(this);" >검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="chkSubmit();">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td>
	        카테고리 : <% DrawSelectBoxCategoryLarge "cdl",cdl %>
			&nbsp;
			<!-- 2009 추가 -->
			소요일수 :
			<select class="select" name="dplusOver">
				<option value="" >전체</option>
				<option value="2" <%= CHKIIF(dplusOver="2","selected","") %> >D+2이상</option>
				<option value="3" <%= CHKIIF(dplusOver="3","selected","") %> >D+3이상</option>
				<option value="4" <%= CHKIIF(dplusOver="4","selected","") %> >D+4이상</option>
			</select>
			~
			<select class="select" name="dplusLower">
				<option value="" >전체</option>
				<option value="7" <%= CHKIIF(dplusLower="7","selected","") %> >D+7이하</option>
				<option value="14" <%= CHKIIF(dplusLower="14","selected","") %> >D+14이하</option>
			</select>
			&nbsp;
			미출고사유 :
			<select class="select" name="MisendReason">
				<option value="">전체</option>
				<option value="">--------</option>
				<option value="00" <%= CHKIIF(MisendReason="00","selected","") %> >입력이전</option>
				<option value="">--------</option>
				<option value="03" <%= CHKIIF(MisendReason="03","selected","") %> >출고지연</option>
				<option value="02" <%= CHKIIF(MisendReason="02","selected","") %> >주문제작</option>
				<option value="08" <%= CHKIIF(MisendReason="08","selected","") %> >수입</option>
				<option value="09" <%= CHKIIF(MisendReason="09","selected","") %> >가구배송</option>
				<option value="04" <%= CHKIIF(MisendReason="04","selected","") %> >예약배송</option>
				<option value="10" <%= CHKIIF(MisendReason="10","selected","") %> >업체휴가</option>
				<option value="07" <%= CHKIIF(MisendReason="07","selected","") %> >고객지정배송</option>
				<option value="">--------</option>
				<option value="05" <%= CHKIIF(MisendReason="05","selected","") %> >품절출고불가</option>
				<option value="">--------</option>
			</select>
			&nbsp;
			처리구분 :
			<select class="select" name="MisendState">
				<option value="">전체</option>
				<!--
				<option value="N" <%= CHKIIF(MisendState="N","selected","") %> >사유미등록전체</option>
				-->
				<option value="0" <%= CHKIIF(MisendState="0","selected","") %> >CS(CALL)미처리</option>
				<option value="4" <%= CHKIIF(MisendState="4","selected","") %> >고객안내</option>
				<option value="6" <%= CHKIIF(MisendState="6","selected","") %> >CS처리완료</option>
			</select>
			&nbsp;
			상태 :
			<select class="select" name="currState">
				<option value="">전체</option>
				<option value="0" <%= CHKIIF(currState="0","selected","") %> >결제완료</option>
				<option value="2" <%= CHKIIF(currState="2","selected","") %> >주문통보</option>
				<option value="3" <%= CHKIIF(currState="3","selected","") %> >주문확인</option>
			</select>
			&nbsp;
			정렬순서 :
			<select class="select" name="sortby">
				<option value="">소요일수</option>
				<option value="makerid" <%= CHKIIF(sortby="makerid","selected","") %> >브랜드</option>
				<option value="orderserial" <%= CHKIIF(sortby="orderserial","selected","") %> >주문번호</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td>
			전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	    	<!--
			<input type="checkbox" class="checkbox" name="excludeall" value="Y" <%= CHKIIF(excludeall="Y","checked","") %> onClick="chkComp2(frm);"> 전부 제외
			-->
			<input type="checkbox" class="checkbox" name="exinmaychulgoday" value="Y" <%= CHKIIF(exinmaychulgoday="Y","checked","") %>> 출고예정일 이전 주문 제외
			<input type="checkbox" class="checkbox" name="exstockout" value="Y" <%= CHKIIF(exstockout="Y","checked","") %>> 품절출고불가 제외
			<!--
			<input type="checkbox" class="checkbox" name="exinneedchulgoday" value="Y" <%= CHKIIF(exinneedchulgoday="Y","checked","") %>> 출고소요일 이전 주문 제외
			-->
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

* D+2, D+4 일은 <font color=red>근무일수 기준</font>입니다.

<p>

<% if (MisendState = "6") and False then %>
* CS처리완료 이면서 주문내역이 표시되면 특정상품 2개 이상 주문 후 <font color=red>일부만 취소</font>한 경우입니다.
<% end if %>

<input type="button" class="button" value="브랜드 배송관련 메모" onClick="jsUpcheBrandBeasongMemo('<%= makerid %>', 'brandmemo');">
<input type="button" class="button" value="상품 배송관련 메모" onClick="jsUpcheItemBeasongMemo('<%= itemid %>', 'itemmemo');">
<input type="button" class="button" value="미출고 사유 일괄입력" onClick="jsMultiMichulgoReason('<%= makerid %>', '<%= itemid %>', 'regallmisendreason');">

<p>

<div id="brandmemo" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmBrandMemo" method="post" action="upchemibeasonglist_process.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="mode" value="modifybrandmemo">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="makerid" value="<%= makerid %>">
	<input type="hidden" name="sitename" value="<%= vSiteName %>">
	<input type="hidden" name="itemid" value="<%= itemid %>">
	<input type="hidden" name="Dtype" value="<%= Dtype %>">
	<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
	<input type="hidden" name="mm1" value="<%= mm1 %>">
	<input type="hidden" name="dd1" value="<%= dd1 %>">
	<input type="hidden" name="yyyy2" value="<%= yyyy2 %>">
	<input type="hidden" name="mm2" value="<%= mm2 %>">
	<input type="hidden" name="dd2" value="<%= dd2 %>">
	<input type="hidden" name="cdl" value="<%= cdl %>">
	<input type="hidden" name="dplusOver" value="<%= dplusOver %>">
	<input type="hidden" name="dplusLower" value="<%= dplusLower %>">
	<input type="hidden" name="exinmaychulgoday" value="<%= exinmaychulgoday %>">
	<input type="hidden" name="exinneedchulgoday" value="<%= exinneedchulgoday %>">
	<input type="hidden" name="sortby" value="<%= sortby %>">
	<input type="hidden" name="MisendReason" value="<%= MisendReason %>">
	<input type="hidden" name="MisendState" value="<%= MisendState %>">
	<input type="hidden" name="currState" value="<%= currState %>">
	<input type="hidden" name="beasongneedday" value="0">				<!-- 브랜드 전체 출고소요일은 사용안함 -->
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30"><b>브랜드ID</b></td>
		<td width="20%" bgcolor="FFFFFF"><%= makerid %></td>
		<td width="10%"></td>
		<td width="25%" bgcolor="FFFFFF"></td>
		<td width="10%">최종수정일</td>
		<td bgcolor="FFFFFF"><%= OCSBrandMemo.Fbeasong_modifyday %></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30"></td>
		<td width="20%" bgcolor="FFFFFF"></td>
		<td width="10%"></td>
		<td width="25%" bgcolor="FFFFFF"></td>
		<td width="10%">작성자</td>
		<td bgcolor="FFFFFF"><%= OCSBrandMemo.Fbeasong_reguserid %></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">브랜드 배송관련 메모</td>
		<td colspan="5" bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="beasong_comment" cols="100" rows="7"><%= OCSBrandMemo.Fbeasong_comment %></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="FFFFFF" colspan = "6" height="35">
			<input type="button" class="button_s" value=" 저장하기 " onClick="submitSaveBrandMemo(frmBrandMemo)">
		</td>
	</tr>
	</form>
</table>

<p>
</div>

<div id="itemmemo" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmItemMemo" method="post" action="upchemibeasonglist_process.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="mode" value="modifyitemmemo">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="makerid" value="<%= makerid %>">
	<input type="hidden" name="sitename" value="<%= vSiteName %>">
	<input type="hidden" name="itemid" value="<%= itemid %>">
	<input type="hidden" name="Dtype" value="<%= Dtype %>">
	<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
	<input type="hidden" name="mm1" value="<%= mm1 %>">
	<input type="hidden" name="dd1" value="<%= dd1 %>">
	<input type="hidden" name="yyyy2" value="<%= yyyy2 %>">
	<input type="hidden" name="mm2" value="<%= mm2 %>">
	<input type="hidden" name="dd2" value="<%= dd2 %>">
	<input type="hidden" name="cdl" value="<%= cdl %>">
	<input type="hidden" name="dplusOver" value="<%= dplusOver %>">
	<input type="hidden" name="dplusLower" value="<%= dplusLower %>">
	<input type="hidden" name="exinmaychulgoday" value="<%= exinmaychulgoday %>">
	<input type="hidden" name="exinneedchulgoday" value="<%= exinneedchulgoday %>">
	<input type="hidden" name="sortby" value="<%= sortby %>">
	<input type="hidden" name="MisendReason" value="<%= MisendReason %>">
	<input type="hidden" name="MisendState" value="<%= MisendState %>">
	<input type="hidden" name="currState" value="<%= currState %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30"><b>상품코드</b></td>
		<td width="20%" bgcolor="FFFFFF" align="left"><%= itemid %></td>
		<td width="10%">상품구분</td>
		<td width="25%" bgcolor="FFFFFF" align="left">
			<input type="radio" name="maketoorderyn" value="N" onClick="jsShowHideItemInfo(frmItemMemo)" <%= CHKIIF(OCSItemMemo.Fmaketoorderyn = "N","checked","") %> > 일반
			<input type="radio" name="maketoorderyn" value="Y" onClick="jsShowHideItemInfo(frmItemMemo)" <%= CHKIIF(OCSItemMemo.Fmaketoorderyn = "Y","checked","") %> > 주문제작(수입)
		</td>
		<td width="10%">재고구분</td>
		<td bgcolor="FFFFFF" align="left">
			<input type="radio" name="stockshortyn" value="N" onClick="jsShowHideItemInfo(frmItemMemo)" <%= CHKIIF(OCSItemMemo.Fstockshortyn = "N","checked","") %> > 정상
			<input type="radio" name="stockshortyn" value="Y" onClick="jsShowHideItemInfo(frmItemMemo)" <%= CHKIIF(OCSItemMemo.Fstockshortyn = "Y","checked","") %> > 재고부족
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" id="stockshort">
		<td width="15%" height="30">재입고구분</td>
		<td width="20%" bgcolor="FFFFFF"  align="left">
			<input type="radio" name="reipgotype" value="1" onClick="jsShowHideItemInfo(frmItemMemo)" <%= CHKIIF(reipgotype = "1","checked","") %> > 1회입고
			<input type="radio" name="reipgotype" value="N" onClick="jsShowHideItemInfo(frmItemMemo)" <%= CHKIIF(reipgotype = "N","checked","") %> > 분할입고
		</td>
		<td width="10%">재입고예정일</td>
		<td width="25%" bgcolor="FFFFFF" align="left">
			<% DrawItemReipgoDateBox item_yyyy1, item_mm1 , item_dd1, item_yyyy2, item_mm2, item_dd2 %>
		</td>
		<td width="10%"></td>
		<td bgcolor="FFFFFF" align="left">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" id="maketoorder">
		<td width="15%" height="30">제작(수입) 필요일수</td>
		<td width="20%" bgcolor="FFFFFF"  align="left">
			<input type="text" class="text" name="beasongneedday" value="<%= OCSItemMemo.Fbeasongneedday %>" size="1" maxlength="3"> 일
		</td>
		<td width="10%"></td>
		<td width="25%" bgcolor="FFFFFF" align="left">
		</td>
		<td width="10%"></td>
		<td bgcolor="FFFFFF" align="left">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30">최종수정일</td>
		<td width="20%" bgcolor="FFFFFF"  align="left">
			<%= OCSItemMemo.Fbeasong_modifyday %>
		</td>
		<td width="10%">작성자</td>
		<td width="25%" bgcolor="FFFFFF" align="left">
			<%= OCSItemMemo.Fbeasong_reguserid %>
		</td>
		<td width="10%"></td>
		<td bgcolor="FFFFFF" align="left">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">상품 배송관련 메모</td>
		<td colspan="5" bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="beasong_comment" cols="100" rows="7"><%= OCSItemMemo.Fbeasong_comment %></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="FFFFFF" colspan = "6" height="35">
			<input type="button" class="button_s" value=" 저장하기 " onClick="submitSaveItemMemo(frmItemMemo)">
		</td>
	</tr>
	</form>
</table>

<p>
</div>

<div id="regallmisendreason" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmMisendInput" method="post" action="upchemibeasonglist_process.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="mode" value="regallmisendreason">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="makerid" value="<%= makerid %>">
	<input type="hidden" name="sitename" value="<%= vSiteName %>">
	<input type="hidden" name="itemid" value="<%= itemid %>">
	<input type="hidden" name="Dtype" value="<%= Dtype %>">
	<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
	<input type="hidden" name="mm1" value="<%= mm1 %>">
	<input type="hidden" name="dd1" value="<%= dd1 %>">
	<input type="hidden" name="yyyy2" value="<%= yyyy2 %>">
	<input type="hidden" name="mm2" value="<%= mm2 %>">
	<input type="hidden" name="dd2" value="<%= dd2 %>">
	<input type="hidden" name="cdl" value="<%= cdl %>">
	<input type="hidden" name="dplusOver" value="<%= dplusOver %>">
	<input type="hidden" name="dplusLower" value="<%= dplusLower %>">
	<input type="hidden" name="exinmaychulgoday" value="<%= exinmaychulgoday %>">
	<input type="hidden" name="exinneedchulgoday" value="<%= exinneedchulgoday %>">
	<input type="hidden" name="sortby" value="<%= sortby %>">
	<input type="hidden" name="MisendReason" value="<%= MisendReason %>">
	<input type="hidden" name="MisendState" value="<%= MisendState %>">
	<input type="hidden" name="currState" value="<%= currState %>">

	<input type="hidden" name="arrdetailidx" value="">
	<input type="hidden" name="arrbaljudate" value="">

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="2" height="30"><b>미출고 사유 일괄입력</b></td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="10%" height="30">미출고사유</td>
		<td bgcolor="FFFFFF" align="left">
			<select class="select" name="regMisendReason" onChange="jsSetMisendReason(frmMisendInput)">
				<option value=""></option>
				<option value="03">출고지연</option>
				<option value="02">주문제작</option>
				<option value="08">수입</option>
				<option value="09">가구배송</option>
				<option value="04">예약배송</option>
				<option value="10">업체휴가</option>
				<option value="07">고객지정배송</option>
			</select>
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="10%">출고예정일</td>
		<input type="hidden" name="chulgoone_yyyy1" value="">
		<input type="hidden" name="chulgoone_mm1" value="">
		<input type="hidden" name="chulgoone_dd1" value="">
		<td bgcolor="FFFFFF" align="left">
			<input type="radio" name="regbeasongdaytype" value="onedate" onClick="jsSetSMSMailText(frmMisendInput)" checked>
		    <input class="text" type="text" name="chulgooneday" value="<%= (chulgoone_yyyy1 + "-" + chulgoone_mm1 + "-" + chulgoone_dd1) %>" size="10" maxlength="10" onKeyup="SetRegBeasongDayType(0);">
		    <a href="javascript:calendarOpen(frmMisendInput.chulgooneday); SetRegBeasongDayType(0);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
			&nbsp;
			<input type="radio" name="regbeasongdaytype" value="datearea" onClick="jsSetSMSMailText(frmMisendInput)">
			<% DrawChulgoDateBox chulgo_yyyy1, chulgo_mm1 , chulgo_dd1, chulgo_yyyy2, chulgo_mm2, chulgo_dd2 %>
			&nbsp;
			<input type="radio" name="regbeasongdaytype" value="dateneed" onClick="jsSetSMSMailText(frmMisendInput)">
			주문통보일 + <input class="text" type="text" name="regbeasongneedday" size="1" value="" onKeyup="SetRegBeasongDayType(2);"> 일
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="7%">고객안내</td>
		<td bgcolor="FFFFFF" align="left">
			<input name="ckSendSMS" type="checkbox" checked  >SMS발송&nbsp;
			<input name="ckSendEmail" type="checkbox" checked  >MAIL발송&nbsp;
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">SMS<br>발송내용</td>
		<td bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="sendsmsmsg" cols="52" rows="5"></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">MAIL<br>발송내용</td>
		<td bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="sendmailmsg" cols="90" rows="7"></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="FFFFFF" colspan = "2" height="35">
			<input type="button" class="button" value="미출고 사유 일괄저장" onclick="multiMisendInput(frmMisendInput);">
		</td>
	</tr>
	</form>
</table>

<p>
</div>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18">
		<% if Dtype="topN" then %>
		검색결과 : <b><% = ojumun.FTotalCount %></b> (최대 <%= ojumun.FPageSize %>건 까지 검색됩니다.)
		<% else %>
			검색결과 : <b><% = ojumun.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ojumun.FTotalpage %></b>
		<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="chkall" onClick="CheckAll(this)"></td>
		<td>브랜드ID</td>
		<td>사이트</td>
		<td width="70">주문번호</td>
		<td width="55">주문자</td>
		<td width="55">수령인</td>
		<td width="50">상품코드</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="30">CS<br>메모</td>
		<td width="30">수량</td>
		<td width="60">주문통보일<br>(기준일)</td>
		<td width="60">주문확인일</td>
		<td width="35">소요<br>일수</td>
		<td width="50">진행상태</td>
		<td width="75">미출고사유</td>
		<td width="60">출고예정일</td>
		<td width="65">처리구분</td>
		<td width="35">상세<br>정보</td>
	</tr>
	<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="18" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<form name="frmBuyPrc_<%= ix %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="FFFFFF">
	<% end if %>
	<%
	IsDisableThis = (ojumun.FMasterItemList(ix).FMisendReason="05")
	if (IsDisableThis = False) and (Not IsNULL(ojumun.FMasterItemList(ix).FMisendipgodate)) then
		if DateDiff("d", ojumun.FMasterItemList(ix).FMisendipgodate, Now()) < 0 then
			IsDisableThis = True
		end if
	end if
	%>
		<td><input type="checkbox" name="detailidx" id="detailidx_<%= ix %>" value="<%= ojumun.FMasterItemList(ix).Fdetailidx %>" <% if IsDisableThis = True then %>disabled<%end if %>></td>
		<input type="hidden" name="baljudate" value="<%= Left(ojumun.FMasterItemList(ix).Fbaljudate,10) %>">
		<td>
			<a href="javascript:searchByMakerId(frm, '<%= ojumun.FMasterItemList(ix).FMakerid %>')">
				<%= ojumun.FMasterItemList(ix).FMakerid %>
			</a>
		</td>
		<td>
			<% if (ojumun.FMasterItemList(ix).Fsitename <> "10x10") then %>
				<%= ojumun.FMasterItemList(ix).Fsitename %>
			<% end if %>
		</td>
		<td><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= ojumun.FMasterItemList(ix).FOrderSerial %>')" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqname %></td>
		<td>

			<a href="javascript:searchByItemId(frm, <%= ojumun.FMasterItemList(ix).FItemid %>)">
				<%= ojumun.FMasterItemList(ix).FItemid %>
			</a>
		</td>
		<td align="left">
			<a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
				<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
					<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
				<% end if %>
		</td>
		<td>
			<% if (ojumun.FMasterItemList(ix).FcsMemoCnt > 0) then %>
				V
			<% end if %>
		</td>
		<td><%= ojumun.FMasterItemList(ix).FItemcnt %></td>

		<td><%= Left(ojumun.FMasterItemList(ix).Fbaljudate,10) %></td>
		<td><%= Left(ojumun.FMasterItemList(ix).Fupcheconfirmdate,10) %></td>
		<td><%= ojumun.FMasterItemList(ix).getNewBeasongDPlusDateStr %></td>
		<td>
		    <% if (detailstate="MOO") then %>

		    <% else %>
    			<% if ojumun.FMasterItemList(ix).FCurrstate = 0 then %>
    			<font color="blue">결제완료</font>
    			<% elseif ojumun.FMasterItemList(ix).FCurrstate = 2 then %>
    			<font color="#000000">주문통보</font>
    			<% elseif ojumun.FMasterItemList(ix).FCurrstate = 3 then %>
    			<font color="#CC9933">주문확인</font>
    			<% elseif ojumun.FMasterItemList(ix).FCurrstate = 7 then %>
    			<font color="#FF0000">출고완료</font>
    			<% end if %>
			<% end if %>
		</td>
		<td>
		    <%= ojumun.FMasterItemList(ix).getMisendText %>
		    <% if not IsNULL(ojumun.FMasterItemList(ix).Fmisendregdate) then %>
		    <br>(<%= Left(ojumun.FMasterItemList(ix).Fmisendregdate,10) %>)
		    <% end if %>
		</td>
		<td><%= ojumun.FMasterItemList(ix).FMisendipgodate %></td>
		<td><%= ojumun.FMasterItemList(ix).getMisendStateText %></td>
		<td>
			<a href="javascript:misendmaster('<%= ojumun.FMasterItemList(ix).FOrderSerial %>');"><img src="/images/icon_search.jpg" border="0"></a>
		</td>
	</tr>
	</form>
	<% next %>
<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18" align="center">
		<% if Dtype="topN" then %>
		최대 <%= ojumun.FPageSize %>건 까지 검색됩니다.
		<% else %>
    		<% if ojumun.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>
    		<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
    			<% if ix>ojumun.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(ix) then %>
    			<font color="red">[<%= ix %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
    			<% end if %>
    		<% next %>

    		<% if ojumun.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
    	<% end if %>
		</td>
	</tr>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
