<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/mdMenu/mdMainCls.asp"-->
<%

dim i, j, k
dim sqlStr, tmpResultArr
dim tmpDOM, tmpXML

'==============================================================================
'// 시간 체크
dim lastPageTime, pageElapsedTime
lastPageTime = Timer

function checkAndWriteElapsedTime(memo)
	pageElapsedTime = Timer - lastPageTime
	lastPageTime = Timer
	response.write "<!-- Page Execute Time Check : " & FormatNumber(pageElapsedTime, 4) & " : " & memo & " -->" & vbCrLf
end function


'==============================================================================
' 전시 카테고리
Dim dispCate, dispCateIndex, dispCateName
dispCate	= req("dispCate", "")

Dim mdMainUserID
mdMainUserID	= req("mdMainUserID", "")

if (dispCate <> "") then
	dispCateName = GetNameFromDispCateCode(dispCate)
end if

function OpenDataIfCateSelected(dispCate)
	OpenDataIfCateSelected = " style='display:none' "
	if (dispCate <> "") then
		OpenDataIfCateSelected = ""
	end if
end function

function OpenDataIfUsernameSelected(username)
	OpenDataIfUsernameSelected = " style='display:none' "
	if (username <> "") then
		OpenDataIfUsernameSelected = ""
	end if
end function


'==============================================================================
'// 업데이트 정보

dim updateNeedInfo
set updateNeedInfo = new ClsIsUpdateNeedItem

if Not IsArray(Application("mdWillFinishEvent")) or (Trim(application("mdTimeWillFinishEvent")) = "") or (DateDiff("s", application("mdTimeWillFinishEvent"), Now() ) > 3 * 60 * 60) then
	Application("mdTimeWillFinishEvent") = Now()
	updateNeedInfo.FwillFinishEvent = True
end if

if Not IsArray(Application("mdEventCount")) or (Trim(application("mdTimeEventCount")) = "") or (DateDiff("s", application("mdTimeEventCount"), Now() ) > 3 * 60 * 60) then
	Application("mdTimeEventCount") = Now()
	updateNeedInfo.FEventCount = True
end if

if Not IsArray(Application("mdUpcheRequest")) or (Trim(application("mdTimeUpcheRequest")) = "") or (DateDiff("s", application("mdTimeUpcheRequest"), Now() ) > 3 * 60 * 60) then
	Application("mdTimeUpcheRequest") = Now()
	updateNeedInfo.FupcheRequest = True
end if

if Not IsArray(Application("mdItemRequest")) or (Trim(application("mdTimeItemRequest")) = "") or (DateDiff("s", application("mdTimeItemRequest"), Now() ) > 3 * 60 * 60) then
	Application("mdTimeItemRequest") = Now()
	updateNeedInfo.FitemRequest = True
end if

if Not IsArray(Application("mdItemSellRequest")) or (Trim(application("mdTimeItemSellRequest")) = "") or (DateDiff("s", application("mdTimeItemSellRequest"), Now() ) > 3 * 60 * 60) then
	Application("mdTimeItemSellRequest") = Now()
	updateNeedInfo.FItemSellRequest = True
end if

if Not IsArray(Application("mdBrandRequest")) or (Trim(application("mdTimeBrandRequest")) = "") or (DateDiff("s", application("mdTimeBrandRequest"), Now() ) > 3 * 60 * 60) then
	Application("mdTimeBrandRequest") = Now()
	updateNeedInfo.FBrandRequest = True
end if

if Not IsArray(Application("mdEventPrize")) or (Trim(application("mdTimeEventPrize")) = "") or (DateDiff("s", application("mdTimeEventPrize"), Now() ) > 3 * 60 * 60) then
	Application("mdTimeEventPrize") = Now()
	updateNeedInfo.FEventPrize = True
end if


'==============================================================================
'// 종료임박 이벤트
dim willFinishEventArr()

if (updateNeedInfo.FwillFinishEvent = True) then
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_Event_WillFinishCnt] '' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		Application("mdWillFinishEvent") = rsget.getRows()
	end if
	rsget.close
end if

tmpResultArr = Application("mdWillFinishEvent")

redim willFinishEventArr(UBound(tmpResultArr, 2))
for i = 0 to UBound(tmpResultArr, 2)
	set willFinishEventArr(i) = new ClsWillFinishEventItem

	willFinishEventArr(i).FNormalCnt = tmpResultArr(2, i)
	willFinishEventArr(i).FAppCnt = tmpResultArr(3, i)

	willFinishEventArr(i).FNormal101Cnt = tmpResultArr(4, i)
	willFinishEventArr(i).FNormal102Cnt = tmpResultArr(5, i)
	willFinishEventArr(i).FNormal103Cnt = tmpResultArr(6, i)
	willFinishEventArr(i).FNormal104Cnt = tmpResultArr(7, i)
	willFinishEventArr(i).FNormal114Cnt = tmpResultArr(8, i)
	willFinishEventArr(i).FNormal106Cnt = tmpResultArr(9, i)
	willFinishEventArr(i).FNormal112Cnt = tmpResultArr(10, i)
	willFinishEventArr(i).FNormal113Cnt = tmpResultArr(11, i)
	willFinishEventArr(i).FNormal115Cnt = tmpResultArr(12, i)
	willFinishEventArr(i).FNormal110Cnt = tmpResultArr(13, i)
	willFinishEventArr(i).FNormal111Cnt = tmpResultArr(14, i)

	willFinishEventArr(i).FApp101Cnt = tmpResultArr(15, i)
	willFinishEventArr(i).FApp102Cnt = tmpResultArr(16, i)
	willFinishEventArr(i).FApp103Cnt = tmpResultArr(17, i)
	willFinishEventArr(i).FApp104Cnt = tmpResultArr(18, i)
	willFinishEventArr(i).FApp114Cnt = tmpResultArr(19, i)
	willFinishEventArr(i).FApp106Cnt = tmpResultArr(20, i)
	willFinishEventArr(i).FApp112Cnt = tmpResultArr(21, i)
	willFinishEventArr(i).FApp113Cnt = tmpResultArr(22, i)
	willFinishEventArr(i).FApp115Cnt = tmpResultArr(23, i)
	willFinishEventArr(i).FApp110Cnt = tmpResultArr(24, i)
	willFinishEventArr(i).FApp111Cnt = tmpResultArr(25, i)

	if (dispCate <> "") then
		dispCateIndex = GetIndexFromDispCateCode(dispCate)

		if (dispCateIndex >= 0) then
			willFinishEventArr(i).FNormalCnt = tmpResultArr((4 + dispCateIndex), i)
			willFinishEventArr(i).FAppCnt = tmpResultArr((15 + dispCateIndex), i)
		end if
	end if
next

set tmpResultArr = Nothing


'==============================================================================
'// 이벤트 통계
dim EventCount

if (updateNeedInfo.FEventCount = True) then
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_Event_Cnt] "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		Application("mdEventCount") = rsget.getRows()
	end if
	rsget.close
end if

tmpResultArr = Application("mdEventCount")

set EventCount = new ClsEventCountItem

EventCount.FtotCount = tmpResultArr(0, 0)
EventCount.Fstate0 = tmpResultArr(1, 0)
EventCount.Fstate1 = tmpResultArr(2, 0)
EventCount.Fstate2 = tmpResultArr(3, 0)
EventCount.Fstate3 = tmpResultArr(4, 0)
EventCount.Fstate5 = tmpResultArr(5, 0)
EventCount.Fstate7 = tmpResultArr(6, 0)
EventCount.Fstate6 = tmpResultArr(7, 0)
EventCount.Fstate9 = tmpResultArr(8, 0)


'==============================================================================
'// 입점 및 계약관리
dim UpcheRequest
dim companyRequest, companyRequestArr()
dim companyContract1, companyContract1Arr()
dim companyContract3, companyContract3Arr()
dim companyInfoModifyReq, companyInfoModifyReqArr()

if (updateNeedInfo.FupcheRequest = True) then
	'// 1. 업체입점문의
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_CompanyRequest_Cnt] "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		companyRequest = rsget.getRows()
	else
		companyRequest = Array()
	end if
	rsget.close

	'// 2. 업체계약관리(업체 오픈)
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_CompanyContract_Cnt] 1 "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		companyContract1 = rsget.getRows()
	else
		companyContract1 = Array()
	end if
	rsget.close

	'// 3. 업체계약관리(업체 확인)
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_CompanyContract_Cnt] 3 "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		companyContract3 = rsget.getRows()
	else
		companyContract3 = Array()
	end if
	rsget.close

	'// 4. 업체정보 등록(변경) 신청
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_CompanyInfoModifyReq_Cnt] 1 "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		companyInfoModifyReq = rsget.getRows()
	else
		companyInfoModifyReq = Array()
	end if
	rsget.close

	Application("mdUpcheRequest") = Array(companyRequest, companyContract1, companyContract3, companyInfoModifyReq)
end if


'// 1. 업체입점문의
tmpResultArr = Application("mdUpcheRequest")(0)

if (UBound(tmpResultArr) <= 0) then
	redim companyRequestArr(0)
else
	redim companyRequestArr(UBound(tmpResultArr, 2))
	for i = 0 to UBound(tmpResultArr, 2)
		set companyRequestArr(i) = new ClsCompanyRequestItem

		companyRequestArr(i).FdispCate = tmpResultArr(0, i)
		companyRequestArr(i).FCateName = tmpResultArr(1, i)
		companyRequestArr(i).Fcount = tmpResultArr(2, i)
	next
end if


'// 2. 업체계약관리(업체 오픈)
tmpResultArr = Application("mdUpcheRequest")(1)

redim companyContract1Arr(UBound(tmpResultArr, 2))
for i = 0 to UBound(tmpResultArr, 2)
	set companyContract1Arr(i) = new ClsCompanyContractItem

	companyContract1Arr(i).FsendUserID = tmpResultArr(0, i)
	companyContract1Arr(i).Fusername = tmpResultArr(1, i)
	companyContract1Arr(i).Fcount = tmpResultArr(2, i)

	if (companyContract1Arr(i).FsendUserID = "") then
		companyContract1Arr(i).FsendUserID = "xxxxxx"
	end if
next


'// 3. 업체계약관리(업체 확인)
tmpResultArr = Application("mdUpcheRequest")(2)

redim companyContract3Arr(UBound(tmpResultArr, 2))
for i = 0 to UBound(tmpResultArr, 2)
	set companyContract3Arr(i) = new ClsCompanyContractItem

	companyContract3Arr(i).FsendUserID = tmpResultArr(0, i)
	companyContract3Arr(i).Fusername = tmpResultArr(1, i)
	companyContract3Arr(i).Fcount = tmpResultArr(2, i)

	if (companyContract3Arr(i).FsendUserID = "") then
		companyContract3Arr(i).FsendUserID = "xxxxxx"
	end if
next


'// 4. 업체정보 등록(변경) 신청
tmpResultArr = Application("mdUpcheRequest")(3)

if (UBound(tmpResultArr) <= 0) then
	redim companyInfoModifyReqArr(0)
else
	redim companyInfoModifyReqArr(UBound(tmpResultArr, 2))
	for i = 0 to UBound(tmpResultArr, 2)
		set companyInfoModifyReqArr(i) = new ClsCompanyInfoModifyReqItem

		companyInfoModifyReqArr(i).FuserID = tmpResultArr(0, i)
		companyInfoModifyReqArr(i).Fusername = tmpResultArr(1, i)
		companyInfoModifyReqArr(i).Fcount = tmpResultArr(2, i)

		if (companyInfoModifyReqArr(i).FuserID = "") then
			companyInfoModifyReqArr(i).FuserID = "xxxxxx"
		end if
	next
end if


'==============================================================================
'// 상품정보
dim ItemRegRequestCount, ItemRegRequestCountArr()
dim UpcheItemModiRequestCount, UpcheItemModiRequestCountArr()

if (updateNeedInfo.FitemRequest = True) then
	'// 1. 승인대기 상품
	sqlStr = " [db_temp].[dbo].[sp_Ten_wait_item_getSummrayList] 'C', '1','CA','' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		ItemRegRequestCount = rsget.getRows()
	else
		ItemRegRequestCount = Array()
	end if
	rsget.close

	'// 2. 업배상품 승인대기
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_UpcheItemModifyReq_Cnt] "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		UpcheItemModiRequestCount = rsget.getRows()
	else
		UpcheItemModiRequestCount = Array()
	end if
	rsget.close

	Application("mdItemRequest") = Array(ItemRegRequestCount, UpcheItemModiRequestCount)
end if

'// 1. 승인대기 상품
tmpResultArr = Application("mdItemRequest")(0)

if (UBound(tmpResultArr) <= 0) then
	redim ItemRegRequestCountArr(0)
else
	redim ItemRegRequestCountArr(UBound(tmpResultArr, 2))
	for i = 0 to UBound(tmpResultArr, 2)
		set ItemRegRequestCountArr(i) = new ClsItemRegRequestCountItem

		ItemRegRequestCountArr(i).FcateCode = tmpResultArr(0, i)
		ItemRegRequestCountArr(i).FcateName = tmpResultArr(1, i)
		ItemRegRequestCountArr(i).Fcount1 = tmpResultArr(3, i)
		ItemRegRequestCountArr(i).Fcount5 = tmpResultArr(4, i)
	next
end if

'// 2. 업배상품 승인대기
tmpResultArr = Application("mdItemRequest")(1)

if (UBound(tmpResultArr) <= 0) then
	redim UpcheItemModiRequestCountArr(0)
else
	redim UpcheItemModiRequestCountArr(UBound(tmpResultArr, 2))
	for i = 0 to UBound(tmpResultArr, 2)
		set UpcheItemModiRequestCountArr(i) = new ClsDispCateItem

		UpcheItemModiRequestCountArr(i).FdispCateCode = tmpResultArr(0, i)
		UpcheItemModiRequestCountArr(i).FdispCateName = tmpResultArr(1, i)
		UpcheItemModiRequestCountArr(i).Fcount = tmpResultArr(2, i)

		if (UpcheItemModiRequestCountArr(i).FdispCateCode = 0) then
			UpcheItemModiRequestCountArr(i).FdispCateCode = ""
		end if
	next
end if

'==============================================================================
'// 상품정보
dim IpgpNotSellCount, IpgpNotSellCountArr()

if (updateNeedInfo.FItemSellRequest = True) then

	'// 3. 판매대기 상품목록
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_IpgoNotSellItem_Cnt] "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		IpgpNotSellCount = rsget.getRows()
	else
		IpgpNotSellCount = Array()
	end if
	rsget.close

	Application("mdItemSellRequest") = Array(IpgpNotSellCount)
end if

'// 3. 판매대기 상품목록
tmpResultArr = Application("mdItemSellRequest")(0)

if (UBound(tmpResultArr) <= 0) then
	redim IpgpNotSellCountArr(0)
else
	redim IpgpNotSellCountArr(UBound(tmpResultArr, 2))
	for i = 0 to UBound(tmpResultArr, 2)
		set IpgpNotSellCountArr(i) = new ClsDispCateItem

		IpgpNotSellCountArr(i).FdispCateCode = tmpResultArr(0, i)
		IpgpNotSellCountArr(i).FdispCateName = tmpResultArr(1, i)
		IpgpNotSellCountArr(i).Fcount = tmpResultArr(2, i)
	next
end if


'==============================================================================
'// 브랜드 정보
dim BrandLookBookCount, BrandLookBookCountArr()
dim BrandShopCollectionCount, BrandShopCollectionCountArr()

if (updateNeedInfo.FBrandRequest = True) then
	'// 1. LOOKBOOK
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_BrandLookBook_Cnt] "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		BrandLookBookCount = rsget.getRows()
	else
		BrandLookBookCount = Array()
	end if
	rsget.close

	'// 2. SHOP_collection
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_BrandShopCollection_Cnt] "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		BrandShopCollectionCount = rsget.getRows()
	else
		BrandShopCollectionCount = Array()
	end if
	rsget.close

	Application("mdBrandRequest") = Array(BrandLookBookCount, BrandShopCollectionCount)
end if

'// 1. LOOKBOOK
tmpResultArr = Application("mdBrandRequest")(0)

if (UBound(tmpResultArr) <= 0) then
	redim BrandLookBookCountArr(0)
else
	redim BrandLookBookCountArr(UBound(tmpResultArr, 2))
	for i = 0 to UBound(tmpResultArr, 2)
		set BrandLookBookCountArr(i) = new ClsDispCateItem

		BrandLookBookCountArr(i).FdispCateCode = tmpResultArr(0, i)
		BrandLookBookCountArr(i).FdispCateName = tmpResultArr(1, i)
		BrandLookBookCountArr(i).Fcount = tmpResultArr(2, i)
	next
end if

'// 2. SHOP_collection
tmpResultArr = Application("mdBrandRequest")(1)

if (UBound(tmpResultArr) <= 0) then
	redim BrandShopCollectionCountArr(0)
else
	redim BrandShopCollectionCountArr(UBound(tmpResultArr, 2))
	for i = 0 to UBound(tmpResultArr, 2)
		set BrandShopCollectionCountArr(i) = new ClsDispCateItem

		BrandShopCollectionCountArr(i).FdispCateCode = tmpResultArr(0, i)
		BrandShopCollectionCountArr(i).FdispCateName = tmpResultArr(1, i)
		BrandShopCollectionCountArr(i).Fcount = tmpResultArr(2, i)
	next
end if


'==============================================================================
'// 당첨 이벤트
dim EventPrize, EventPrizeArr()

if (updateNeedInfo.FEventPrize = True) then

	'// 1. 당첨 이벤트
	sqlStr = " [db_sitemaster].[dbo].[usp_Ten_EventPrize_Cnt] "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		EventPrize = rsget.getRows()
	else
		EventPrize = Array()
	end if
	rsget.close

	Application("mdEventPrize") = Array(EventPrize)
end if

'// 1. 당첨 이벤트
tmpResultArr = Application("mdEventPrize")(0)

if (UBound(tmpResultArr) <= 0) then
	redim EventPrizeArr(0)
else
	redim EventPrizeArr(UBound(tmpResultArr, 2))
	for i = 0 to UBound(tmpResultArr, 2)
		set EventPrizeArr(i) = new CEventPrizeItem

		EventPrizeArr(i).FeventCode = tmpResultArr(0, i)
		EventPrizeArr(i).FeventName = tmpResultArr(1, i)
		EventPrizeArr(i).FeventKind = tmpResultArr(2, i)
		EventPrizeArr(i).FuserID = tmpResultArr(3, i)
		EventPrizeArr(i).FuserName = tmpResultArr(4, i)
		EventPrizeArr(i).FdDay = tmpResultArr(5, i)
		EventPrizeArr(i).FprizeDay = tmpResultArr(6, i)
	next
end if

%>
<script language="JavaScript" src="/cscenter/js/convert.date.js"></script>
<script language='javascript'>

function showHideTR(id) {
	tr = document.getElementsByTagName("tr");

	for (var i = 0; i < tr.length; i++) {
		if (tr[i].id == id) {
			if ( tr[i].style.display=="none" ) {
				tr[i].style.display = "";
			} else {
				tr[i].style.display = "none";
			}
		}
	}
}

function popOpenEvent(dispCate, eventstate, selDate, iSD, iED, eventkind) {
    var window_width = 1280;
    var window_height = 960;

	var url = "/admin/eventmanage/event/?menupos=870";
	url = url + "&eventstate=" + eventstate;
	url = url + "&selDate=" + selDate;
	url = url + "&iSD=" + iSD;
	url = url + "&iED=" + iED;
	url = url + "&eventkind=" + eventkind;
	url = url + "&disp=" + dispCate;

    var popwin = window.open(url,"popOpenEvent","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function RefreshData(v) {
	var frm = document.frmAct;

	frm.mode.value = "RefreshData";
	frm.mdTime.value = v;
	frm.submit();
}

var mdTimeWillFinishEvent = new Date(getDateFromFormat("<%= application("mdTimeWillFinishEvent") %>", "yyyy-MM-dd a h:mm:ss"));
var mdTimeEventCount = new Date(getDateFromFormat("<%= application("mdTimeEventCount") %>", "yyyy-MM-dd a h:mm:ss"));
var mdTimeUpcheRequest = new Date(getDateFromFormat("<%= application("mdTimeUpcheRequest") %>", "yyyy-MM-dd a h:mm:ss"));
var mdTimeItemRequest = new Date(getDateFromFormat("<%= application("mdTimeItemRequest") %>", "yyyy-MM-dd a h:mm:ss"));
var mdTimeItemSellRequest = new Date(getDateFromFormat("<%= application("mdTimeItemSellRequest") %>", "yyyy-MM-dd a h:mm:ss"));
var mdTimeBrandRequest = new Date(getDateFromFormat("<%= application("mdTimeBrandRequest") %>", "yyyy-MM-dd a h:mm:ss"));
var mdTimeEventPrize = new Date(getDateFromFormat("<%= application("mdTimeEventPrize") %>", "yyyy-MM-dd a h:mm:ss"));

function DisplayClock() {
	var v = new Date();

	var objTimeWillFinishEvent = document.getElementById("objTimeWillFinishEvent");
	var objTimeEventCount = document.getElementById("objTimeEventCount");
	var objTimeUpcheRequest = document.getElementById("objTimeUpcheRequest");
	var objTimeItemRequest = document.getElementById("objTimeItemRequest");
	var objTimeItemSellRequest = document.getElementById("objTimeItemSellRequest");
	var objTimeBrandRequest = document.getElementById("objTimeBrandRequest");
	var objTimeEventPrize = document.getElementById("objTimeEventPrize");

	objTimeWillFinishEvent.innerHTML = GetDateDiffString(v.getTime() - mdTimeWillFinishEvent.getTime());
	objTimeEventCount.innerHTML = GetDateDiffString(v.getTime() - mdTimeEventCount.getTime());
	objTimeUpcheRequest.innerHTML = GetDateDiffString(v.getTime() - mdTimeUpcheRequest.getTime());
	objTimeItemRequest.innerHTML = GetDateDiffString(v.getTime() - mdTimeItemRequest.getTime());
	objTimeItemSellRequest.innerHTML = GetDateDiffString(v.getTime() - mdTimeItemSellRequest.getTime());
	objTimeBrandRequest.innerHTML = GetDateDiffString(v.getTime() - mdTimeBrandRequest.getTime());
	objTimeEventPrize.innerHTML = GetDateDiffString(v.getTime() - mdTimeEventPrize.getTime());

	setTimeout('DisplayClock();','1000');
}

function GetDateDiffString(v) {
	var result = "";

	if (v < (60 * 1000)) {
		v = v / 1000;
		result = parseInt(v) + "초 전";
	} else if (v < (60 * 60 * 1000)) {
		v = v / (60 * 1000);
		result = parseInt(v) + "분 전";
	} else {
		result =  "1시간 전";
	}

	return result;
}

window.onload = function() {
	DisplayClock();
}

/*
function GetCategoryName(dispCateCode) {
	var ret = "";
	var item = document.getElementById('dispCate');

	if (item == undefined) {
		return ret;
	}

	if (item.value == "") {
		return ret;
	}

	var selIndex = item.selectedIndex;

	return item.options[selIndex].text;
}
*/

</script>

<% Call checkAndWriteElapsedTime("010") %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="33%" valign="top">
	    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">

        <tr valign="top">
            <td>
				<!--  aaaa -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>상품관리 - 판매대기</b>
										(<span id="objTimeItemSellRequest"></span>) <a href="javascript:RefreshData('mdTimeItemSellRequest')"><img src="/images/icon_reload.gif" border="0"></a>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
										&nbsp;
            						</td>
            					</tr>
            					<tr height="25">
            						<td>판매 대기</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="/admin/shopmaster/item_new_list.asp?menupos=653" target="_blank">판매대기상품목록</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% for i = 0 to UBound(IpgpNotSellCountArr) %>
								<% if (dispCate = "") or (dispCate = IpgpNotSellCountArr(i).FdispCateCode) then %>
            					<tr height="25" id="IpgpNotSellCount">
            						<td>
										&nbsp;&nbsp;&nbsp; - <%= IpgpNotSellCountArr(i).FdispCateName %>
									</td>
            						<td align="right">
										<a href="/admin/shopmaster/item_new_list.asp?menupos=653&disp=<%= IpgpNotSellCountArr(i).FdispCateCode %>" target="_blank">
											<b><%= IpgpNotSellCountArr(i).Fcount %></b> 건
										</a>
									</td>
            					</tr>
								<% end if %>
								<% next %>
            				</table>
            			</td>
            		</tr>
            	</table>
        	    <!--  aaaa -->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
				<!--  aaaa -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>상품관리 - 승인대기</b>
										(<span id="objTimeItemRequest"></span>) <a href="javascript:RefreshData('mdTimeItemRequest')"><img src="/images/icon_reload.gif" border="0"></a>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
										&nbsp;
            						</td>
            					</tr>
            					<tr height="25">
            						<td>승인 대기</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>
										&nbsp; * <a href="/admin/itemmaster/item_confirm_master.asp?menupos=121" target="_blank">승인대기 상품목록</a>
										<a href="javascript:showHideTR('ItemRegRequestCount');">[펼치기]</a>
									</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% for i = 0 to UBound(ItemRegRequestCountArr) %>
								<% if (dispCate = "") or IsNull(ItemRegRequestCountArr(i).FcateCode) or (dispCate = ("" &ItemRegRequestCountArr(i).FcateCode)) then %>
            					<tr height="25" id="ItemRegRequestCount">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										<% if ItemRegRequestCountArr(i).FcateName <> "" then %>
											<%= ItemRegRequestCountArr(i).FcateName %>
										<% else %>
											<font color="red"><b>전시 카테고리 미지정</b></font>
										<% end if %>
									</td>
            						<td align="right">
										<% if IsNull(ItemRegRequestCountArr(i).FcateCode) then %>
											대기 : <a href="/admin/itemmaster/item_confirm.asp?sLT=C&makerid=&onlyNotSet=Y&sCS=1" target="_blank">
												<b><%= ItemRegRequestCountArr(i).Fcount1 %></b> 건
											</a>
											/
											재등록 : <a href="/admin/itemmaster/item_confirm.asp?sLT=C&makerid=&onlyNotSet=Y&sCS=5" target="_blank">
												<b><%= ItemRegRequestCountArr(i).Fcount5 %></b> 건
											</a>
										<% else %>
											대기 : <a href="/admin/itemmaster/item_confirm.asp?sLT=C&makerid=&disp=<%= ItemRegRequestCountArr(i).FcateCode %>&sCS=1" target="_blank">
												<b><%= ItemRegRequestCountArr(i).Fcount1 %></b> 건
											</a>
											/
											재등록 : <a href="/admin/itemmaster/item_confirm.asp?sLT=C&makerid=&disp=<%= ItemRegRequestCountArr(i).FcateCode %>&sCS=5" target="_blank">
												<b><%= ItemRegRequestCountArr(i).Fcount5 %></b> 건
											</a>
										<% end if %>
									</td>
            					</tr>
								<% end if %>
								<% next %>
								<tr height="25">
            						<td>&nbsp; * <a href="/admin/itemmaster/item_modReq_confirm.asp?menupos=1660" target="_blank">업배상품 승인대기 목록</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% for i = 0 to UBound(UpcheItemModiRequestCountArr) %>
								<% if (dispCate = "") or (dispCate = CStr(UpcheItemModiRequestCountArr(i).FdispCateCode)) then %>
            					<tr height="25" id="UpcheItemModiRequestCount">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										<% if UpcheItemModiRequestCountArr(i).FdispCateCode <> "" then %>
											<%= UpcheItemModiRequestCountArr(i).FdispCateName %>
										<% else %>
											<font color="red"><b>전시 카테고리 미지정</b></font>
										<% end if %>
									</td>
            						<td align="right">
										<% if UpcheItemModiRequestCountArr(i).FdispCateCode <> "" then %>
											<a href="/admin/itemmaster/item_modReq_confirm.asp?menupos=1660&disp=<%= UpcheItemModiRequestCountArr(i).FdispCateCode %>" target="_blank">
												<b><%= UpcheItemModiRequestCountArr(i).Fcount %></b> 건
											</a>
										<% else %>
											<a href="/admin/itemmaster/item_modReq_confirm.asp?menupos=1660&onlyNotSet=Y" target="_blank">
												<b><%= UpcheItemModiRequestCountArr(i).Fcount %></b> 건
											</a>
										<% end if %>

									</td>
            					</tr>
								<% end if %>
								<% next %>
            				</table>
            			</td>
            		</tr>
            	</table>
        	    <!--  aaaa -->
			</td>
		</tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

		</table>
	</td>
	<td width="10"></td>
	<td width="33%" valign="top">
	    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">

		<tr valign="top">
            <td>
				<!--  aaaa -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>입점 및 계약관리</b>
										(<span id="objTimeUpcheRequest"></span>) <a href="javascript:RefreshData('mdTimeUpcheRequest')"><img src="/images/icon_reload.gif" border="0"></a>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
										&nbsp;
            						</td>
            					</tr>
            					<tr height="25">
            						<td>
										<a href="/admin/board/upche/req_list.asp?menupos=1069&disp=<%= dispCate %>" target="_blank">업체입점문의(오늘 접수건)</a>
										<a href="javascript:showHideTR('UpcheRequest');">[펼치기]</a>
									</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% if (UBound(companyRequestArr) > 0) then %>
								<% for i = 0 to UBound(companyRequestArr) %>
								<% if (dispCate = "") or (companyRequestArr(i).FdispCate = "0") or (dispCate = companyRequestArr(i).FdispCate) then %>
            					<tr height="25" id="UpcheRequest" <%= OpenDataIfCateSelected(dispCate) %> >
            						<td>
										&nbsp; - <%= companyRequestArr(i).FCateName %>
									</td>
            						<td align="right">
										<a href="/admin/board/upche/req_list.asp?menupos=1069&disp=<%= companyRequestArr(i).FdispCate %>" target="_blank">
											<b><%= companyRequestArr(i).Fcount %></b> 건
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<% end if %>
								<% next %>
								<% end if %>
								<tr height="25">
            						<td>
										<a href="/admin/member/contract/ctrList.asp?menupos=1619" target="_blank">업체계약관리(발송자 기준)</a>
									</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>
										&nbsp; * <a href="/admin/member/contract/ctrList.asp?menupos=1619&ContractState=1" target="_blank">업체 오픈</a>
										<a href="javascript:showHideTR('companyContract1');">[펼치기]</a>
									</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% for i = 0 to UBound(companyContract1Arr) %>
								<% if (mdMainUserID = "") or (mdMainUserID = companyContract1Arr(i).Fusername) or (companyContract1Arr(i).Fusername = "퇴사자") then %>
            					<tr height="25" id="companyContract1" <%= OpenDataIfUsernameSelected(mdMainUserID) %> >
            						<td>
										&nbsp;&nbsp;&nbsp; - <%= companyContract1Arr(i).Fusername %>
									</td>
            						<td align="right">
										<a href="/admin/member/contract/ctrList.asp?menupos=1619&ContractState=1&sendUserID=<%= companyContract1Arr(i).FsendUserID %>" target="_blank">
											<b><%= companyContract1Arr(i).Fcount %></b> 건
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<% end if %>
								<% next %>
            					<tr height="25">
            						<td>
										&nbsp; * <a href="/admin/member/contract/ctrList.asp?menupos=1619&ContractState=3" target="_blank">업체 확인</a>
										<a href="javascript:showHideTR('companyContract3');">[펼치기]</a>
									</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% for i = 0 to UBound(companyContract3Arr) %>
								<% if (mdMainUserID = "") or (mdMainUserID = companyContract3Arr(i).Fusername) or (companyContract3Arr(i).Fusername = "퇴사자") then %>
            					<tr height="25" id="companyContract3" <%= OpenDataIfUsernameSelected(mdMainUserID) %> >
            						<td>
										&nbsp;&nbsp;&nbsp; - <%= companyContract3Arr(i).Fusername %>
									</td>
            						<td align="right">
										<a href="/admin/member/contract/ctrList.asp?menupos=1619&ContractState=3&sendUserID=<%= companyContract3Arr(i).FsendUserID %>" target="_blank">
											<b><%= companyContract3Arr(i).Fcount %></b> 건
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<% end if %>
								<% next %>
            					<tr height="25">
            						<td>
										<a href="/admin/member/partner/?menupos=1453&reqstatus=1" target="_blank">업체정보 등록(변경) 신청</a>
									</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% if (UBound(companyInfoModifyReqArr) <= 0) then %>
            					<tr height="25" id="companyInfoModifyReq">
            						<td>
										&nbsp;&nbsp;&nbsp; - 신청중인 건이 없습니다.
									</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% else %>
								<% for i = 0 to UBound(companyInfoModifyReqArr) %>
            					<tr height="25" id="companyInfoModifyReq">
            						<td>
										&nbsp; - <%= companyInfoModifyReqArr(i).Fusername %>
									</td>
            						<td align="right">
										<a href="/admin/member/partner/?menupos=1453&reqstatus=1&reqname=<%= companyInfoModifyReqArr(i).Fusername %>" target="_blank">
											<b><%= companyInfoModifyReqArr(i).Fcount %></b> 건
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<% next %>
								<% end if %>
            				</table>
            			</td>
            		</tr>
            	</table>
        	    <!--  aaaa -->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
				<!--  aaaa -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>브랜드관리</b>
										(<span id="objTimeBrandRequest"></span>) <a href="javascript:RefreshData('mdTimeBrandRequest')"><img src="/images/icon_reload.gif" border="0"></a>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
										&nbsp;
            						</td>
            					</tr>
								<tr height="25">
            						<td>LOOKBOOK</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% for i = 0 to UBound(BrandLookBookCountArr) %>
								<% if (dispCate = "") or (dispCate = BrandLookBookCountArr(i).FdispCateCode) or (BrandLookBookCountArr(i).FdispCateCode = "") then %>
            					<tr height="25" id="BrandLookBookCount">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										<% if BrandLookBookCountArr(i).FdispCateCode <> "" then %>
											<%= BrandLookBookCountArr(i).FdispCateName %>
										<% else %>
											미지정
										<% end if %>
									</td>
            						<td align="right">
										<a href="/admin/brand/lookbook/index.asp?menupos=1599&standardCateCode=<%= BrandLookBookCountArr(i).FdispCateCode %>" target="_blank">
											<b><%= BrandLookBookCountArr(i).Fcount %></b> 건
										</a>
									</td>
            					</tr>
								<% end if %>
								<% next %>
								<tr height="25">
            						<td>SHOP Collection</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% for i = 0 to UBound(BrandShopCollectionCountArr) %>
								<% if (dispCate = "") or (dispCate = BrandShopCollectionCountArr(i).FdispCateCode) or (BrandShopCollectionCountArr(i).FdispCateCode = "") then %>
            					<tr height="25" id="BrandLookBookCount">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										<% if BrandShopCollectionCountArr(i).FdispCateCode <> "" then %>
											<%= BrandShopCollectionCountArr(i).FdispCateName %>
										<% else %>
											미지정
										<% end if %>
									</td>
            						<td align="right">
										<a href="/admin/brand/shop/collection/index.asp?menupos=1599&standardCateCode=<%= BrandShopCollectionCountArr(i).FdispCateCode %>" target="_blank">
											<b><%= BrandShopCollectionCountArr(i).Fcount %></b> 건
										</a>
									</td>
            					</tr>
								<% end if %>
								<% next %>
            				</table>
            			</td>
            		</tr>
            	</table>
        	    <!--  aaaa -->
        	</td>
        </tr>


        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
				<!--  aaaa -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>당첨이벤트관리</b>
										(<span id="objTimeEventPrize"></span>) <a href="javascript:RefreshData('mdTimeEventPrize')"><img src="/images/icon_reload.gif" border="0"></a>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
										&nbsp;
            						</td>
            					</tr>
								<tr height="25">
            						<td>당첨이벤트목록(당첨자 등록이전)</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<% for i = 0 to UBound(EventPrizeArr) %>
								<% if (mdMainUserID = "") or (mdMainUserID = EventPrizeArr(i).FuserName) or (EventPrizeArr(i).FuserName = "퇴사자") then %>
            					<tr height="25" id="BrandLookBookCount">
            						<td>
										&nbsp;&nbsp;&nbsp; - [<%= EventPrizeArr(i).FuserName %>] <a href="/admin/eventmanage/event/index.asp?menupos=870&selEvt=evt_code&sEtxt=<%= EventPrizeArr(i).FeventCode %>" target="_blank"><%= EventPrizeArr(i).FeventName %></a>
										<% if (EventPrizeArr(i).FeventKind <> 19) and (EventPrizeArr(i).FeventKind <> 25) and (EventPrizeArr(i).FeventKind <> 26) then %>
										<a href="<%=wwwURL%>/event/eventmain.asp?eventid=<%= EventPrizeArr(i).FeventCode %>" target="_blank">[WEB보기]</a>
										<% end if %>
									</td>
            						<td align="right">
										<b><%= EventPrizeArr(i).GetDDayStr %></b>
									</td>
            					</tr>
								<% end if %>
								<% next %>
            				</table>
            			</td>
            		</tr>
            	</table>
        	    <!--  aaaa -->
        	</td>
        </tr>

		</table>
	</td>
	<td width="10"></td>
	<td valign="top">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        <tr valign="top">
            <td>
                <!-- 새로고침 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
						<form name="frm" method="post" action="md_main.asp">
						<input type="hidden" name="menupos" value="<%= menupos %>">
                        <tr height="25">
                        	<td>
            			    	<img src="/images/icon_star.gif" align="absbottom">
								<b>카테고리</b> : <%=fnDispCateSelectBox(1, "", "dispCate", dispCate, "") %>
								&nbsp;
								<b>이름 : </b>
								<input type="text" class="text" name="mdMainUserID" value="<%= mdMainUserID %>" size="10">
								<input type="button" class="button" value="검색" onclick="document.frm.submit();">
            			    </td>
            			    <td align="right">
            			    	<a href="javascript:document.frm.submit();">
        				        새로고침
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
						</form>
            	        </table>
            	    </td>
            	</tr>
            	</table>
            	<!-- 새로고침 끝 -->
            </td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>매출조회</b>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
										&nbsp;
            						</td>
            					</tr>
            					<tr height="25">
            						<td>당일매출(주문일 기준)</td>
            						<td align="right">
										<a href="/admin/maechul/statistic/statistic_category_datamart.asp?menupos=1495&date_gijun=regdate&syear=2014&smonth=9&sday=29&eyear=2014&emonth=9&eday=29&isBanpum=all&categbn=D" target="_blank">
										[당일]
										</a>

										<a href="/admin/maechul/statistic/statistic_category_datamart.asp?menupos=1495&date_gijun=regdate&syear=2014&smonth=9&sday=23&eyear=2014&emonth=9&eday=29&isBanpum=all&categbn=D" target="_blank">
										[주간]
										</a>

										<a href="/admin/maechul/statistic/statistic_category_datamart.asp?menupos=1495&date_gijun=regdate&syear=2014&smonth=8&sday=30&eyear=2014&emonth=9&eday=29&isBanpum=all&categbn=D" target="_blank">
										[월간]
										</a>
									</td>
            					</tr>

								<!--
            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>123,456,789</b>
									</td>
            					</tr>
            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>123,456,789</b>
									</td>
            					</tr>
            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>123,456,789</b>
									</td>
            					</tr>
            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>123,456,789</b>
									</td>
            					</tr>
            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>123,456,789</b>
									</td>
            					</tr>

            					<tr height="25">
            						<td>분기별 수익달성율(출고일 기준)</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>

            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>57%</b>
									</td>
            					</tr>
            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>57%</b>
									</td>
            					</tr>
            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>57%</b>
									</td>
            					</tr>
            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>57%</b>
									</td>
            					</tr>
            					<tr height="25" id="aaaaa">
            						<td>
										&nbsp;&nbsp;&nbsp; -
										디자인문구
									</td>
            						<td align="right">
										<b>57%</b>
									</td>
            					</tr>
								-->

								<!--
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/report/channel_bestseller.asp?sitename=10x10&menupos=302" target="_blank">카테고리별 베스트셀러</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/report/channelupchesellamount.asp?menupos=306" target="_blank">카테고리별 브랜드매출집계</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/maechul/statistic/statistic_category.asp?menupos=1484" target="_blank">카테고리별매출 - 실시간</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/maechul/statistic/statistic_category_datamart.asp?menupos=1495" target="_blank">카테고리별매출 - MART</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>업체별</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/report/upchesellamount.asp?menupos=101" target="_blank">업체별매출집계</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/newreport/newbrandsum.asp?menupos=633" target="_blank">신규업체매출</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								-->
            				</table>
            			</td>
            		</tr>
            	</table>
        	    <!--  aaaa -->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
				<!--  aaaa -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>업무협조</b>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
										&nbsp;
            						</td>
            					</tr>

            					<tr height="25">
            						<td>촬영요청</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="/admin/photo_req/request_list.asp?menupos=1419" target="_blank">촬영요청리스트</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="/admin/photo_req/request_cal.asp?menupos=1420" target="_blank">촬영요청스케줄</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>

            				</table>
            			</td>
            		</tr>
            	</table>
        	    <!--  aaaa -->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>상품소싱</b>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
										&nbsp;
            						</td>
            					</tr>

            					<tr height="25">
            						<td>상품소싱</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="/admin/newstorage/orderlist.asp?menupos=537" target="_blank">주문서관리</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="/admin/stock/brandcurrentstock.asp?menupos=708" target="_blank">브랜드별재고현황</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>

            					<tr height="25">
            						<td>매출조회</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="/admin/ordermaster/oneitembuylist.asp?menupos=77" target="_blank">판매내역[특정상품]</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="/admin/upchejungsan/upcheselllist.asp?menupos=138" target="_blank">판매내역[브랜드]</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>

            				</table>
            			</td>
            		</tr>
            	</table>
			</td>
		</tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
				<!--  aaaa -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>메인페이지관리</b>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
										&nbsp;
            						</td>
            					</tr>
            					<tr height="25">
            						<td>메인페이지</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/categorymaster/category_md_choice.asp?menupos=886" target="_blank">MD`S PICK</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/sitemaster/main_md_recommend_flash.asp?menupos=643" target="_blank">엠디추천상품</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/sitemaster/main_manager.asp?menupos=919" target="_blank">메인페이지관리</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/categorymaster/category_main_pageItem.asp?menupos=949" target="_blank">TODAY`S HOT</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>카테고리메인</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/categorymaster/category_manager.asp?menupos=952" target="_blank">카테고리 페이지 관리</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="/admin/categorymaster/category_main_EventBanner.asp?menupos=967" target="_blank">카테고리 이벤트 배너 관리</a></td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>

            				</table>
            			</td>
            		</tr>
            	</table>
        	    <!--  aaaa -->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
				<!--  aaaa -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>이벤트관리 - 종료임박</b>
										(<span id="objTimeWillFinishEvent"></span>) <a href="javascript:RefreshData('mdTimeWillFinishEvent')"><img src="/images/icon_reload.gif" border="0"></a>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
            							<a href="javascript:popOpenEvent('', '', '', '', '', '')">바로가기 <img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            						</td>
            					</tr>

            					<tr height="25">
            						<td>
										일반 이벤트
										<% if (dispCateName <> "") then %>(<%= dispCateName %>)<% end if %>
									</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(Now(), 10) %>', '<%= Left(DateAdd("d", 0, Now()), 10) %>', '1,12,13,16,17,23,24')">당일종료</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(Now(), 10) %>', '<%= Left(DateAdd("d", 0, Now()), 10) %>', '1,12,13,16,17,23,24')">
											<b><%= willFinishEventArr(0).FNormalCnt %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(DateAdd("d", 1, Now()), 10) %>', '<%= Left(DateAdd("d", 1, Now()), 10) %>', '1,12,13,16,17,23,24')">내일종료</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(DateAdd("d", 1, Now()), 10) %>', '<%= Left(DateAdd("d", 1, Now()), 10) %>', '1,12,13,16,17,23,24')">
											<b><%= willFinishEventArr(1).FNormalCnt %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<% for i = 2 to 6 %>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(DateAdd("d", i, Now()), 10) %>', '<%= Left(DateAdd("d", i, Now()), 10) %>', '1,12,13,16,17,23,24')"><%= Left(DateAdd("d", i, Now()), 10) %> (<%= GetWeekDayName(Left(DateAdd("d", i, Now()), 10)) %>)</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(DateAdd("d", i, Now()), 10) %>', '<%= Left(DateAdd("d", i, Now()), 10) %>', '1,12,13,16,17,23,24')">
											<b><%= willFinishEventArr(i).FNormalCnt %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<% next %>

            					<tr height="25">
            						<td>
										모바일 or 앱 이벤트
										<% if (dispCateName <> "") then %>(<%= dispCateName %>)<% end if %>
									</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(Now(), 10) %>', '<%= Left(DateAdd("d", 0, Now()), 10) %>', '19,25,26')">당일종료</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(Now(), 10) %>', '<%= Left(DateAdd("d", 0, Now()), 10) %>', '19,25,26')">
											<b><%= willFinishEventArr(0).FAppCnt %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(DateAdd("d", 1, Now()), 10) %>', '<%= Left(DateAdd("d", 1, Now()), 10) %>', '19,25,26')">내일종료</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(DateAdd("d", 1, Now()), 10) %>', '<%= Left(DateAdd("d", 1, Now()), 10) %>', '19,25,26')">
											<b><%= willFinishEventArr(1).FAppCnt %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<% for i = 2 to 6 %>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(DateAdd("d", i, Now()), 10) %>', '<%= Left(DateAdd("d", i, Now()), 10) %>', '19,25,26')"><%= Left(DateAdd("d", i, Now()), 10) %> (<%= GetWeekDayName(Left(DateAdd("d", i, Now()), 10)) %>)</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('<%= dispCate %>', '6', 'E', '<%= Left(DateAdd("d", i, Now()), 10) %>', '<%= Left(DateAdd("d", i, Now()), 10) %>', '19,25,26')">
											<b><%= willFinishEventArr(i).FAppCnt %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<% next %>

            				</table>
            			</td>
            		</tr>
            	</table>
        	    <!--  aaaa -->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
				<!--  aaaa -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        			<tr bgcolor="<%= adminColor("menubar") %>">
        				<td>
        					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
								<tr height="25">
            						<td style="border-bottom:1px solid #BABABA">
            			    			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>이벤트관리</b>
																				(<span id="objTimeEventCount"></span>) <a href="javascript:RefreshData('mdTimeEventCount')"><img src="/images/icon_reload.gif" border="0"></a>
            						</td>
            						<td align="right" style="border-bottom:1px solid #BABABA">
            							<a href="javascript:popOpenEvent('', '', '', '', '', '')">바로가기 <img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            						</td>
            					</tr>
            					<tr height="25">
            						<td>진행상태</td>
            						<td align="right">
										&nbsp;
									</td>
            					</tr>
            					<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('', '0', '', '', '', '')">등록대기</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('', '0', '', '', '', '')">
											<b><%= EventCount.Fstate0 %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('', '1', '', '', '', '')">승인반려</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('', '1', '', '', '', '')">
											<b><%= EventCount.Fstate1 %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('', '2', '', '', '', '')">승인요청</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('', '2', '', '', '', '')">
											<b><%= EventCount.Fstate2 %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('', '3', '', '', '', '')">이미지등록요청</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('', '3', '', '', '', '')">
											<b><%= EventCount.Fstate3 %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('', '5', '', '', '', '')">오픈요청(이미지등록완료)</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('', '5', '', '', '', '')">
											<b><%= EventCount.Fstate5 %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('', '7', '', '', '', '')">오픈예정</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('', '7', '', '', '', '')">
											<b><%= EventCount.Fstate7 %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
								<tr height="25">
            						<td>&nbsp; * <a href="javascript:popOpenEvent('', '6', '', '', '', '')">오픈</a></td>
            						<td align="right">
										<a href="javascript:popOpenEvent('', '6', '', '', '', '')">
											<b><%= EventCount.Fstate6 %></b>
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
            					</tr>
            				</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>

		</table>
	</td>
</tr>
</table>

<form name="frmAct" method="post" action="md_main_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="mdTime" value="">
</form>

<% Call checkAndWriteElapsedTime("011") %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
