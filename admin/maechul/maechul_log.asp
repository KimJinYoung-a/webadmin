<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 클래스
' Hieditor : 2011.04.22 이상구 생성
'			 2013.11.21 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechulLogCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim research, page

dim actDivCode, targetGbn

dim orgPay_yyyy1, orgPay_yyyy2, orgPay_mm1, orgPay_mm2, orgPay_dd1, orgPay_dd2
dim actDate_yyyy1, actDate_yyyy2, actDate_mm1, actDate_mm2, actDate_dd1, actDate_dd2

dim orgPay_fromDate, orgPay_toDate
dim actDate_fromDate, actDate_toDate

dim chkOrgPay, chkActDate
dim chkGrpByOrderserial, chkOnlyDiff

dim yyyy, mm, dd, tmpDate
dim searchfield, searchtext

dim excTPL

Dim i

research = requestCheckvar(request("research"),10)
page = requestCheckvar(request("page"),10)

actDivCode = requestCheckvar(request("actDivCode"),10)
targetGbn = requestCheckvar(request("targetGbn"),10)

orgPay_yyyy1   = request("orgPay_yyyy1")
orgPay_mm1     = request("orgPay_mm1")
orgPay_dd1     = request("orgPay_dd1")
orgPay_yyyy2   = request("orgPay_yyyy2")
orgPay_mm2     = request("orgPay_mm2")
orgPay_dd2     = request("orgPay_dd2")

actDate_yyyy1   = request("actDate_yyyy1")
actDate_mm1     = request("actDate_mm1")
actDate_dd1     = request("actDate_dd1")
actDate_yyyy2   = request("actDate_yyyy2")
actDate_mm2     = request("actDate_mm2")
actDate_dd2     = request("actDate_dd2")

chkOrgPay     	= request("chkOrgPay")
chkActDate     	= request("chkActDate")
chkGrpByOrderserial     	= request("chkGrpByOrderserial")
chkOnlyDiff     	= request("chkOnlyDiff")

searchfield 	= request("searchfield")
searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")

excTPL 	= request("excTPL")

if (page="") then page = 1
if (chkOrgPay="") and (research = "") then chkOrgPay = "Y"
if (research = "") then
	excTPL = "Y"
end if

if (orgPay_yyyy1="") then
	orgPay_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	orgPay_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 2)

	''orgPay_fromDate = DateSerial(Cstr(Year(now())), 5, 1)
	''orgPay_toDate = DateSerial(Cstr(Year(now())), 6, 1)

	orgPay_yyyy1 = Cstr(Year(orgPay_fromDate))
	orgPay_mm1 = Cstr(Month(orgPay_fromDate))
	orgPay_dd1 = Cstr(day(orgPay_fromDate))

	tmpDate = DateAdd("d", -1, orgPay_toDate)
	orgPay_yyyy2 = Cstr(Year(tmpDate))
	orgPay_mm2 = Cstr(Month(tmpDate))
	orgPay_dd2 = Cstr(day(tmpDate))
else
	orgPay_fromDate = DateSerial(orgPay_yyyy1, orgPay_mm1, orgPay_dd1)
	orgPay_toDate = DateSerial(orgPay_yyyy2, orgPay_mm2, orgPay_dd2+1)
end if

if (actDate_yyyy1="") then
	actDate_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	actDate_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 2)

	'' actDate_fromDate = DateSerial(Cstr(Year(now())), 5, 1)
	'' actDate_toDate = DateSerial(Cstr(Year(now())), 6, 1)

	actDate_yyyy1 = Cstr(Year(actDate_fromDate))
	actDate_mm1 = Cstr(Month(actDate_fromDate))
	actDate_dd1 = Cstr(day(actDate_fromDate))

	tmpDate = DateAdd("d", -1, actDate_toDate)
	actDate_yyyy2 = Cstr(Year(tmpDate))
	actDate_mm2 = Cstr(Month(tmpDate))
	actDate_dd2 = Cstr(day(tmpDate))
else
	actDate_fromDate = DateSerial(actDate_yyyy1, actDate_mm1, actDate_dd1)
	actDate_toDate = DateSerial(actDate_yyyy2, actDate_mm2, actDate_dd2+1)
end if

Dim oCMaechulLog
set oCMaechulLog = new CMaechulLog
	oCMaechulLog.FPageSize = 100
	oCMaechulLog.FCurrPage = page

	if (chkOrgPay = "Y") then
		oCMaechulLog.FRectOrgPayStartDate = orgPay_fromDate
		oCMaechulLog.FRectOrgPayEndDate = orgPay_toDate
	end if

	if (chkActDate = "Y") then
		oCMaechulLog.FRectActDateStartDate = actDate_fromDate
		oCMaechulLog.FRectActDateEndDate = actDate_toDate
	end if

	oCMaechulLog.FRectActDivCode = actDivCode
	oCMaechulLog.FRectChkGrpByOrderserial = chkGrpByOrderserial
	oCMaechulLog.FRectChkOnlyDiff = chkOnlyDiff

	oCMaechulLog.FRectSearchField = searchfield
	oCMaechulLog.FRectSearchText = searchtext

	oCMaechulLog.FRectTargetGbn = targetGbn

	oCMaechulLog.FRectExcTPL = excTPL

	oCMaechulLog.GetMaechulLog

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

/*
function jsGetOnPGData(pgid) {
	var frm = document.frmAct;

	if (pgid == "inicis") {
		frm.mode.value = "getonpgdata";
	} else if (pgid == "uplus") {
		frm.mode.value = "getonpgdatauplus";
	} else {
		alert("ERROR");
		return;
	}

	if (confirm("PG데이타(ON " + pgid + ") 를 가져오기 하시겠습니까?") == true) {
		frm.submit();
	}
}

function jsMatchPGData() {
	var frm = document.frmAct;

	frm.mode.value = "matchpgdata";

	if (confirm("자동매칭(10x10) 하시겠습니까?") == true) {
		frm.submit();
	}
}

function jsMatchFingersPGData() {
	var frm = document.frmAct;

	frm.mode.value = "matchfingerspgdata";

	if (confirm("자동매칭(핑거스) 하시겠습니까?") == true) {
		frm.submit();
	}
}

function jsMatchGiftCardPGData() {
	var frm = document.frmAct;

	frm.mode.value = "matchgiftcardpgdata";

	if (confirm("자동매칭(기프트) 하시겠습니까?") == true) {
		frm.submit();
	}
}

function popUploadKCPPGData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("popRegKCPPGDataFile_on.asp","popUploadKCPPGData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsMatchCancel(logidx) {
	var frm = document.frmAct;

	frm.logidx.value = logidx;
	frm.mode.value = "matchcancelOnline";

	if (confirm("[취소]내역 매칭 하시겠습니까?") == true) {
		frm.submit();
	}
}

function jsDuplicateMatchCancel(logidx) {
	var frm = document.frmAct;

	frm.logidx.value = logidx;
	frm.mode.value = "matchcancelOnlineDup";

	if (confirm("[취소]내역 중복승인 매칭 하시겠습니까?") == true) {
		frm.submit();
	}
}
 */

function jsReloadOrgOrder() {
	var frm = document.frm;

	if (confirm("!!!! 최대 60초까지 시간이 소요됩니다. !!!!\n\n원주문 내역을 재작성하시겠습니까?") == true) {
		frm.startdate.value = "<%= orgPay_fromDate %>";
		frm.enddate.value = "<%= orgPay_yyyy2 %>-<%= Format00(2, orgPay_mm2) %>-<%= Format00(2, orgPay_dd2) %>";
		frm.method.value = "post";
		frm.mode.value = "reorgorder";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

function jsReloadOrgOrderFingers() {
	var frm = document.frm;

	if (confirm("!!!! 최대 60초까지 시간이 소요됩니다. !!!!\n\n원주문 내역을 재작성하시겠습니까?") == true) {
		frm.startdate.value = "<%= orgPay_fromDate %>";
		frm.enddate.value = "<%= orgPay_yyyy2 %>-<%= Format00(2, orgPay_mm2) %>-<%= Format00(2, orgPay_dd2) %>";
		frm.method.value = "post";
		frm.mode.value = "reorgorderfingers";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}


function jsReloadCSOrder() {
	var frm = document.frm;

	if (confirm("!!!! 최대 60초까지 시간이 소요됩니다. !!!!\n\n원주문 내역을 재작성하시겠습니까?") == true) {
		frm.startdate.value = "<%= orgPay_fromDate %>";
		frm.enddate.value = "<%= orgPay_yyyy2 %>-<%= Format00(2, orgPay_mm2) %>-<%= Format00(2, orgPay_dd2) %>";
		frm.method.value = "post";
		frm.mode.value = "recsorder";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

function jsReloadCSOrderFingers() {
	var frm = document.frm;

	if (confirm("!!!! 최대 60초까지 시간이 소요됩니다. !!!!\n\n원주문 내역을 재작성하시겠습니까?") == true) {
		frm.startdate.value = "<%= orgPay_fromDate %>";
		frm.enddate.value = "<%= orgPay_yyyy2 %>-<%= Format00(2, orgPay_mm2) %>-<%= Format00(2, orgPay_dd2) %>";
		frm.method.value = "post";
		frm.mode.value = "recsorderfingers";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

Date.prototype.yyyymmdd = function() {
	var yyyy = this.getFullYear().toString();
	var mm = (this.getMonth()+1).toString(); // getMonth() is zero-based
	var dd  = this.getDate().toString();

	return yyyy + '-' + (mm > 9 ? mm : "0" + mm) + '-' + (dd > 9 ? dd : "0" + dd);
};

function jsReloadOrgOrderOne(orderserial) {
	var frm = document.frm;

	if (confirm("!!!! 최대 60초까지 시간이 소요됩니다. !!!!\n\n원주문 내역을 재작성하시겠습니까?") == true) {
		var nowdate = new Date();

		frm.startdate.value = "2008-01-01";
		frm.enddate.value = nowdate.yyyymmdd();
		frm.orderserial.value = orderserial;
		frm.method.value = "post";
		frm.mode.value = "reorgorderone";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

function jsReloadCSOrderOne(orderserial) {
	var frm = document.frm;

	if (confirm("!!!! 최대 60초까지 시간이 소요됩니다. !!!!\n\n원주문 내역을 재작성하시겠습니까?") == true) {
		var nowdate = new Date();

		frm.startdate.value = "2012-01-01";
		frm.enddate.value = nowdate.yyyymmdd();
		frm.orderserial.value = orderserial;
		frm.method.value = "post";
		frm.mode.value = "recsorderone";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

function jsReloadOrgOrderNCSOne(orderserial){
    var frm = document.frm;

	if (confirm(orderserial+' 재작성 하시겠습니까?')){
	    frm.orderserial.value = orderserial;
		frm.method.value = "post";
		frm.mode.value = "reOrgorderCSone";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

function jsReloadOrgOrderNCSOneOFF(orderserial){
    var frm = document.frm;

	if (confirm(orderserial+' 재작성 하시겠습니까?')){
	    frm.orderserial.value = orderserial;
		frm.method.value = "post";
		frm.mode.value = "reOrgorderCSoneOFF";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

function jsReloadOrgOrderNCSOneACA(orderserial){
    var frm = document.frm;

	if (confirm(orderserial+' 재작성 하시겠습니까?')){
	    frm.orderserial.value = orderserial;
		frm.method.value = "post";
		frm.mode.value = "reOrgorderCSoneACA";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

function popUploadReMakeOrder() {
	var popwin = window.open("popUploadRemakeOrder_on.asp","popUploadReMakeOrder","width=600 height=400 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mode" value="">
<input type="hidden" name="startdate" value="">
<input type="hidden" name="enddate" value="">
<input type="hidden" name="orderserial" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 매출구분 : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
		&nbsp;&nbsp;
		* 주문구분 :
		<select class="select" name="actDivCode">
			<option value=""></option>
			<option value="A" <% if (actDivCode = "A") then %>selected<% end if %> >원주문</option>
			<option value="C" <% if (actDivCode = "C") then %>selected<% end if %> >취소주문</option>
			<option value="H" <% if (actDivCode = "H") then %>selected<% end if %> >상품변경</option>
			<option value="E" <% if (actDivCode = "E") then %>selected<% end if %> >교환주문</option>
			<option value="M" <% if (actDivCode = "M") then %>selected<% end if %> >반품주문</option>
			<option value="CC" <% if (actDivCode = "CC") then %>selected<% end if %> >취소정상화주문</option>
			<option value="HH" <% if (actDivCode = "HH") then %>selected<% end if %> >상품변경취소주문</option>
			<option value="EE" <% if (actDivCode = "EE") then %>selected<% end if %> >교환취소주문</option>
			<option value="MM" <% if (actDivCode = "MM") then %>selected<% end if %> >반품취소주문</option>
		</select>
		&nbsp;&nbsp;
		* 검색조건 :
		<select class="select" name="searchfield">
			<option value=""></option>
			<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >주문번호</option>
			<option value="sitename" <% if (searchfield = "sitename") then %>selected<% end if %> >매출처</option>
		</select>
		<input type="text" class="text" name="searchtext" value="<%= searchtext %>">
		&nbsp;
		<input type="checkbox" name="excTPL" value="Y" <% if (excTPL = "Y") then %>checked<% end if %> >
		3PL 매출 제외
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="chkOrgPay" value="Y" <% if (chkOrgPay = "Y") then %>checked<% end if %> >
		원결제일자 :
		<% DrawDateBoxdynamic orgPay_yyyy1, "orgPay_yyyy1", orgPay_yyyy2, "orgPay_yyyy2", orgPay_mm1, "orgPay_mm1", orgPay_mm2, "orgPay_mm2", orgPay_dd1, "orgPay_dd1", orgPay_dd2, "orgPay_dd2" %>
		&nbsp;&nbsp;
		<input type="checkbox" name="chkActDate" value="Y" <% if (chkActDate = "Y") then %>checked<% end if %> >
		결제일자(처리일자) :
		<% DrawDateBoxdynamic actDate_yyyy1, "actDate_yyyy1", actDate_yyyy2, "actDate_yyyy2", actDate_mm1, "actDate_mm1", actDate_mm2, "actDate_mm2", actDate_dd1, "actDate_dd1", actDate_dd2, "actDate_dd2" %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
		<input type="checkbox" name="chkGrpByOrderserial" value="Y" <% if (chkGrpByOrderserial = "Y") then %>checked<% end if %> >
		주문번호별합계표시
		&nbsp;
		<input type="checkbox" name="chkOnlyDiff" value="Y" <% if (chkOnlyDiff = "Y") then %>checked<% end if %> >
		오차내역만 표시(합계표시일 경우, 매출총액)
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p />

* OK+신용 주문에서 OKCASHBAG 결제로그가 생성되지 않는 경우, <font color="red">결제로그를 삭제 후</font> 주문로그 재작성하세요.

<p />

[검색합계]
<% if (oCMaechulLog.FREsultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
	<td width="110" rowspan="2">소비자가<br>합계</td>
	<td width="110" rowspan="2">판매가<br>(할인가)</td>
	<td width="110" rowspan="2">상품쿠폰<br>적용가</td>
	<td colspan="3">보너스쿠폰</td>
	<td width="80" rowspan="2">
		기타할인<br>(올앳)
	</td>
	<% end if %>
	<td width="110" rowspan="2">매출총액</td>
	<td width="110" rowspan="2">마일리지</td>
	<td width="110" rowspan="2">예치금</td>
	<td width="110" rowspan="2">기프트</td>
	<td width="110" rowspan="2">실결제액</td>
	<td width="110" rowspan="2">업체정산액</td>
	<td width="110" rowspan="2"><b>회계매출</b></td>
	<td width="110" rowspan="2">구매마일리지</td>
	<td rowspan="2">비고</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% if (C_InspectorUser = False) then %>
	<td width="80">비율쿠폰</td>
	<td width="80">정액쿠폰</td>
	<td width="80">배송비쿠폰</td>
<% end if %>
</tr>

<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<% if (C_InspectorUser = False) then %>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.ForgTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FsubtotalpriceCouponNotApplied, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.Ftotalsum, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FOneItem.FtotalBonusCouponDiscount - oCMaechulLog.FOneItem.FtotalPriceBonusCouponDiscount - oCMaechulLog.FOneItem.FtotalBeasongBonusCouponDiscount), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalPriceBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalBeasongBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.Fallatdiscountprice, 0) %></td>
	<% end if %>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalMaechulPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FmileTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FdepositTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FgiftTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.GetRealPayPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalUpcheJungsanCash, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FOneItem.FtotalMaechulPrice - oCMaechulLog.FOneItem.FtotalUpcheJungsanCash), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalMileage, 0) %></td>
	<td></td>
</tr>
</table>
<% end if %>
<p>

<% if True or (C_ADMIN_AUTH = True) then %>
	<% if (searchfield = "orderserial") and (searchtext <> "") then %>
	<!--
		<input type="button" class="button" value="원주문재작성(<%= searchtext %>)" onClick="jsReloadOrgOrderOne('<%= searchtext %>')">
		<input type="button" class="button" value="CS주문재작성(<%= searchtext %>)" onClick="jsReloadCSOrderOne('<%= searchtext %>')">
    -->
        <input type="button" class="button" value="ON 재작성(<%= searchtext %>)" onClick="jsReloadOrgOrderNCSOne('<%= searchtext %>')">
		&nbsp;
		<input type="button" class="button" value="OFF 재작성(<%= searchtext %>)" onClick="jsReloadOrgOrderNCSOneOFF('<%= searchtext %>')">
		&nbsp;
		<input type="button" class="button" value="ACA 재작성(<%= searchtext %>)" onClick="jsReloadOrgOrderNCSOneACA('<%= searchtext %>')">
	<% else %>
		&nbsp;
		&nbsp;
		<!-- 느려서 쿼리 돌다가 타임아웃 난다.
		<%= orgPay_fromDate %> ~ <%= orgPay_yyyy2 %>-<%= Format00(2, orgPay_mm2) %>-<%= Format00(2, orgPay_dd2) %>
		<% if (DateDiff("d", orgPay_fromDate, orgPay_yyyy2 + "-" + Format00(2, orgPay_mm2) + "-" + Format00(2, orgPay_dd2)) > 3) then %>
		<font color="red">내역 재작성은 기간(원결제일자)이 3일 이내일 경우만 가능합니다.</font>
		<% else %>
			<input type="button" class="button" value="원주문재작성" onClick="jsReloadOrgOrder()">
			<input type="button" class="button" value="CS주문재작성" onClick="jsReloadCSOrder()">
			&nbsp;
			<input type="button" class="button" value="원주문재작성(핑거스)" onClick="jsReloadOrgOrderFingers()">
			<input type="button" class="button" value="CS주문재작성(핑거스)" onClick="jsReloadCSOrderFingers()">
		<% end if %>
		-->
	<% end if %>

    <input type="button" class="button" value="재작성큐등록(ON)" onClick="popUploadReMakeOrder();">

<% end if %>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		검색결과 : <b><%= oCMaechulLog.FTotalcount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCMaechulLog.FTotalPage %></b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80" rowspan="2">구분</td>
	<td width="60" rowspan="2">매출처</td>
	<td width="100" rowspan="2">주문번호</td>
	<!--
	<td width="60" rowspan="2">결제방법</td>
	-->
	<td width="70" rowspan="2">원결제일</td>
	<td width="70" rowspan="2">결제일<br>(처리일)</td>
	<% if (C_InspectorUser = False) then %>
	<td width="55" rowspan="2">소비자가<br>합계</td>
	<td width="55" rowspan="2">판매가<br>(할인가)</td>
	<td width="55" rowspan="2">상품쿠폰<br>적용가</td>
	<td width="180" colspan="3">보너스쿠폰</td>
	<td width="50" rowspan="2">
		기타할인<br>(올앳)
	</td>
	<% end if %>
	<td width="60" rowspan="2">매출총액</td>
	<td width="50" rowspan="2">마일리지</td>
	<td width="50" rowspan="2">예치금</td>
	<td width="50" rowspan="2">기프트</td>
	<td width="80" rowspan="2">실결제액</td>
	<td width="60" rowspan="2">업체<br>정산액</td>
	<td width="60" rowspan="2"><b>회계매출</b></td>
	<td width="40" rowspan="2">구매<br>마일<br>리지</td>
	<td width="70" rowspan="2">등록일</td>
	<td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
	<td width="55">비율쿠폰</td>
	<td width="55">정액쿠폰</td>
	<td width="55">배송비<br>쿠폰</td>
	<% end if %>
</tr>

<% for i=0 to oCMaechulLog.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCMaechulLog.FItemList(i).GetActDivCodeName %></td>
	<td><%= oCMaechulLog.FItemList(i).Fsitename %></td>
	<td><% if (chkGrpByOrderserial = "Y") then %><%= oCMaechulLog.FItemList(i).Forderserial %><% else %><%= oCMaechulLog.FItemList(i).GetFullOrderSerial %><% end if %></td>
	<!--
	<td><%= oCMaechulLog.FItemList(i).JumunMethodName %></td>
	-->
	<td>
		<acronym title="<%= oCMaechulLog.FItemList(i).Fipkumdate %>"><%= Left(oCMaechulLog.FItemList(i).Fipkumdate, 10) %></acronym>
	</td>
	<td>
		<% if (chkGrpByOrderserial <> "Y") then %>
		<acronym title="<%= oCMaechulLog.FItemList(i).FactDate %>"><%= Left(oCMaechulLog.FItemList(i).FactDate, 10) %></acronym>
		<% end if %>
	</td>
	<% if (C_InspectorUser = False) then %>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ForgTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FsubtotalpriceCouponNotApplied, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).Ftotalsum, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalBonusCouponDiscount - oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalPriceBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).Fallatdiscountprice, 0) %></td>
	<% end if %>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalMaechulPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FmileTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FdepositTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FgiftTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).GetRealPayPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalMaechulPrice - oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalMileage, 0) %></td>
	<td>
		<acronym title="<%= oCMaechulLog.FItemList(i).Fregdate %>"><%= Left(oCMaechulLog.FItemList(i).Fregdate, 10) %></acronym>
	</td>
	<td>
		<% if (oCMaechulLog.FItemList(i).FrealTotalsum <> 0) then %>
		<font color="red"><%= FormatNumber(oCMaechulLog.FItemList(i).FrealTotalsum, 0) %></font>
		<% end if %>
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="30" align="center">
		<% if oCMaechulLog.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCMaechulLog.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCMaechulLog.StartScrollPage to oCMaechulLog.FScrollCount + oCMaechulLog.StartScrollPage - 1 %>
			<% if i>oCMaechulLog.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCMaechulLog.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<%
set oCMaechulLog = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
