<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매출로그
' Hieditor : 2011.04.22 이상구 생성
'###########################################################
%>
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

dim dategbn
dim orgPay_yyyy1, orgPay_yyyy2, orgPay_mm1, orgPay_mm2, orgPay_dd1, orgPay_dd2

dim orgPay_fromDate, orgPay_toDate

dim chkGrpByOrderserial, chkOnlyDiff

dim yyyy, mm, dd, tmpDate
dim searchfield, searchtext

Dim i

dim excTPL

research = requestCheckvar(request("research"),10)
page = requestCheckvar(request("page"),10)

actDivCode = requestCheckvar(request("actDivCode"),10)
targetGbn = requestCheckvar(request("targetGbn"),10)
dategbn     = requestCheckvar(request("dategbn"),10)
orgPay_yyyy1   = request("orgPay_yyyy1")
orgPay_mm1     = request("orgPay_mm1")
orgPay_dd1     = request("orgPay_dd1")
orgPay_yyyy2   = request("orgPay_yyyy2")
orgPay_mm2     = request("orgPay_mm2")
orgPay_dd2     = request("orgPay_dd2")


chkGrpByOrderserial     	= request("chkGrpByOrderserial")
chkOnlyDiff     	= request("chkOnlyDiff")

searchfield 	= request("searchfield")
searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")

excTPL 	= request("excTPL")

if (page="") then page = 1
if (dategbn="") then dategbn="orgPay"

if (research = "") then
	excTPL = "Y"
end if

if (orgPay_yyyy1="") then
	orgPay_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	orgPay_toDate =  DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 2) ''DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	'' orgPay_fromDate = DateSerial(Cstr(Year(now())), 5, 1)
	'' orgPay_toDate = DateSerial(Cstr(Year(now())), 6, 1)

	orgPay_yyyy1 = Cstr(Year(orgPay_fromDate))
	orgPay_mm1 = Cstr(Month(orgPay_fromDate))
	orgPay_dd1 = Cstr(day(orgPay_fromDate))

	tmpDate = DateAdd("d", -1, orgPay_toDate)
	''tmpDate = DateAdd("m", -1, orgPay_toDate)
	orgPay_yyyy2 = Cstr(Year(tmpDate))
	orgPay_mm2 = Cstr(Month(tmpDate))
	orgPay_dd2 = Cstr(day(tmpDate))
else
	orgPay_fromDate = DateSerial(orgPay_yyyy1, orgPay_mm1, orgPay_dd1)
	orgPay_toDate = DateSerial(orgPay_yyyy2, orgPay_mm2, orgPay_dd2+1)
end if



Dim oCMaechulLog
set oCMaechulLog = new CMaechulLog
	oCMaechulLog.FPageSize = 500
	oCMaechulLog.FCurrPage = page

	oCMaechulLog.FRectActDivCode = actDivCode
	oCMaechulLog.FRectDategbn = dategbn

	if (oCMaechulLog.FRectDategbn="ActDate") then
	    oCMaechulLog.FRectActDateStartDate = orgPay_fromDate
	    oCMaechulLog.FRectActDateEndDate = orgPay_toDate
	else
	    oCMaechulLog.FRectOrgPayStartDate = orgPay_fromDate
	    oCMaechulLog.FRectOrgPayEndDate = orgPay_toDate
    end if
	''oCMaechulLog.FRectChkGrpByOrderserial = chkGrpByOrderserial

	oCMaechulLog.FRectSearchField = searchfield
	oCMaechulLog.FRectSearchText = searchtext
    oCMaechulLog.FRectTargetGbn = targetGbn

	if targetGbn = "" and chkOnlyDiff <> "" then
		response.write "<script>alert('먼저 매출구분을 지정하세요.');</script>"
	else
		oCMaechulLog.FRectChkOnlyDiff = chkOnlyDiff
	end if

	oCMaechulLog.FRectExcTPL = excTPL

    ''if (research<>"") then ''최초 검색안함.
	    oCMaechulLog.GetMaechulLogSum
    ''end if

Dim sumTotalMileage, sumAccountSell, summileTotalPrice
dim sumdepositTotalPrice,sumgiftTotalPrice
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
		&nbsp;
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
		&nbsp;
		<select name="dategbn">
		<option value="orgPay" <%=CHKIIF(dategbn="orgPay","selected","")%> >원결제일자
		<option value="ActDate" <%=CHKIIF(dategbn="ActDate","selected","")%> >결제일자(처리일자)
		</select>
		<% DrawDateBoxdynamic orgPay_yyyy1, "orgPay_yyyy1", orgPay_yyyy2, "orgPay_yyyy2", orgPay_mm1, "orgPay_mm1", orgPay_mm2, "orgPay_mm2", orgPay_dd1, "orgPay_dd1", orgPay_dd2, "orgPay_dd2" %>
		&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	    <!--
		&nbsp;
		<input type="checkbox" name="chkGrpByOrderserial" value="Y" <% if (chkGrpByOrderserial = "Y") then %>checked<% end if %> >
		주문번호별합계표시
		-->
		&nbsp;
		<input type="checkbox" name="chkOnlyDiff" value="Y" <% if (chkOnlyDiff = "Y") then %>checked<% end if %> >
		오차내역만 표시(합계표시일 경우, 매출총액)
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<h5>테스트중...</h5>

<p>



<p>

<% if (C_ADMIN_AUTH = True) then %>
[관리자뷰] :
	<%= orgPay_fromDate %> ~ <%= orgPay_yyyy2 %>-<%= Format00(2, orgPay_mm2) %>-<%= Format00(2, orgPay_dd2) %>
	&nbsp;
	&nbsp;
	<% if (DateDiff("d", orgPay_fromDate, orgPay_yyyy2 + "-" + Format00(2, orgPay_mm2) + "-" + Format00(2, orgPay_dd2)) > 3) then %>
	<font color="red">내역 재작성은 기간(원결제일자)이 3일 이내일 경우만 가능합니다.</font>
	<% else %>
		<input type="button" class="button" value="원주문재작성" onClick="jsReloadOrgOrder()">
		<input type="button" class="button" value="CS주문재작성" onClick="jsReloadCSOrder()">
	<% end if %>
<% end if %>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">

		검색결과 : <b><%= oCMaechulLog.FResultCount %></b>
		&nbsp;
	<!--
		페이지 : <b><%= page %> / <%= oCMaechulLog.FTotalPage %></b>
	-->
	<% if oCMaechulLog.FResultCount=oCMaechulLog.FPageSize then %>
	&nbsp;(최대 <%=oCMaechulLog.FPageSize%> 건 표시)
	<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (oCMaechulLog.FRectDategbn="ActDate") then %>
	<td width="70" rowspan="2">결제일<br>(처리일)</td>
	<% else %>
	<td width="70" rowspan="2">원결제일</td>
	<% end if %>
	<% if (C_InspectorUser = False) then %>
	<td width="55" rowspan="2">소비자가<br>합계</td>
	<td width="55" rowspan="2">판매가<br>(할인가)</td>
	<td width="55" rowspan="2">상품쿠폰<br>적용가</td>
	<td width="180" colspan="3">보너스쿠폰</td>
	<td width="50" rowspan="2">
		기타할인<br>(올앳)
	</td>
	<% end if %>
	<td width="100" rowspan="2">매출총액</td>
	<td width="100" rowspan="2">사용 마일리지</td>
	<td width="100" rowspan="2">예치금</td>
	<td width="100" rowspan="2">기프트</td>
	<td width="100" rowspan="2">실결제액</td>
	<td width="100" rowspan="2">업체<br>정산액</td>
	<td width="100" rowspan="2"><b>회계매출</b></td>
	<td width="100" rowspan="2">구매 마일리지</td>
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
<%
summileTotalPrice = summileTotalPrice+oCMaechulLog.FItemList(i).FmileTotalPrice
sumdepositTotalPrice = sumdepositTotalPrice+oCMaechulLog.FItemList(i).FdepositTotalPrice
sumgiftTotalPrice = sumgiftTotalPrice+oCMaechulLog.FItemList(i).FgiftTotalPrice
sumAccountSell  = sumAccountSell + (oCMaechulLog.FItemList(i).FtotalMaechulPrice - oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash)
sumTotalMileage = sumTotalMileage + oCMaechulLog.FItemList(i).FtotalMileage
%>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>

	<td>
		<acronym title="<%= oCMaechulLog.FItemList(i).Fipkumdate %>"><%= Left(oCMaechulLog.FItemList(i).Fipkumdate, 10) %></acronym>
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
	<% if (dategbn="orgPay") and (chkOnlyDiff<>"") then %>
		<% if (oCMaechulLog.FItemList(i).FtotalMaechulPrice-oCMaechulLog.FItemList(i).FrealTotalsum <> 0) then %>
		<font color="red"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalMaechulPrice-oCMaechulLog.FItemList(i).FrealTotalsum, 0) %></font><br>
		<% end if %>
        <% if (oCMaechulLog.FItemList(i).FmileTotalPrice-oCMaechulLog.FItemList(i).FrealSpendmileage <> 0) then %>
		<font color="blue"><%= FormatNumber(oCMaechulLog.FItemList(i).FmileTotalPrice-oCMaechulLog.FItemList(i).FrealSpendmileage, 0) %></font>
		<% end if %>
	<% end if %>
	</td>
</tr>
<% next %>
<tr align="right" bgcolor="FFFFFF">
    <td align="center" >합계</td>
	<% if (C_InspectorUser = False) then %>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
	<% end if %>
    <td></td>
    <td><%= FormatNumber(summileTotalPrice,0) %></td>
    <td><%= FormatNumber(sumdepositTotalPrice,0) %></td>
    <td><%= FormatNumber(sumgiftTotalPrice,0) %></td>
    <td></td>
    <td></td>
    <td><%= FormatNumber(sumAccountSell,0) %></td>
    <td><%= FormatNumber(sumTotalMileage,0) %></td>
    <td></td>
</tr>

<!--
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
-->
</table>

<%
set oCMaechulLog = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
