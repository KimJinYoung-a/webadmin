<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 클래스
' Hieditor : 2011.04.22 이상구 생성
'			 2013.11.21 한용민 수정
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

dim research

dim targetGbn, excTPL, dategbn, showlevel
dim vatinclude, mwdiv_beasongdiv, vPurchasetype

dim yyyy1, mm1, dd1, yyyy2, mm2, dd2

dim yyyy, mm, dd, tmpDate
dim fromDate, toDate

Dim i

research = requestCheckvar(request("research"),10)

targetGbn   = request("targetGbn")
excTPL   = request("excTPL")
showlevel   = request("showlevel")

vatinclude     = requestcheckvar(request("vatinclude"),1)
mwdiv_beasongdiv     = requestcheckvar(request("mwdiv_beasongdiv"),10)
vPurchasetype = request("purchasetype")

dategbn   = request("dategbn")
yyyy1   = request("yyyy1")
mm1   = request("mm1")
dd1   = request("dd1")
yyyy2   = request("yyyy2")
mm2   = request("mm2")
dd2   = request("dd2")

if (research = "") then
	excTPL = "Y"
    ''showlevel = "Y"
	dategbn = "ActDate"
	targetGbn = "ON"
end if

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

Dim oCMaechulLog
set oCMaechulLog = new CMaechulLog
	oCMaechulLog.FPageSize = 500
	oCMaechulLog.FCurrPage = 1

	oCMaechulLog.FRectDategbn = dategbn
	oCMaechulLog.FRectStartDate = fromDate
	oCMaechulLog.FRectEndDate = toDate

	''oCMaechulLog.FRectActDivCode = actDivCode
	''oCMaechulLog.FRectChkGrpByOrderserial = chkGrpByOrderserial
	''oCMaechulLog.FRectChkOnlyDiff = chkOnlyDiff

	''oCMaechulLog.FRectSearchField = searchfield
	''oCMaechulLog.FRectSearchText = searchtext

	oCMaechulLog.FRectTargetGbn = targetGbn

	oCMaechulLog.FRectExcTPL = excTPL
    oCMaechulLog.FRectShowLevel = showlevel

	oCMaechulLog.FRectvatinclude = vatinclude
	oCMaechulLog.FRectmwdiv_beasongdiv = mwdiv_beasongdiv
	oCMaechulLog.FRectPurchasetype = vPurchasetype

	oCMaechulLog.GetMaechulLogByMonth

dim ToTorgOrderCnt, ToTcancelOrderCnt, ToTreturnOrderCnt, ToTorgTotalPrice, ToTsubtotalpriceCouponNotApplied, ToTtotalsum, ToTtotalBonusCouponDiscount, ToTtotalPriceBonusCouponDiscount, ToTtotalBeasongBonusCouponDiscount, ToTallatdiscountprice, ToTtotalMaechulPrice
dim ToTmileTotalPrice, ToTgiftTotalPrice, ToTdepositTotalPrice, ToTGetRealPayPrice, ToTtotalUpcheJungsanCash, ToTtotalMileage

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
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
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 매출구분 : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
		&nbsp;&nbsp;
		* 과세구분 : <% drawSelectBoxVatYN "vatinclude", vatinclude %>
		&nbsp;&nbsp;
		* 매입구분 : <% drawmwdiv_beasongdiv "mwdiv_beasongdiv", mwdiv_beasongdiv , "" %>
		&nbsp;&nbsp;
		* 구매유형 : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
		&nbsp;&nbsp;
		<input type="checkbox" name="excTPL" value="Y" <% if (excTPL = "Y") then %>checked<% end if %> > 3PL 매출 제외
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 날짜 :
		<select class="select" name="dategbn">
			<option value="ActDate" <%=CHKIIF(dategbn="ActDate","selected","")%> >결제일자(처리일자)</option>
			<option value="PayDate" <%=CHKIIF(dategbn="PayDate","selected","")%> >원결제일자</option>
		</select>
		<% DrawDateBoxdynamic yyyy1, "yyyy1", yyyy2, "yyyy2", mm1, "mm1", mm2, "mm2", dd1, "dd1", dd2, "dd2" %>
        <input type="checkbox" name="showlevel" value="Y" <%= CHKIIF(showlevel="Y", "checked", "") %>> 회원등급 표시
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>

	* 주문건수<br>
	&nbsp; - 원주문 : 최초결제 주문만<br>
	&nbsp; - 취소 : 취소주문, 취소정상화주문<br>
	&nbsp; - 반품 : 반품주문, 반품취소주문<br>

<p>


<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50" rowspan="2">기준월</td>
	<td width="120" rowspan="2">매출처</td>
	<td width="70" rowspan="2">채널</td>
    <td width="70" rowspan="2">회원등급</td>
	<td width="120" colspan="3">주문건수</td>
	<td width="85" rowspan="2">소비자가<br>합계</td>
	<td width="85" rowspan="2">판매가<br>(할인가)</td>
	<td width="85" rowspan="2">상품쿠폰<br>적용가</td>
	<td width="210" colspan="3">보너스쿠폰</td>
	<td width="50" rowspan="2">
		기타할인<br>(올앳)
	</td>
	<td width="85" rowspan="2">매출총액</td>
	<td width="65" rowspan="2">마일리지</td>
	<td width="65" rowspan="2">기프트</td>
	<td width="65" rowspan="2">예치금</td>
	<td width="85" rowspan="2">실결제액</td>
	<td width="85" rowspan="2">업체<br>정산액</td>
	<td width="85" rowspan="2"><b>회계매출</b></td>
	<td width="65" rowspan="2">구매<br>마일<br>리지</td>
	<td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">원주문</td>
	<td width="40">취소</td>
	<td width="40">반품</td>
	<td width="70">비율쿠폰</td>
	<td width="70">정액쿠폰</td>
	<td width="70">배송비<br>쿠폰</td>
</tr>

<% for i=0 to oCMaechulLog.FresultCount -1 %>
<%
ToTorgOrderCnt = ToTorgOrderCnt + oCMaechulLog.FItemList(i).ForgOrderCnt
ToTcancelOrderCnt = ToTcancelOrderCnt + oCMaechulLog.FItemList(i).FcancelOrderCnt
ToTreturnOrderCnt = ToTreturnOrderCnt + oCMaechulLog.FItemList(i).FreturnOrderCnt
ToTorgTotalPrice = ToTorgTotalPrice + oCMaechulLog.FItemList(i).ForgTotalPrice
ToTsubtotalpriceCouponNotApplied = ToTsubtotalpriceCouponNotApplied + oCMaechulLog.FItemList(i).FsubtotalpriceCouponNotApplied
ToTtotalsum = ToTtotalsum + oCMaechulLog.FItemList(i).Ftotalsum
ToTtotalBonusCouponDiscount = ToTtotalBonusCouponDiscount + oCMaechulLog.FItemList(i).FtotalBonusCouponDiscount
ToTtotalPriceBonusCouponDiscount = ToTtotalPriceBonusCouponDiscount + oCMaechulLog.FItemList(i).FtotalPriceBonusCouponDiscount
ToTtotalBeasongBonusCouponDiscount = ToTtotalBeasongBonusCouponDiscount + oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount
ToTallatdiscountprice = ToTallatdiscountprice + oCMaechulLog.FItemList(i).Fallatdiscountprice
ToTtotalMaechulPrice = ToTtotalMaechulPrice + oCMaechulLog.FItemList(i).FtotalMaechulPrice
ToTmileTotalPrice = ToTmileTotalPrice + oCMaechulLog.FItemList(i).FmileTotalPrice
ToTgiftTotalPrice = ToTgiftTotalPrice + oCMaechulLog.FItemList(i).FgiftTotalPrice
ToTdepositTotalPrice = ToTdepositTotalPrice + oCMaechulLog.FItemList(i).FdepositTotalPrice
ToTGetRealPayPrice = ToTGetRealPayPrice + oCMaechulLog.FItemList(i).GetRealPayPrice
ToTtotalUpcheJungsanCash = ToTtotalUpcheJungsanCash + oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash
ToTtotalMileage = ToTtotalMileage + oCMaechulLog.FItemList(i).FtotalMileage
%>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCMaechulLog.FItemList(i).Fyyyymm %></td>
	<td><%= oCMaechulLog.FItemList(i).Fsitename %></td>
	<td><%= oCMaechulLog.FItemList(i).GetSellChannelName %></td>
    <td>
        <%= CHKIIF(showlevel="Y", getUserLevelStr(oCMaechulLog.FItemList(i).Fuserlevel), "-") %>
    </td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ForgOrderCnt, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FcancelOrderCnt, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FreturnOrderCnt, 0) %></td>

	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ForgTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FsubtotalpriceCouponNotApplied, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).Ftotalsum, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalBonusCouponDiscount - oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalPriceBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).Fallatdiscountprice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalMaechulPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FmileTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FgiftTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FdepositTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).GetRealPayPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalMaechulPrice - oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalMileage, 0) %></td>
	<td>

	</td>
</tr>
<% next %>
<tr align="center" bgcolor="FFFFFF">
	<td colspan="4">합계</td>
	<td align="right"><%= FormatNumber(ToTorgOrderCnt,0) %></td>
	<td align="right"><%= FormatNumber(ToTcancelOrderCnt,0) %></td>
	<td align="right"><%= FormatNumber(ToTreturnOrderCnt,0) %></td>
	<td align="right"><%= FormatNumber(ToTorgTotalPrice,0) %></td>
	<td align="right"><%= FormatNumber(ToTsubtotalpriceCouponNotApplied,0) %></td>
	<td align="right"><%= FormatNumber(ToTtotalsum,0) %></td>
	<td align="right"><%= FormatNumber((ToTtotalBonusCouponDiscount - ToTtotalBeasongBonusCouponDiscount),0) %></td>
	<td align="right"><%= FormatNumber(ToTtotalPriceBonusCouponDiscount,0) %></td>
	<td align="right"><%= FormatNumber(ToTtotalBeasongBonusCouponDiscount,0) %></td>
	<td align="right"><%= FormatNumber(ToTallatdiscountprice,0) %></td>
	<td align="right"><%= FormatNumber(ToTtotalMaechulPrice,0) %></td>
	<td align="right"><%= FormatNumber(ToTmileTotalPrice,0) %></td>
	<td align="right"><%= FormatNumber(ToTgiftTotalPrice,0) %></td>
	<td align="right"><%= FormatNumber(ToTdepositTotalPrice,0) %></td>
	<td align="right"><%= FormatNumber(ToTGetRealPayPrice,0) %></td>
	<td align="right"><%= FormatNumber(ToTtotalUpcheJungsanCash,0) %></td>
	<td align="right"><%= FormatNumber((ToTtotalMaechulPrice - ToTtotalUpcheJungsanCash),0) %></td>
	<td align="right"><%= FormatNumber(ToTtotalMileage,0) %></td>
	<td></td>
</tr>
</form>
</table>

<%
set oCMaechulLog = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
