<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� Ŭ����
' Hieditor : 2011.04.22 �̻� ����
'			 2013.11.21 �ѿ�� ����
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

	if (confirm("!!!! �ִ� 60�ʱ��� �ð��� �ҿ�˴ϴ�. !!!!\n\n���ֹ� ������ ���ۼ��Ͻðڽ��ϱ�?") == true) {
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

	if (confirm("!!!! �ִ� 60�ʱ��� �ð��� �ҿ�˴ϴ�. !!!!\n\n���ֹ� ������ ���ۼ��Ͻðڽ��ϱ�?") == true) {
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

	if (confirm(orderserial+' ���ۼ� �Ͻðڽ��ϱ�?')){
	    frm.orderserial.value = orderserial;
		frm.method.value = "post";
		frm.mode.value = "reOrgorderCSone";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

function jsReloadOrgOrderNCSOneOFF(orderserial){
    var frm = document.frm;

	if (confirm(orderserial+' ���ۼ� �Ͻðڽ��ϱ�?')){
	    frm.orderserial.value = orderserial;
		frm.method.value = "post";
		frm.mode.value = "reOrgorderCSoneOFF";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

function jsReloadOrgOrderNCSOneACA(orderserial){
    var frm = document.frm;

	if (confirm(orderserial+' ���ۼ� �Ͻðڽ��ϱ�?')){
	    frm.orderserial.value = orderserial;
		frm.method.value = "post";
		frm.mode.value = "reOrgorderCSoneACA";
		frm.action = "maechul_log_process.asp";

		frm.submit();
	}
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �˻� ���� -->
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
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���ⱸ�� : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
		&nbsp;&nbsp;
		* �������� : <% drawSelectBoxVatYN "vatinclude", vatinclude %>
		&nbsp;&nbsp;
		* ���Ա��� : <% drawmwdiv_beasongdiv "mwdiv_beasongdiv", mwdiv_beasongdiv , "" %>
		&nbsp;&nbsp;
		* �������� : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
		&nbsp;&nbsp;
		<input type="checkbox" name="excTPL" value="Y" <% if (excTPL = "Y") then %>checked<% end if %> > 3PL ���� ����
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ��¥ :
		<select class="select" name="dategbn">
			<option value="ActDate" <%=CHKIIF(dategbn="ActDate","selected","")%> >��������(ó������)</option>
			<option value="PayDate" <%=CHKIIF(dategbn="PayDate","selected","")%> >����������</option>
		</select>
		<% DrawDateBoxdynamic yyyy1, "yyyy1", yyyy2, "yyyy2", mm1, "mm1", mm2, "mm2", dd1, "dd1", dd2, "dd2" %>
        <input type="checkbox" name="showlevel" value="Y" <%= CHKIIF(showlevel="Y", "checked", "") %>> ȸ����� ǥ��
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p>

	* �ֹ��Ǽ�<br>
	&nbsp; - ���ֹ� : ���ʰ��� �ֹ���<br>
	&nbsp; - ��� : ����ֹ�, �������ȭ�ֹ�<br>
	&nbsp; - ��ǰ : ��ǰ�ֹ�, ��ǰ����ֹ�<br>

<p>


<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50" rowspan="2">���ؿ�</td>
	<td width="120" rowspan="2">����ó</td>
	<td width="70" rowspan="2">ä��</td>
    <td width="70" rowspan="2">ȸ�����</td>
	<td width="120" colspan="3">�ֹ��Ǽ�</td>
	<td width="85" rowspan="2">�Һ��ڰ�<br>�հ�</td>
	<td width="85" rowspan="2">�ǸŰ�<br>(���ΰ�)</td>
	<td width="85" rowspan="2">��ǰ����<br>���밡</td>
	<td width="210" colspan="3">���ʽ�����</td>
	<td width="50" rowspan="2">
		��Ÿ����<br>(�þ�)
	</td>
	<td width="85" rowspan="2">�����Ѿ�</td>
	<td width="65" rowspan="2">���ϸ���</td>
	<td width="65" rowspan="2">����Ʈ</td>
	<td width="65" rowspan="2">��ġ��</td>
	<td width="85" rowspan="2">�ǰ�����</td>
	<td width="85" rowspan="2">��ü<br>�����</td>
	<td width="85" rowspan="2"><b>ȸ�����</b></td>
	<td width="65" rowspan="2">����<br>����<br>����</td>
	<td rowspan="2">���</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">���ֹ�</td>
	<td width="40">���</td>
	<td width="40">��ǰ</td>
	<td width="70">��������</td>
	<td width="70">��������</td>
	<td width="70">��ۺ�<br>����</td>
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
	<td colspan="4">�հ�</td>
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
