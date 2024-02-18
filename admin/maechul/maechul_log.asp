<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� Ŭ����
' Hieditor : 2011.04.22 �̻� ����
'			 2013.11.21 �ѿ�� ����
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

	if (confirm("PG����Ÿ(ON " + pgid + ") �� �������� �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsMatchPGData() {
	var frm = document.frmAct;

	frm.mode.value = "matchpgdata";

	if (confirm("�ڵ���Ī(10x10) �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsMatchFingersPGData() {
	var frm = document.frmAct;

	frm.mode.value = "matchfingerspgdata";

	if (confirm("�ڵ���Ī(�ΰŽ�) �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsMatchGiftCardPGData() {
	var frm = document.frmAct;

	frm.mode.value = "matchgiftcardpgdata";

	if (confirm("�ڵ���Ī(����Ʈ) �Ͻðڽ��ϱ�?") == true) {
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

	if (confirm("[���]���� ��Ī �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsDuplicateMatchCancel(logidx) {
	var frm = document.frmAct;

	frm.logidx.value = logidx;
	frm.mode.value = "matchcancelOnlineDup";

	if (confirm("[���]���� �ߺ����� ��Ī �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}
 */

function jsReloadOrgOrder() {
	var frm = document.frm;

	if (confirm("!!!! �ִ� 60�ʱ��� �ð��� �ҿ�˴ϴ�. !!!!\n\n���ֹ� ������ ���ۼ��Ͻðڽ��ϱ�?") == true) {
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

	if (confirm("!!!! �ִ� 60�ʱ��� �ð��� �ҿ�˴ϴ�. !!!!\n\n���ֹ� ������ ���ۼ��Ͻðڽ��ϱ�?") == true) {
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

	if (confirm("!!!! �ִ� 60�ʱ��� �ð��� �ҿ�˴ϴ�. !!!!\n\n���ֹ� ������ ���ۼ��Ͻðڽ��ϱ�?") == true) {
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

	if (confirm("!!!! �ִ� 60�ʱ��� �ð��� �ҿ�˴ϴ�. !!!!\n\n���ֹ� ������ ���ۼ��Ͻðڽ��ϱ�?") == true) {
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

function popUploadReMakeOrder() {
	var popwin = window.open("popUploadRemakeOrder_on.asp","popUploadReMakeOrder","width=600 height=400 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
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
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���ⱸ�� : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
		&nbsp;&nbsp;
		* �ֹ����� :
		<select class="select" name="actDivCode">
			<option value=""></option>
			<option value="A" <% if (actDivCode = "A") then %>selected<% end if %> >���ֹ�</option>
			<option value="C" <% if (actDivCode = "C") then %>selected<% end if %> >����ֹ�</option>
			<option value="H" <% if (actDivCode = "H") then %>selected<% end if %> >��ǰ����</option>
			<option value="E" <% if (actDivCode = "E") then %>selected<% end if %> >��ȯ�ֹ�</option>
			<option value="M" <% if (actDivCode = "M") then %>selected<% end if %> >��ǰ�ֹ�</option>
			<option value="CC" <% if (actDivCode = "CC") then %>selected<% end if %> >�������ȭ�ֹ�</option>
			<option value="HH" <% if (actDivCode = "HH") then %>selected<% end if %> >��ǰ��������ֹ�</option>
			<option value="EE" <% if (actDivCode = "EE") then %>selected<% end if %> >��ȯ����ֹ�</option>
			<option value="MM" <% if (actDivCode = "MM") then %>selected<% end if %> >��ǰ����ֹ�</option>
		</select>
		&nbsp;&nbsp;
		* �˻����� :
		<select class="select" name="searchfield">
			<option value=""></option>
			<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >�ֹ���ȣ</option>
			<option value="sitename" <% if (searchfield = "sitename") then %>selected<% end if %> >����ó</option>
		</select>
		<input type="text" class="text" name="searchtext" value="<%= searchtext %>">
		&nbsp;
		<input type="checkbox" name="excTPL" value="Y" <% if (excTPL = "Y") then %>checked<% end if %> >
		3PL ���� ����
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="chkOrgPay" value="Y" <% if (chkOrgPay = "Y") then %>checked<% end if %> >
		���������� :
		<% DrawDateBoxdynamic orgPay_yyyy1, "orgPay_yyyy1", orgPay_yyyy2, "orgPay_yyyy2", orgPay_mm1, "orgPay_mm1", orgPay_mm2, "orgPay_mm2", orgPay_dd1, "orgPay_dd1", orgPay_dd2, "orgPay_dd2" %>
		&nbsp;&nbsp;
		<input type="checkbox" name="chkActDate" value="Y" <% if (chkActDate = "Y") then %>checked<% end if %> >
		��������(ó������) :
		<% DrawDateBoxdynamic actDate_yyyy1, "actDate_yyyy1", actDate_yyyy2, "actDate_yyyy2", actDate_mm1, "actDate_mm1", actDate_mm2, "actDate_mm2", actDate_dd1, "actDate_dd1", actDate_dd2, "actDate_dd2" %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
		<input type="checkbox" name="chkGrpByOrderserial" value="Y" <% if (chkGrpByOrderserial = "Y") then %>checked<% end if %> >
		�ֹ���ȣ���հ�ǥ��
		&nbsp;
		<input type="checkbox" name="chkOnlyDiff" value="Y" <% if (chkOnlyDiff = "Y") then %>checked<% end if %> >
		���������� ǥ��(�հ�ǥ���� ���, �����Ѿ�)
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p />

* OK+�ſ� �ֹ����� OKCASHBAG �����αװ� �������� �ʴ� ���, <font color="red">�����α׸� ���� ��</font> �ֹ��α� ���ۼ��ϼ���.

<p />

[�˻��հ�]
<% if (oCMaechulLog.FREsultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
	<td width="110" rowspan="2">�Һ��ڰ�<br>�հ�</td>
	<td width="110" rowspan="2">�ǸŰ�<br>(���ΰ�)</td>
	<td width="110" rowspan="2">��ǰ����<br>���밡</td>
	<td colspan="3">���ʽ�����</td>
	<td width="80" rowspan="2">
		��Ÿ����<br>(�þ�)
	</td>
	<% end if %>
	<td width="110" rowspan="2">�����Ѿ�</td>
	<td width="110" rowspan="2">���ϸ���</td>
	<td width="110" rowspan="2">��ġ��</td>
	<td width="110" rowspan="2">����Ʈ</td>
	<td width="110" rowspan="2">�ǰ�����</td>
	<td width="110" rowspan="2">��ü�����</td>
	<td width="110" rowspan="2"><b>ȸ�����</b></td>
	<td width="110" rowspan="2">���Ÿ��ϸ���</td>
	<td rowspan="2">���</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% if (C_InspectorUser = False) then %>
	<td width="80">��������</td>
	<td width="80">��������</td>
	<td width="80">��ۺ�����</td>
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
		<input type="button" class="button" value="���ֹ����ۼ�(<%= searchtext %>)" onClick="jsReloadOrgOrderOne('<%= searchtext %>')">
		<input type="button" class="button" value="CS�ֹ����ۼ�(<%= searchtext %>)" onClick="jsReloadCSOrderOne('<%= searchtext %>')">
    -->
        <input type="button" class="button" value="ON ���ۼ�(<%= searchtext %>)" onClick="jsReloadOrgOrderNCSOne('<%= searchtext %>')">
		&nbsp;
		<input type="button" class="button" value="OFF ���ۼ�(<%= searchtext %>)" onClick="jsReloadOrgOrderNCSOneOFF('<%= searchtext %>')">
		&nbsp;
		<input type="button" class="button" value="ACA ���ۼ�(<%= searchtext %>)" onClick="jsReloadOrgOrderNCSOneACA('<%= searchtext %>')">
	<% else %>
		&nbsp;
		&nbsp;
		<!-- ������ ���� ���ٰ� Ÿ�Ӿƿ� ����.
		<%= orgPay_fromDate %> ~ <%= orgPay_yyyy2 %>-<%= Format00(2, orgPay_mm2) %>-<%= Format00(2, orgPay_dd2) %>
		<% if (DateDiff("d", orgPay_fromDate, orgPay_yyyy2 + "-" + Format00(2, orgPay_mm2) + "-" + Format00(2, orgPay_dd2)) > 3) then %>
		<font color="red">���� ���ۼ��� �Ⱓ(����������)�� 3�� �̳��� ��츸 �����մϴ�.</font>
		<% else %>
			<input type="button" class="button" value="���ֹ����ۼ�" onClick="jsReloadOrgOrder()">
			<input type="button" class="button" value="CS�ֹ����ۼ�" onClick="jsReloadCSOrder()">
			&nbsp;
			<input type="button" class="button" value="���ֹ����ۼ�(�ΰŽ�)" onClick="jsReloadOrgOrderFingers()">
			<input type="button" class="button" value="CS�ֹ����ۼ�(�ΰŽ�)" onClick="jsReloadCSOrderFingers()">
		<% end if %>
		-->
	<% end if %>

    <input type="button" class="button" value="���ۼ�ť���(ON)" onClick="popUploadReMakeOrder();">

<% end if %>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		�˻���� : <b><%= oCMaechulLog.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCMaechulLog.FTotalPage %></b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80" rowspan="2">����</td>
	<td width="60" rowspan="2">����ó</td>
	<td width="100" rowspan="2">�ֹ���ȣ</td>
	<!--
	<td width="60" rowspan="2">�������</td>
	-->
	<td width="70" rowspan="2">��������</td>
	<td width="70" rowspan="2">������<br>(ó����)</td>
	<% if (C_InspectorUser = False) then %>
	<td width="55" rowspan="2">�Һ��ڰ�<br>�հ�</td>
	<td width="55" rowspan="2">�ǸŰ�<br>(���ΰ�)</td>
	<td width="55" rowspan="2">��ǰ����<br>���밡</td>
	<td width="180" colspan="3">���ʽ�����</td>
	<td width="50" rowspan="2">
		��Ÿ����<br>(�þ�)
	</td>
	<% end if %>
	<td width="60" rowspan="2">�����Ѿ�</td>
	<td width="50" rowspan="2">���ϸ���</td>
	<td width="50" rowspan="2">��ġ��</td>
	<td width="50" rowspan="2">����Ʈ</td>
	<td width="80" rowspan="2">�ǰ�����</td>
	<td width="60" rowspan="2">��ü<br>�����</td>
	<td width="60" rowspan="2"><b>ȸ�����</b></td>
	<td width="40" rowspan="2">����<br>����<br>����</td>
	<td width="70" rowspan="2">�����</td>
	<td rowspan="2">���</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
	<td width="55">��������</td>
	<td width="55">��������</td>
	<td width="55">��ۺ�<br>����</td>
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
