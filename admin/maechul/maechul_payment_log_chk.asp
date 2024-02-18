<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� Ŭ����
' Hieditor : 2011.04.22 �̻� ����
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
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim yyyy, mm, dd
dim fromDate ,toDate, tmpDate
dim targetGbn

Dim i

research = requestCheckvar(request("research"),10)
page = requestCheckvar(request("page"),10)

yyyy1   = request("yyyy1")
mm1     = request("mm1")
dd1     = request("dd1")
yyyy2   = request("yyyy2")
mm2     = request("mm2")
dd2     = request("dd2")

targetGbn		= requestCheckvar(request("targetGbn"),10)

if (page="") then page = 1
if (targetGbn = "") then
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

Dim oCMaechulPaymentLog
set oCMaechulPaymentLog = new CMaechulLog
	oCMaechulPaymentLog.FPageSize = 100
	oCMaechulPaymentLog.FCurrPage = page

	oCMaechulPaymentLog.FRectStartdate = fromDate
	oCMaechulPaymentLog.FRectEndDate = toDate

	oCMaechulPaymentLog.FRectTargetGbn = targetGbn

	if (DateDiff("d", fromDate, toDate) = 1) then
		oCMaechulPaymentLog.FRectChkGrpByOrderserial = "Y"
	end if

	oCMaechulPaymentLog.GetMaechulPaymentLogCheck

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function jsSetDate(yyyy, mm, dd) {
	var frm = document.frm;

	frm.yyyy1.value = yyyy;
	frm.mm1.value = mm;
	frm.dd1.value = dd;

	frm.yyyy2.value = yyyy;
	frm.mm2.value = mm;
	frm.dd2.value = dd;

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

function jsDelMatch(orderserial, suborderserial, paydivcode) {
	var frm = document.frmAct;

	frm.mode.value = "delmatch";

	frm.orderserial.value = orderserial;
	frm.suborderserial.value = suborderserial;
	frm.paydivcode.value = paydivcode;

	if (confirm("������Ī ������ ���� �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsNormalizeMatch(orderserial, suborderserial, orgorderserial, payreqprice) {
	var frm = document.frmAct;

	frm.mode.value = "normalizematch";

	frm.orderserial.value = orderserial;
	frm.suborderserial.value = suborderserial;
	frm.orgorderserial.value = orgorderserial;
	frm.payreqprice.value = payreqprice;

	if (confirm("����ֹ�-����ȭ�ֹ� ��Ī �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsRefundProcMatch(orderserial, suborderserial) {
	var frm = document.frmAct;

	frm.mode.value = "matchRefundProc";

	frm.orderserial.value = orderserial;
	frm.suborderserial.value = suborderserial;

	if (confirm("ȯ�������� ���� �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsAddPayLog(orderserial, suborderserial) {
	var frm = document.frmAct;

	frm.mode.value = "addpaylog";

	frm.orderserial.value = orderserial;
	frm.suborderserial.value = suborderserial;

	if (confirm("�����α� �߰����� �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsPopSearchRefundCS(targetGbn, orderserial, suborderserial, orgorderserial, chgorderserial, reqDate, reqPrice) {
    var window_width = 1200;
    var window_height = 500;

    var popwin = window.open("popSearchRefundCS.asp?targetGbn=" + targetGbn + "&orderserial=" + orderserial + "&suborderserial=" + suborderserial + "&orgorderserial=" + orgorderserial + "&chgorderserial=" + chgorderserial + "&reqDate=" + reqDate + "&reqPrice=" + reqPrice,"jsPopSearchRefundCS","width=" + window_width + " height=" + window_height + " scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}
 */

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;
		���ⱸ�� : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
		&nbsp;
		�ֹ�����(ó������) :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p>

	<font color="red">������ ����!!</font><br>

	* ó�����ڸ� �Ϸ�� �����ϸ� �ֹ���ȣ�� ǥ�õ˴ϴ�.<br />
    * �ֹ������ = ����α�, ������û�� = �����α�

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oCMaechulPaymentLog.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCMaechulPaymentLog.FTotalPage %></b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">����</td>
	<td width="80">������<br>(ó����)</td>
	<td width="100">�ֹ���ȣ</td>
	<td width="100">�ֹ������</td>
	<td width="80">������û��</td>
	<td width="80">����</td>

	<td>���</td>
</tr>

<% for i=0 to oCMaechulPaymentLog.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCMaechulPaymentLog.FItemList(i).FtargetGbn %></td>
	<td>
		<a href="javascript:jsSetDate('<%= Left(oCMaechulPaymentLog.FItemList(i).Factdate, 4) %>', '<%= Right(Left(oCMaechulPaymentLog.FItemList(i).Factdate, 7), 2) %>', '<%= Right(oCMaechulPaymentLog.FItemList(i).Factdate, 2) %>')">
			<%= oCMaechulPaymentLog.FItemList(i).Factdate %>
		</a>
	</td>
	<td><%= oCMaechulPaymentLog.FItemList(i).Forderserial %></td>
	<td align="right"><%= FormatNumber(oCMaechulPaymentLog.FItemList(i).FtotalOrderMaechulPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulPaymentLog.FItemList(i).FtotalpayreqPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulPaymentLog.FItemList(i).FtotalOrderMaechulPrice - oCMaechulPaymentLog.FItemList(i).FtotalpayreqPrice, 0) %></td>

	<td>

	</td>
</tr>
<% next %>

</form>
</table>

<%
set oCMaechulPaymentLog = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
