<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� Ŭ����
' Hieditor : 2011.04.22 �̻� ����
'			 2013.11.14 �ѿ�� ����
'###########################################################
Server.ScriptTimeOut = 180
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


dim research, page, pagesize, Dategbn, gubunstartdate, gubunenddate
dim actDivCode, sitename, mwdiv_beasongdiv, makerid
dim orgPay_yyyy1, orgPay_yyyy2, orgPay_mm1, orgPay_mm2, orgPay_dd1, orgPay_dd2, vatinclude, targetGbn
dim actDate_yyyy1, actDate_yyyy2, actDate_mm1, actDate_mm2, actDate_dd1, actDate_dd2
dim chulgoDate_yyyy1, chulgoDate_yyyy2, chulgoDate_mm1, chulgoDate_mm2, chulgoDate_dd1, chulgoDate_dd2
dim jFixedDt_yyyy1, jFixedDt_yyyy2, jFixedDt_mm1, jFixedDt_mm2, jFixedDt_dd1, jFixedDt_dd2, chkjFixedDt
dim orgPay_fromDate, orgPay_toDate
dim actDate_fromDate, actDate_toDate, chulgoDate_fromDate, chulgoDate_toDate, jFixedDt_fromDate, jFixedDt_toDate
dim chkOrgPay, chkActDate, chkChulgoDate, yyyy, mm, dd, tmpDate, searchfield, searchtext, michulgoOnly, miJungsanOnly, i

dim showStatistic, showOnlyStatistic
dim excTPL, excZeroPrice
dim exc6month

	makerid = requestCheckvar(request("makerid"),32)
	research = requestCheckvar(request("research"),10)
	page = requestCheckvar(request("page"),10)
	mwdiv_beasongdiv     = requestcheckvar(request("mwdiv_beasongdiv"),4)
	actDivCode = requestCheckvar(request("actDivCode"),10)
	vatinclude     = requestcheckvar(request("vatinclude"),1)
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
	chulgoDate_yyyy1   = request("chulgoDate_yyyy1")
	chulgoDate_mm1     = request("chulgoDate_mm1")
	chulgoDate_dd1     = request("chulgoDate_dd1")
	chulgoDate_yyyy2   = request("chulgoDate_yyyy2")
	chulgoDate_mm2     = request("chulgoDate_mm2")
	chulgoDate_dd2     = request("chulgoDate_dd2")

	jFixedDt_yyyy1   = request("jFixedDt_yyyy1")
	jFixedDt_mm1     = request("jFixedDt_mm1")
	jFixedDt_dd1     = request("jFixedDt_dd1")
	jFixedDt_yyyy2   = request("jFixedDt_yyyy2")
	jFixedDt_mm2     = request("jFixedDt_mm2")
	jFixedDt_dd2     = request("jFixedDt_dd2")

	targetGbn     = requestcheckvar(request("targetGbn"),16)
	chkOrgPay     	= request("chkOrgPay")
	chkActDate     	= request("chkActDate")
	chkChulgoDate   = request("chkChulgoDate")
	chkjFixedDt		= request("chkjFixedDt")
	searchfield 	= request("searchfield")
	searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")
	michulgoOnly 	= request("michulgoOnly")
    miJungsanOnly 	= request("miJungsanOnly")
	showStatistic 	= request("showStatistic")
	showOnlyStatistic = request("showOnlyStatistic")
	excZeroPrice 	= request("excZeroPrice")
	pagesize 	= request("pagesize")

	excTPL 	= request("excTPL")
	exc6month 	= request("exc6month")

if (page="") then page = 1
if (pagesize="") then pagesize = 20
if (chkOrgPay="") and (chkChulgoDate="") and (chkActDate="") and (chkjFixedDt="") and (research = "") then chkOrgPay = "Y"

if (research = "") then
	excTPL = "Y"
	exc6month = "Y"
end if

if (orgPay_yyyy1="") then
	orgPay_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	orgPay_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 2)

	''orgPay_fromDate = DateSerial(Cstr(Year(now())), 5, 1)
	''orgPay_toDate = DateSerial(Cstr(Year(now())), 5, 2)

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

	''actDate_fromDate = DateSerial(Cstr(Year(now())), 5, 1)
	''actDate_toDate = DateSerial(Cstr(Year(now())), 5, 2)

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

if (chulgoDate_yyyy1="") then
	chulgoDate_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	chulgoDate_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 2)

	''chulgoDate_fromDate = DateSerial(Cstr(Year(now())), 5, 1)
	''chulgoDate_toDate = DateSerial(Cstr(Year(now())), 5, 2)

	chulgoDate_yyyy1 = Cstr(Year(chulgoDate_fromDate))
	chulgoDate_mm1 = Cstr(Month(chulgoDate_fromDate))
	chulgoDate_dd1 = Cstr(day(chulgoDate_fromDate))

	tmpDate = DateAdd("d", -1, chulgoDate_toDate)
	chulgoDate_yyyy2 = Cstr(Year(tmpDate))
	chulgoDate_mm2 = Cstr(Month(tmpDate))
	chulgoDate_dd2 = Cstr(day(tmpDate))
else
	chulgoDate_fromDate = DateSerial(chulgoDate_yyyy1, chulgoDate_mm1, chulgoDate_dd1)
	chulgoDate_toDate = DateSerial(chulgoDate_yyyy2, chulgoDate_mm2, chulgoDate_dd2+1)
end if

if (jFixedDt_yyyy1="") then
	jFixedDt_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	jFixedDt_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 2)

	jFixedDt_yyyy1 = Cstr(Year(jFixedDt_fromDate))
	jFixedDt_mm1 = Cstr(Month(jFixedDt_fromDate))
	jFixedDt_dd1 = Cstr(day(jFixedDt_fromDate))

	tmpDate = DateAdd("d", -1, jFixedDt_toDate)
	jFixedDt_yyyy2 = Cstr(Year(tmpDate))
	jFixedDt_mm2 = Cstr(Month(tmpDate))
	jFixedDt_dd2 = Cstr(day(tmpDate))
else
	jFixedDt_fromDate = DateSerial(jFixedDt_yyyy1, jFixedDt_mm1, jFixedDt_dd1)
	jFixedDt_toDate = DateSerial(jFixedDt_yyyy2, jFixedDt_mm2, jFixedDt_dd2+1)
end if

Dim oCMaechulLog
set oCMaechulLog = new CMaechulLog
	oCMaechulLog.FPageSize = pagesize
	oCMaechulLog.FCurrPage = page

	if (chkOrgPay = "Y") then
		oCMaechulLog.FRectOrgPayStartDate = orgPay_fromDate
		oCMaechulLog.FRectOrgPayEndDate = orgPay_toDate
		oCMaechulLog.FRectDategbn = "orgPay"
		Dategbn="orgPay"
		gubunstartdate = orgPay_fromDate
		gubunenddate = orgPay_toDate
	end if

	if (chkActDate = "Y") then
		oCMaechulLog.FRectActDateStartDate = actDate_fromDate
		oCMaechulLog.FRectActDateEndDate = actDate_toDate
		oCMaechulLog.FRectDategbn = "ActDate"
		Dategbn="ActDate"
		gubunstartdate = actDate_fromDate
		gubunenddate = actDate_toDate
	end if

	if (chkChulgoDate = "Y") then
		oCMaechulLog.FRectChulgoDateStartDate = chulgoDate_fromDate
		oCMaechulLog.FRectChulgoDateEndDate = chulgoDate_toDate
		oCMaechulLog.FRectDategbn = "chulgoDate"
		Dategbn="chulgoDate"
		gubunstartdate = chulgoDate_fromDate
		gubunenddate = chulgoDate_toDate
	end if

	if (chkjFixedDt = "Y") then
		oCMaechulLog.FRectjFixedDtStartDate = jFixedDt_fromDate
		oCMaechulLog.FRectjFixedDtEndDate = jFixedDt_toDate
		oCMaechulLog.FRectDategbn = "jFixedDt"
		Dategbn="jFixedDt"
		gubunstartdate = jFixedDt_fromDate
		gubunenddate = jFixedDt_toDate
	end if

	oCMaechulLog.FRectmwdiv_beasongdiv = mwdiv_beasongdiv
	oCMaechulLog.FRectActDivCode = actDivCode
	oCMaechulLog.FRectvatinclude = vatinclude
	oCMaechulLog.FRectSearchField = searchfield
	oCMaechulLog.FRectSearchText = searchtext
	oCMaechulLog.FRecttargetGbn = targetGbn
	oCMaechulLog.FRectMichulgoOnly = michulgoOnly
    oCMaechulLog.FRectMiJungsanOnly = miJungsanOnly
	oCMaechulLog.FRectmakerid = makerid

	oCMaechulLog.FRectExcTPL = excTPL
	oCMaechulLog.FRectExcZeroPrice = excZeroPrice
	oCMaechulLog.FRectShowStatistic = showStatistic
    oCMaechulLog.FRectshowOnlyStatistic = showOnlyStatistic
	oCMaechulLog.FRectExc6month = exc6month

	oCMaechulLog.GetMaechulDetailLog
	''oCMaechulLog.GetMaechulDetailLogSUM
%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popXL(makerid, startDate, endDate, vatinclude, mwdiv) {
	var showEndDate;

	if (makerid == "") {
		alert("�귣�� �˻� �� �����ٿ� �����մϴ�.");
		return;
	}

	if (vatinclude == "") {
		alert("�������� �˻� �� �����ٿ� �����մϴ�.");
		return;
	}

	if (mwdiv == "") {
		alert("���Ա��� �˻� �� �����ٿ� �����մϴ�.");
		return;
	}

	if ((startDate == "") || (endDate == "")) {
		alert("��¥ �˻� �� �����ٿ� �����մϴ�.");
		return;
	}

	<% if (chulgoDate_toDate <> "") then %>
	showEndDate = '<%= Left(DateAdd("d", -1, chulgoDate_toDate), 10) %>';
	<% end if %>
	if (confirm("�����ޱ�\n\n - �귣�� : " + makerid + "\n - ������� : " + startDate + " ~ " + showEndDate + "\n - �������� : " + vatinclude + "\n - ���Ա��� : " + mwdiv + "\n\n�����Ͻðڽ��ϱ�?")) {
		var popwin = window.open("maechul_detail_log_XL.asp?Dategbn=<%= Dategbn %>&makerid=" + makerid + "&startDate=" + startDate + "&endDate=" + endDate + "&vatinclude=" + vatinclude + "&mwdiv=" + mwdiv,"reActAccMonthSummary","width=1000,height=1000 scrollbars=yes resizable=yes");
		popwin.focus();
	}
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

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="page" value="">
		<input type="hidden" name="research" value="on">
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
				&nbsp;&nbsp;
				* �������� : <% drawSelectBoxVatYN "vatinclude", vatinclude %>
				&nbsp;&nbsp;
				* ���Ա��� : <% drawmwdiv_beasongdiv "mwdiv_beasongdiv", mwdiv_beasongdiv , "" %>
				&nbsp;&nbsp;
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
				&nbsp;&nbsp;
				<input type="checkbox" name="chkChulgoDate" value="Y" <% if (chkChulgoDate = "Y") then %>checked<% end if %> >
				������� :
				<% DrawDateBoxdynamic chulgoDate_yyyy1, "chulgoDate_yyyy1", chulgoDate_yyyy2, "chulgoDate_yyyy2", chulgoDate_mm1, "chulgoDate_mm1", chulgoDate_mm2, "chulgoDate_mm2", chulgoDate_dd1, "chulgoDate_dd1", chulgoDate_dd2, "chulgoDate_dd2" %>
				&nbsp;&nbsp;
				<input type="checkbox" name="chkjFixedDt" value="Y" <% if (chkjFixedDt = "Y") then %>checked<% end if %> >
				����Ȯ������ :
				<% DrawDateBoxdynamic jFixedDt_yyyy1, "jFixedDt_yyyy1", jFixedDt_yyyy2, "jFixedDt_yyyy2", jFixedDt_mm1, "jFixedDt_mm1", jFixedDt_mm2, "jFixedDt_mm2", jFixedDt_dd1, "jFixedDt_dd1", jFixedDt_dd2, "jFixedDt_dd2" %>
			</td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF" >
			<td align="left">
				* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				&nbsp;&nbsp;
				<input type="checkbox" name="michulgoOnly" value="Y" <% if (michulgoOnly = "Y") then %>checked<% end if %> >
				������ֹ���
				&nbsp;&nbsp;
				<input type="checkbox" name="miJungsanOnly" value="Y" <% if (miJungsanOnly = "Y") then %>checked<% end if %> >
				������Ǹ�
				&nbsp;&nbsp;
				<input type="checkbox" name="excZeroPrice" value="Y" <% if (excZeroPrice = "Y") then %>checked<% end if %> >
				�Һ��ڰ� 0�� ��ǰ ����
				&nbsp;&nbsp;
				<input type="checkbox" name="showStatistic" value="Y" <% if (showStatistic = "Y") then %>checked<% end if %> >
				�˻��հ�ǥ��
				&nbsp;&nbsp;

				<input type="checkbox" name="showOnlyStatistic" value="Y" <% if (showOnlyStatistic = "Y") then %>checked<% end if %> >
				�հ�<b>��</b>ǥ��
				&nbsp;&nbsp;

				��ǥ�� :
				<select class="select" name="pagesize">
					<option value="20" <% if (pagesize = "20") then %>selected<% end if %> >20</option>
					<option value="100" <% if (pagesize = "100") then %>selected<% end if %> >100</option>
					<option value="1000" <% if (pagesize = "1000") then %>selected<% end if %> >1000</option>
					<option value="3000" <% if (pagesize = "3000") then %>selected<% end if %> >3000</option>
				</select>
				&nbsp;&nbsp;
				<input type="checkbox" name="exc6month" value="Y" <% if (exc6month = "Y") then %>checked<% end if %> >
				3�������� �ֹ�����
			</td>
		</tr>
	</form>
</table>
<!-- �˻� �� -->

<h5>�׽�Ʈ��...</h5>

<p>
<% if (showStatistic = "Y") then %>
[�˻��հ�]
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
	<td align="right"><%= FormatNumber((oCMaechulLog.FOneItem.FtotalBonusCouponDiscount - oCMaechulLog.FOneItem.FtotalBeasongBonusCouponDiscount), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalPriceBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalBeasongBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.Fallatdiscountprice, 0) %></td>
	<% end if %>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalMaechulPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalUpcheJungsanCash, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FOneItem.FtotalMaechulPrice - oCMaechulLog.FOneItem.FtotalUpcheJungsanCash), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FOneItem.FtotalMileage, 0) %></td>
	<td></td>
</tr>
</table>

<p>
<% end if %>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		�� �ӵ��� ������ ��� ������ ���ð� ��ٷ� �ּ���. ���ϰ� ū ������ �Դϴ�.
	</td>
	<td align="right">
		<input type="button" class="button" value="�����ޱ�(<%= makerid %>)" onclick="popXL('<%= makerid %>', '<%= gubunstartdate %>', '<%= gubunenddate %>', '<%= vatinclude %>', '<%= mwdiv_beasongdiv %>');">
	</td>
</tr>
</table>
<!-- �׼� �� -->

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
	<td width="60" rowspan="2">����</td>
	<td width="60" rowspan="2">����ó</td>
	<td width="100" rowspan="2">�ֹ���ȣ</td>
	<td width="100" rowspan="2">���ֹ���ȣ</td>
	<td width="70" rowspan="2">��������</td>
	<td width="70" rowspan="2">������<br>(ó����)</td>
	<td width="30" rowspan="2">����<br>����</td>
	<td width="30" rowspan="2">��ǰ<br>�ͼ�</td>
	<td width="30" rowspan="2">����<br>����</td>
	<td rowspan="2">�귣��</td>
	<td width="60" rowspan="2">��ǰ�ڵ�</td>
	<td width="60" rowspan="2">�ɼ��ڵ�</td>
	<td rowspan="2">��ǰ��[�ɼǸ�]</td>
	<td width="30" rowspan="2">����</td>
	<% if (C_InspectorUser = False) then %>
	<td width="55" rowspan="2">�Һ��ڰ�<br>�հ�</td>
	<td width="55" rowspan="2">�ǸŰ�<br>(���ΰ�)</td>
	<td width="55" rowspan="2">��ǰ����<br>���밡</td>
	<td colspan="3">���ʽ�����</td>
	<td width="50" rowspan="2">
		��Ÿ����<br>(�þ�)
	</td>
	<% end if %>
	<td width="60" rowspan="2">�����Ѿ�</td>
	<td width="60" rowspan="2">��ü<br>�����</td>
	<td width="60" rowspan="2"><b>ȸ�����</b></td>
	<td width="70" rowspan="2">�����</td>
	<td width="70" rowspan="2">������</td>
	<td width="40" rowspan="2">����<br>����<br>����</td>
	<td width="70" rowspan="2">�����</td>
	<td width="70" rowspan="2">���<br>���԰�</td>
	<td rowspan="2">���</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
	<td width="45">����<br>����</td>
	<td width="45">����<br>����</td>
	<td width="45">��ۺ�<br>����</td>
	<% end if %>
</tr>


<% for i=0 to oCMaechulLog.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCMaechulLog.FItemList(i).GetActDivCodeName %></td>
	<td><%= oCMaechulLog.FItemList(i).Fsitename %></td>
	<td><%= oCMaechulLog.FItemList(i).GetFullOrderSerial %></td>
	<td><%= oCMaechulLog.FItemList(i).Forgorderserial %></td>
	<td>
		<acronym title="<%= oCMaechulLog.FItemList(i).Fipkumdate %>"><%= Left(oCMaechulLog.FItemList(i).Fipkumdate, 10) %></acronym>
	</td>
	<td>
		<acronym title="<%= oCMaechulLog.FItemList(i).FactDate %>"><%= Left(oCMaechulLog.FItemList(i).FactDate, 10) %></acronym>
	</td>
	<td><%= oCMaechulLog.FItemList(i).GetVatIncludeName %></td>
	<td><%= oCMaechulLog.FItemList(i).FtargetGbn %></td>
	<td><%= oCMaechulLog.FItemList(i).GetOMWdivName %></td>
	<td><%= oCMaechulLog.FItemList(i).Fmakerid %></td>
	<td><%= oCMaechulLog.FItemList(i).Fitemid %></td>
	<td><%= oCMaechulLog.FItemList(i).Fitemoption %></td>
	<td align="left">
		<acronym title="<%= oCMaechulLog.FItemList(i).Fitemname %>[<%= oCMaechulLog.FItemList(i).Fitemoptionname %>]"><%= Left(( oCMaechulLog.FItemList(i).Fitemname + "[" +  oCMaechulLog.FItemList(i).Fitemoptionname + "]"), 12) %></acronym>...
	</td>
	<td>
		<% if (Abs(oCMaechulLog.FItemList(i).Fitemno) <> 1) then %>
			<font color="red"><%= oCMaechulLog.FItemList(i).Fitemno %></font>
		<% else %>
			<%= oCMaechulLog.FItemList(i).Fitemno %>
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
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalMaechulPrice - oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash), 0) %></td>
	<td>
		<% if (oCMaechulLog.FItemList(i).Fbeasongdate <> "") then %>
		<acronym title="<%= oCMaechulLog.FItemList(i).Fbeasongdate %>"><%= Left(oCMaechulLog.FItemList(i).Fbeasongdate, 10) %></acronym>
		<% end if %>
	</td>
	<td><%=oCMaechulLog.FItemList(i).FDTLjFixedDt%></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalMileage, 0) %></td>
	<td>
		<acronym title="<%= oCMaechulLog.FItemList(i).Fregdate %>"><%= Left(oCMaechulLog.FItemList(i).Fregdate, 10) %></acronym>
	</td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FavgipgoPrice, 0) %></td>
	<td></td>
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
