<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����α�(�ֹ���)
' Hieditor : 2011.04.22 �̻� ����
'			 2020.07.24 �ѿ�� ����(�����α׸�Ī�߰�)
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
<!-- #include virtual="/lib/classes/maechul/incMaechulFunction.asp"-->
<%
dim research, page, yyyy1,yyyy2,mm1,mm2,dd1,dd2, yyyy3, mm3, dd3, yyyy4, mm4, yyyy, mm, dd, fromDate ,toDate, tmpDate
dim searchfield, searchtext, payDivCode, PGgubun, PGuserid, targetGbn, dateGubun, currCSOrderserial, showDelMatchBtn
dim showOnlyPriceNotMatch, matchState, excNoPay, excNoReqPay, excHP, excGift, i, oCMaechulPaymentLog
dim pagesize, incActPayMonthDiff
	research = requestCheckvar(request("research"),10)
	page = requestCheckvar(request("page"),10)
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	yyyy3   = request("yyyy3")
	mm3     = request("mm3")
	dd3     = request("dd3")
	yyyy4   = request("yyyy4")
	mm4     = request("mm4")
	searchfield 	= request("searchfield")
	searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")
	payDivCode     	= request("payDivCode")
	PGgubun     	= request("PGgubun")
	PGuserid     	= request("PGuserid")
	matchState     	= request("matchState")
	targetGbn		= requestCheckvar(request("targetGbn"),10)
	showOnlyPriceNotMatch     	= request("showOnlyPriceNotMatch")
	excNoPay     				= request("excNoPay")
	excNoReqPay     			= request("excNoReqPay")
	excHP     					= request("excHP")
	excGift    					= request("excGift")
	dateGubun     				= request("dateGubun")
    pagesize     				= request("pagesize")
    incActPayMonthDiff			= request("incActPayMonthDiff")


if (page="") then page = 1
if (pagesize="") then pagesize = 100

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 2) ''DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)  ''�Ϸ�ġ�� ����

	''fromDate = DateSerial(Cstr(Year(now())), 4, 1)
	''toDate = DateSerial(Cstr(Year(now())), 5, 1)

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

if (yyyy3 = "") then
	yyyy3 = yyyy1
	mm3 = mm1
	dd3 = dd1
end if

if (yyyy4 = "") then
	yyyy4 = yyyy1
	mm4 = mm1
end if

''���� 1��
tmpDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
yyyy = Cstr(Year(tmpDate))
mm = Format00(2, Cstr(Month(tmpDate)))
dd = Format00(2, Cstr(day(tmpDate)))

set oCMaechulPaymentLog = new CMaechulLog
	oCMaechulPaymentLog.FPageSize = pagesize
	oCMaechulPaymentLog.FCurrPage = page
	oCMaechulPaymentLog.FRectStartdate = fromDate
	oCMaechulPaymentLog.FRectEndDate = toDate
	oCMaechulPaymentLog.FRectSearchField = searchfield
	oCMaechulPaymentLog.FRectSearchText = searchtext
	oCMaechulPaymentLog.FRectDateGubun = dateGubun

	if (searchfield <> "orderserial") or (searchtext = "") then
		'// �˻����� �ִ°�� �Ʒ� ���� ����

		oCMaechulPaymentLog.FRectTargetGbn = targetGbn
		oCMaechulPaymentLog.FRectPayDivCode = payDivCode
		oCMaechulPaymentLog.FRectPGgubun = PGgubun
		oCMaechulPaymentLog.FRectPGuserid = PGuserid
		oCMaechulPaymentLog.FRectMatchState = matchState
		oCMaechulPaymentLog.FRectShowOnlyPriceNotMatch = showOnlyPriceNotMatch
		oCMaechulPaymentLog.FRectExcNoPay = excNoPay
		oCMaechulPaymentLog.FRectExcNoReqPay = excNoReqPay
		oCMaechulPaymentLog.FRectExcHP = excHP
		oCMaechulPaymentLog.FRectExcGift = excGift
        oCMaechulPaymentLog.FRectIncActPayMonthDiff = incActPayMonthDiff
	end if

	oCMaechulPaymentLog.GetMaechulPaymentLog

if (searchfield = "orderserial") and (searchtext <> "") then
	showDelMatchBtn = "Y"
end if

%>

<script type='text/javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

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

function jsMatchCancel(orderserial, suborderserial) {
	var frm = document.frmAct;

	frm.orderserial.value = orderserial;
	frm.suborderserial.value = suborderserial;
	frm.mode.value = "matchCancel";

	if (confirm("[���]���� ��Ī �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsMatchReturn(orderserial, suborderserial) {
	var frm = document.frmAct;

	frm.orderserial.value = orderserial;
	frm.suborderserial.value = suborderserial;
	frm.mode.value = "matchReturn";

	if (confirm("[��ǰ]���� ��Ҹ�Ī �Ͻðڽ��ϱ�?") == true) {
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

// �����α� ��Ī
function jsMatchPay() {
	var frm = document.frmAct;
	var frmDate = document.frmDate;

	if (frmDate.yyyy3.value=='' || frmDate.mm3.value=='' || frmDate.dd3.value==''){
		alert('��Ī�Ͻ� �������� �����ϼ���.');
		return;
	}

	var base_yyyymmdd = "<%= yyyy %>-<%= Format00(2,mm) %>-<%= Format00(2,dd) %>";
	var curr_yyyymmdd = frmDate.yyyy3.value + "-" + frmDate.mm3.value + "-" + frmDate.dd3.value;

	if (curr_yyyymmdd.length != 10) {
		alert("��¥ ������ 0000-00-00 �̾�� �մϴ�.");
		return;
	}

	if (curr_yyyymmdd < base_yyyymmdd) {
		alert("��¥������ " + base_yyyymmdd + " ���ĸ� �����մϴ�.");
		return;
	}

	frm.mode.value = "matchByDay";
	frm.yyyymmdd.value = curr_yyyymmdd;

	if (confirm("�����α� ��Ī �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

// �����α� ��Ī 5�Ͼ�
function jsMatchPaydaypart(daypart) {
	if (frmDate.yyyy4.value=='' || frmDate.mm4.value==''){
		alert('��Ī�Ͻ� �������� �����ϼ���.');
		return;
	}
	var curr_yyyymmdd = frmDate.yyyy4.value + "-" + frmDate.mm4.value
	if (curr_yyyymmdd.length != 7) {
		alert("��¥ ������ 0000-00 �̾�� �մϴ�.");
		return;
	}

	if (daypart==''){
		alert('�����ڰ� �����ϴ�.');
		return;
	}
	frmAct.daypart.value=daypart;

	if (confirm("�����α� ��Ī �Ͻðڽ��ϱ�?") == true) {
		frmAct.mode.value = "matchByDaydaypart";
		frmAct.yyyymmdd.value = curr_yyyymmdd;
		frmAct.submit();
	}

}

function jsPopSearchRefundCS(targetGbn, orderserial, suborderserial, orgorderserial, chgorderserial, reqDate, reqPrice) {
    var window_width = 1200;
    var window_height = 500;

    var popwin = window.open("popSearchRefundCS.asp?targetGbn=" + targetGbn + "&orderserial=" + orderserial + "&suborderserial=" + suborderserial + "&orgorderserial=" + orgorderserial + "&chgorderserial=" + chgorderserial + "&reqDate=" + reqDate + "&reqPrice=" + reqPrice,"jsPopSearchRefundCS","width=" + window_width + " height=" + window_height + " scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsSetRefunding() {

    var frm = document.frmAct;
    var dateGubun = '<%= dateGubun %>';
    // dateGubun

    if ((dateGubun != 'payreqdate') && (dateGubun != 'paydate')) {
        alert('��¥ ������ ������(ó���� �Ǵ� ������) �϶��� ��밡���մϴ�.');
        return;
    }

	if (confirm("ȯ������ �ϰ����� �Ͻðڽ��ϱ�?") == true) {
		frmAct.mode.value = "setrefunding";
        frmAct.dateGubun.value = dateGubun;
		frmAct.startDt.value = '<%= Left(fromDate, 10) %>';
        frmAct.endDt.value = '<%= Left(toDate, 10) %>';
		frmAct.submit();
	}
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;
		���ⱸ�� : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
		&nbsp;
		<select name="dateGubun"  class="select">
			<option value="payreqdate" <% if dateGubun = "payreqdate" then %>selected<% end if %> >������(ó����)</option>
			<option value="paydate" <% if dateGubun = "paydate" then %>selected<% end if %> >������(������)</option>
			<option value="maeipdate" <% if dateGubun = "maeipdate" then %>selected<% end if %> >ī��������</option>
			<option value="mayipkumdate" <% if dateGubun = "mayipkumdate" then %>selected<% end if %> >�Աݿ�����</option>
		</select> :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		&nbsp;
		* �˻����� :
		<select class="select" name="searchfield">
		<option value=""></option>
		<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >�ֹ���ȣ</option>
		<option value="pggubun" <% if (searchfield = "pggubun") then %>selected<% end if %> >PG��</option>
		<option value="pguserid" <% if (searchfield = "pguserid") then %>selected<% end if %> >PG�� ID</option>
		<option value="pgkey" <% if (searchfield = "pgkey") then %>selected<% end if %> >PGkey</option>
		<option value="pgcskey" <% if (searchfield = "pgcskey") then %>selected<% end if %> >PGCSkey</option>
		<option value="payReqPrice" <% if (searchfield = "payReqPrice") then %>selected<% end if %> >��û��</option>
		<option value="realPayPrice" <% if (searchfield = "realPayPrice") then %>selected<% end if %> >���ξ�</option>
		</select>
		<input type="text" class="text" name="searchtext" value="<%= searchtext %>">
		&nbsp;
		* ��Ī���� :
		<select class="select" name="matchState">
		<option value=""></option>
		<option value="Y" <% if (matchState = "Y") then %>selected<% end if %> >��ü ��Ī �Ϸ�</option>
		<option value="A" <% if (matchState = "A") then %>selected<% end if %> >�ڵ���Ī �Ϸ�</option>
		<option value="H" <% if (matchState = "H") then %>selected<% end if %> >�����Ī �Ϸ�</option>
		<option value="R" <% if (matchState = "R") then %>selected<% end if %> >ȯ�� ������</option>
		<option value="X" <% if (matchState = "X") then %>selected<% end if %> >��Ī����</option>
		</select>
        * ǥ�ð��� :
		<select class="select" name="pagesize">
			<option value="100">100</option>
			<option value="500" <%= CHKIIF(pagesize="500", "selected", "")%> >500</option>
			<option value="1000" <%= CHKIIF(pagesize="1000", "selected", "")%> >1000</option>
			<option value="2500" <%= CHKIIF(pagesize="2500", "selected", "")%> >2500</option>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
		������� :
		<select class="select" name="payDivCode">
			<option value=""></option>
			<option value="7" <% if (payDivCode = "7") then %>selected<% end if %> >������(����)</option>
			<option value="100" <% if (payDivCode = "100") then %>selected<% end if %> >�ſ�</option>
			<option value="20" <% if (payDivCode = "20") then %>selected<% end if %> >�ǽð�</option>
			<option value="50" <% if (payDivCode = "50") then %>selected<% end if %> >������</option>
			<option value="80" <% if (payDivCode = "80") then %>selected<% end if %> >All@</option>
			<option value="90" <% if (payDivCode = "90") then %>selected<% end if %> >��ǰ��</option>
			<option value="110" <% if (payDivCode = "110") then %>selected<% end if %> >OK+�ſ�</option>
			<option value="400" <% if (payDivCode = "400") then %>selected<% end if %> >�ڵ���</option>
			<option value="550" <% if (payDivCode = "550") then %>selected<% end if %> >������</option>
			<option value="560" <% if (payDivCode = "560") then %>selected<% end if %> >����Ƽ��</option>
			<option value="">------------</option>
			<option value="mil" <% if (payDivCode = "mil") then %>selected<% end if %> >���ϸ���</option>
			<option value="dep" <% if (payDivCode = "dep") then %>selected<% end if %> >��ġ��</option>
			<option value="gif" <% if (payDivCode = "gif") then %>selected<% end if %> >����Ʈī��</option>
			<option value="">------------</option>
			<option value="77" <% if (payDivCode = "77") then %>selected<% end if %> >������ȯ��</option>
			<option value="6" <% if (payDivCode = "6") then %>selected<% end if %> >�������Ա�</option>
			<option value="rmi" <% if (payDivCode = "rmi") then %>selected<% end if %> >���ϸ���ȯ��</option>
			<option value="rde" <% if (payDivCode = "rde") then %>selected<% end if %> >��ġ��ȯ��</option>
			<option value="">------------</option>
			<option value="0" <% if (payDivCode = "0") then %>selected<% end if %> >��������</option>
			<option value="XXX" <% if (payDivCode = "XXX") then %>selected<% end if %> >XXX</option>
		</select>
		&nbsp;
		PG�� :
		<select name="PGgubun" class="select">
			<option value="">--����--</option>
			<%Call sbGetOptPGgubun(PGgubun)%>
		</select>
		&nbsp;
		PG�� ID : 
		<select name="PGuserid" class="select">
			<option value="">--����--</option>
			<%Call sbGetOptPGID(PGuserid)%>
		</select>
		<% 'Call DrawSelectBoxPGUserid("PGuserid", PGuserid, "") %>
		&nbsp;
		<input type="checkbox" name="showOnlyPriceNotMatch" value="Y" <% if (showOnlyPriceNotMatch = "Y") then  %>checked<% end if %> >
		�ݾ�(��û, ����) ����ġ ������
		&nbsp;
		<input type="checkbox" name="excNoPay" value="Y" <% if (excNoPay = "Y") then  %>checked<% end if %> >
		�������� ���� ����
		&nbsp;
		<input type="checkbox" name="excNoReqPay" value="Y" <% if (excNoReqPay = "Y") then  %>checked<% end if %> disabled>
		������û�� 0������ ����
		&nbsp;
		<input type="checkbox" name="excHP" value="Y" <% if (excHP = "Y") then  %>checked<% end if %> >
		�ڵ��� ������ ����
		&nbsp;
		<input type="checkbox" name="excGift" value="Y" <% if (excGift = "Y") then  %>checked<% end if %> >
		������,����Ƽ�� ������ ����
        &nbsp;
		<input type="checkbox" name="incActPayMonthDiff" value="Y" <% if (incActPayMonthDiff = "Y") then  %>checked<% end if %> >
		���ο�-ó���� ����ġ��
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>

<!-- �׼� ���� -->
<form name="frmDate" method="get" action="" style="margin:0px;">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">������ ����!!</font>
	</td>
	<td align="right">
	</td>
</tr>
<tr>
	<td align="left">
		������(ó����) :
		<br>
		1�ϴ��� : <% Call DrawOneDateBoxdynamic("yyyy3", yyyy3, "mm3", Format00(2,mm3), "dd3", Format00(2,dd3), "", "", "", "") %>
		<input type="button" class="button" value="��Ī" onClick="jsMatchPay();" >
		<br>
		5�ϴ��� :
		<% call DrawYMBoxdynamic("yyyy4", yyyy4, "mm4", mm4, "") %>
		<input type="button" class="button" value="��Ī(1��~5��)" onClick="jsMatchPaydaypart('1');" >
		<input type="button" class="button" value="��Ī(6��~10��)" onClick="jsMatchPaydaypart('2');" >
		<input type="button" class="button" value="��Ī(11��~15��)" onClick="jsMatchPaydaypart('3');" >
		<input type="button" class="button" value="��Ī(16��~20��)" onClick="jsMatchPaydaypart('4');" >
		<input type="button" class="button" value="��Ī(21��~25��)" onClick="jsMatchPaydaypart('5');" >
		<input type="button" class="button" value="��Ī(26��~����)" onClick="jsMatchPaydaypart('6');" >
	</td>
	<td align="right">
        <input type="button" value="ȯ������ �ϰ�����" onClick="jsSetRefunding()">
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oCMaechulPaymentLog.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCMaechulPaymentLog.FTotalPage %></b>
        &nbsp;
        ��û�Ѿ� : <b><%= FormatNumber(oCMaechulPaymentLog.FTotalPayReqPrice, 0) %></b>
        &nbsp;
        �����Ѿ� : <b><%= FormatNumber(oCMaechulPaymentLog.FTotalRealPayPrice, 0) %></b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">����</td>
	<td width="100">�ֹ���ȣ</td>
	<td width="100">���ֹ���ȣ</td>
	<td width="80">������û��</td>
	<td width="80">������<br>(ó����)</td>

	<td width="80">�������</td>
	<td width="100">PG��</td>
	<td width="100">PG�� ID</td>
	<td>PG�� Key</td>
	<td>PG�� CSKey</td>
	<td width="80">�ǽ��ξ�</td>
	<td width="80">������</td>
	<td width="80">�Աݿ�����</td>
	<td width="80">������<br>(������)</td>
	<td width="80">ī��������</td>
	<td width="80">�Աݿ�����</td>

	<td width="80">����</td>
	<td width="80">��Ī����</td>

	<td width="80">�����</td>
	<td>���</td>
</tr>

<% for i=0 to oCMaechulPaymentLog.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCMaechulPaymentLog.FItemList(i).GetActDivCodeName %></td>
	<td><%= oCMaechulPaymentLog.FItemList(i).GetFullOrderSerial %></td>
	<td><%= oCMaechulPaymentLog.FItemList(i).Forgorderserial %></td>
	<td align="right"><%= FormatNumber(oCMaechulPaymentLog.FItemList(i).FpayReqPrice, 0) %></td>
	<td>
		<acronym title="<%= oCMaechulPaymentLog.FItemList(i).FpayReqDate %>"><%= Left(oCMaechulPaymentLog.FItemList(i).FpayReqDate, 10) %></acronym>
	</td>

	<td><%= oCMaechulPaymentLog.FItemList(i).GetPayDivCodeName %></td>
	<td><%= oCMaechulPaymentLog.FItemList(i).FPGgubun %></td>
	<td><%= oCMaechulPaymentLog.FItemList(i).FPGuserid %></td>
	<td><%= oCMaechulPaymentLog.FItemList(i).FPGkey %></td>
	<td><%= oCMaechulPaymentLog.FItemList(i).FPGCSkey %></td>
	<td align="right">
		<% if (oCMaechulPaymentLog.FItemList(i).FrealPayPrice <> "") then %>
			<%= FormatNumber(oCMaechulPaymentLog.FItemList(i).FrealPayPrice, 0) %>
		<% end if %>
	</td>
	<td align="right">
	    <% if (oCMaechulPaymentLog.FItemList(i).FcommPrice <> "") then %>
	        <%= FormatNumber(-1 * oCMaechulPaymentLog.FItemList(i).FcommPrice + -1 * oCMaechulPaymentLog.FItemList(i).FcommVatPrice, 0) %>
	    <% end if %>
	</td>
	<td align="right">
	    <% if (oCMaechulPaymentLog.FItemList(i).FjungsanPrice <> "") then %>
	        <%= FormatNumber(oCMaechulPaymentLog.FItemList(i).FjungsanPrice, 0) %>
	    <% end if %>

	</td>
	<td>
		<acronym title="<%= oCMaechulPaymentLog.FItemList(i).FpayDate %>"><%= Left(oCMaechulPaymentLog.FItemList(i).FpayDate, 10) %></acronym>
	</td>
	<td>
		<acronym title="<%= oCMaechulPaymentLog.FItemList(i).FmaeipDate %>"><%= Left(oCMaechulPaymentLog.FItemList(i).FmaeipDate, 10) %></acronym>
	</td>
	<td>
		<acronym title="<%= oCMaechulPaymentLog.FItemList(i).FmayIpkumDate %>"><%= Left(oCMaechulPaymentLog.FItemList(i).FmayIpkumDate, 10) %></acronym>
	</td>

	<td align="right">
		<% if (oCMaechulPaymentLog.FItemList(i).FrealPayPrice <> "") then %>
			<% if (oCMaechulPaymentLog.FItemList(i).FpayReqPrice <> oCMaechulPaymentLog.FItemList(i).FrealPayPrice) then %><font color="red"><% end if %>
			<%= FormatNumber((oCMaechulPaymentLog.FItemList(i).FrealPayPrice - oCMaechulPaymentLog.FItemList(i).FpayReqPrice), 0) %>
		<% else %>
			<% if (oCMaechulPaymentLog.FItemList(i).FpayReqPrice <> 0) then %><font color="red"><% end if %>
			<%= FormatNumber(-1 * oCMaechulPaymentLog.FItemList(i).FpayReqPrice, 0) %></font>
		<% end if %>
	</td>
	<td>
		<%= oCMaechulPaymentLog.FItemList(i).GetMatchMethodName %>
	</td>

	<td>
		<acronym title="<%= oCMaechulPaymentLog.FItemList(i).Fregdate %>"><%= Left(oCMaechulPaymentLog.FItemList(i).Fregdate, 10) %></acronym>
	</td>
	<td>
		<% if (oCMaechulPaymentLog.FItemList(i).FpayDivCode = "XXX") then %>

			<font color="gray"><%= oCMaechulPaymentLog.FItemList(i).GetMayPayMethodName %></font>
			<% if (oCMaechulPaymentLog.FItemList(i).FactDivCode <> "MM") and (oCMaechulPaymentLog.FItemList(i).FactDivCode <> "EE") and (oCMaechulPaymentLog.FItemList(i).FactDivCode <> "HH") then %>
			<input type="button" class="button" value="CS" class="csbutton" style="width:30px;" onclick="jsPopSearchRefundCS('<%= oCMaechulPaymentLog.FItemList(i).FtargetGbn %>', '<%= oCMaechulPaymentLog.FItemList(i).Forderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fsuborderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Forgorderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fchgorderserial %>', '<%= Left(oCMaechulPaymentLog.FItemList(i).FpayReqDate, 10) %>','<%= oCMaechulPaymentLog.FItemList(i).FpayReqPrice %>');">
			<% end if %>
			<% if (oCMaechulPaymentLog.FItemList(i).FactDivCode = "CC") or (oCMaechulPaymentLog.FItemList(i).FactDivCode = "EE") or (oCMaechulPaymentLog.FItemList(i).FactDivCode = "MM") or (oCMaechulPaymentLog.FItemList(i).FactDivCode = "HH") then %>
				<input type="button" class="button" value="��Ī" class="csbutton" style="width:30px;" onclick="jsNormalizeMatch('<%= oCMaechulPaymentLog.FItemList(i).Forderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fsuborderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Forgorderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).FpayReqPrice %>');">
			<% end if %>
			<% if (oCMaechulPaymentLog.FItemList(i).FactDivCode <> "A") then  %>
				<input type="button" class="button" value="ȯ������" class="csbutton" style="width:60px;" onclick="jsRefundProcMatch('<%= oCMaechulPaymentLog.FItemList(i).Forderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fsuborderserial %>');">
			<% end if %>
			<% if (oCMaechulPaymentLog.FItemList(i).FactDivCode = "C") then %>
				<input type="button" class="button" value="��Ҹ�Ī" class="csbutton" style="width:60px;" onclick="jsMatchCancel('<%= oCMaechulPaymentLog.FItemList(i).Forderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fsuborderserial %>');">
			<% end if %>
            <% if (oCMaechulPaymentLog.FItemList(i).FactDivCode = "M") then %>
				<input type="button" class="button" value="����" class="csbutton" style="width:60px;" onclick="jsMatchCancel('<%= oCMaechulPaymentLog.FItemList(i).Forderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fsuborderserial %>');">
                <input type="button" class="button" value="���" class="csbutton" style="width:60px;" onclick="jsMatchReturn('<%= oCMaechulPaymentLog.FItemList(i).Forderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fsuborderserial %>');">
			<% end if %>

		<% else %>

			<% if (oCMaechulPaymentLog.FItemList(i).FpayDivCode <> "0") and (oCMaechulPaymentLog.FItemList(i).FactDivCode <> "MM") and (oCMaechulPaymentLog.FItemList(i).FactDivCode <> "CC") and (oCMaechulPaymentLog.FItemList(i).FactDivCode <> "EE") then %>
				<input type="button" class="button" value="��Ī" class="csbutton" style="width:30px;" onclick="jsPopSearchRefundCS('<%= oCMaechulPaymentLog.FItemList(i).FtargetGbn %>', '<%= oCMaechulPaymentLog.FItemList(i).Forderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fsuborderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Forgorderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fchgorderserial %>', '<%= Left(oCMaechulPaymentLog.FItemList(i).FpayReqDate, 10) %>','<%= oCMaechulPaymentLog.FItemList(i).FpayReqPrice %>');">
			<% end if %>

			<% if (showDelMatchBtn = "Y") then %>

				<% 'if (oCMaechulPaymentLog.FItemList(i).FPGgubun <> "mileage") then %>
					<input type="button" class="button" value="��Ī����" class="csbutton" style="width:60px;" onclick="jsDelMatch('<%= oCMaechulPaymentLog.FItemList(i).Forderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fsuborderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).FpayDivCode %>');">
				<% 'end if %>
			<% end if %>

			<% 'if (oCMaechulPaymentLog.FItemList(i).FpayReqPrice <> oCMaechulPaymentLog.FItemList(i).FrealPayPrice) then %>
				<input type="button" class="button" value="�α��߰�" class="csbutton" style="width:60px;" onclick="jsAddPayLog('<%= oCMaechulPaymentLog.FItemList(i).Forderserial %>', '<%= oCMaechulPaymentLog.FItemList(i).Fsuborderserial %>');">
			<% 'end if %>

		<% end if %>
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if oCMaechulPaymentLog.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCMaechulPaymentLog.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCMaechulPaymentLog.StartScrollPage to oCMaechulPaymentLog.FScrollCount + oCMaechulPaymentLog.StartScrollPage - 1 %>
			<% if i>oCMaechulPaymentLog.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCMaechulPaymentLog.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<form name="frmAct" method="post" action="/admin/maechul/refundMatchRefund_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="orderserial" value="">
	<input type="hidden" name="suborderserial" value="">
	<input type="hidden" name="orgorderserial" value="">
	<input type="hidden" name="paydivcode" value="">
	<input type="hidden" name="payreqprice" value="">
	<input type="hidden" name="yyyymmdd" value="">
    <input type="hidden" name="startDt" value="">
    <input type="hidden" name="endDt" value="">
    <input type="hidden" name="dateGubun" value="">
	<input type="hidden" name="daypart" value="">
</form>

<%
set oCMaechulPaymentLog = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
