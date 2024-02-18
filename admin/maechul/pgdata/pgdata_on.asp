<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸�
' Hieditor : 2011.04.22 �̻� ����
'			 2023.05.31 �ѿ�� ����(�˻����� �߰� / ������� : ����������, pg���� : convinienspay)
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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/maechul/incMaechulFunction.asp"-->
<%

dim research, page, pagesize
dim excmatchfinish, onlypricenotequal
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim yyyy3,yyyy4,mm3,mm4,dd3,dd4
dim yyyy, mm, dd
dim fromDate ,toDate, tmpDate
dim fromDate2 ,toDate2
dim sitename
dim appDivCode, ipkumdate
dim searchfield, searchtext
dim PGuserid, appMethod
dim pggubun
dim showjumunlog, showjumunlogNotMatch, chkSearchIpkumDate, chkSearchAppDate
dim reasonGubun

Dim i

	research = requestCheckvar(request("research"),10)
	page = requestCheckvar(request("page"),10)
	excmatchfinish = requestCheckvar(request("excmatchfinish"),10)
	onlypricenotequal = requestCheckvar(request("onlypricenotequal"),10)

	yyyy1   = requestCheckvar(request("yyyy1"),32)
	mm1     = requestCheckvar(request("mm1"),32)
	dd1     = requestCheckvar(request("dd1"),32)
	yyyy2   = requestCheckvar(request("yyyy2"),32)
	mm2     = requestCheckvar(request("mm2"),32)
	dd2     = requestCheckvar(request("dd2"),32)

	yyyy3   = requestCheckvar(request("yyyy3"),32)
	mm3     = requestCheckvar(request("mm3"),32)
	dd3     = requestCheckvar(request("dd3"),32)
	yyyy4   = requestCheckvar(request("yyyy4"),32)
	mm4     = requestCheckvar(request("mm4"),32)
	dd4     = requestCheckvar(request("dd4"),32)

	sitename		= requestCheckvar(request("sitename"),32)
	appDivCode 		= requestCheckvar(request("appDivCode"),32)
	ipkumdate 		= requestCheckvar(request("ipkumdate"),32)
	PGuserid 		= requestCheckvar(request("PGuserid"),32)
	appMethod 		= requestCheckvar(request("appMethod"),32)

	searchfield 	= requestCheckvar(request("searchfield"),32)
	searchtext 		= Replace(Replace(requestCheckvar(request("searchtext"),64), "'", ""), Chr(34), "")

	pggubun 		= requestCheckvar(request("pggubun"),32)
	reasonGubun 	= requestCheckvar(request("reasonGubun"),32)

	showjumunlog 				= requestCheckvar(request("showjumunlog"),32)
	showjumunlogNotMatch 		= requestCheckvar(request("showjumunlogNotMatch"),32)
	chkSearchIpkumDate 			= requestCheckvar(request("chkSearchIpkumDate"),32)
	chkSearchAppDate 			= requestCheckvar(request("chkSearchAppDate"),32)
	pagesize					= requestCheckvar(request("pagesize"),32)

if (chkSearchIpkumDate="") then chkSearchAppDate = "Y"
if (page="") then page = 1
if (pagesize="") then pagesize = 100

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

	fromDate2 = fromDate
	toDate2 = toDate
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

if (yyyy3="") then
	fromDate2 = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate2 = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	yyyy3 = Cstr(Year(fromDate2))
	mm3 = Cstr(Month(fromDate2))
	dd3 = Cstr(day(fromDate2))

	tmpDate = DateAdd("d", -1, toDate2)
	yyyy4 = Cstr(Year(tmpDate))
	mm4 = Cstr(Month(tmpDate))
	dd4 = Cstr(day(tmpDate))
else
	fromDate2 = DateSerial(yyyy3, mm3, dd3)
	toDate2 = DateSerial(yyyy4, mm4, dd4+1)
end if

Dim oCPGData
set oCPGData = new CPGData
	oCPGData.FPageSize = pagesize
	oCPGData.FCurrPage = page

	oCPGData.FRectExcMatchFinish   	= excmatchfinish
	oCPGData.FRectOnlyPriceNotEqual   	= onlypricenotequal

	if (chkSearchAppDate = "Y") and (chkSearchIpkumDate = "Y") then
		oCPGData.FRectDateType = "A"
	elseif (chkSearchIpkumDate = "Y") then
		oCPGData.FRectDateType = "B"
	else
		oCPGData.FRectDateType = ""
	end if

	if (chkSearchAppDate = "Y") then
		oCPGData.FRectStartdate = fromDate
		oCPGData.FRectEndDate = toDate
	end if

	if (chkSearchIpkumDate = "Y") then
		oCPGData.FRectStartIpkumdate = fromDate2
		oCPGData.FRectEndIpkumDate = toDate2
	end if

	oCPGData.FRectPGuserid = PGuserid
	oCPGData.FRectAppMethod = appMethod
	oCPGData.FRectSiteName = sitename
	oCPGData.FRectAppDivCode = appDivCode
	oCPGData.FRectIpkumdate = ipkumdate

	oCPGData.FRectSearchField = searchfield
	oCPGData.FRectSearchText = searchtext

	oCPGData.FRectPGGubun = pggubun
	oCPGData.FRectReasonGubun = reasonGubun

	oCPGData.FRectShowJumunLog 			= showjumunlog
	oCPGData.FRectShowJumunLogNotMatch 	= showjumunlogNotMatch

    oCPGData.getPGDataList_ON

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function jsGetOnPGData(pgid) {
	var frm = document.frmAct;
	var yyyymmdd = document.getElementById("yyyymmdd");

	if (pgid == "inicis") {
		frm.mode.value = "getonpgdata";
        alert('�ߺ�');
        return;
	} else if (pgid == "inicishp") {
		frm.mode.value = "getonpgdatahp";
        alert('�ߺ�');
        return;
	} else if (pgid == "uplus") {
		frm.mode.value = "getonpgdatauplus";
        alert('�ߺ�');
        return;
	} else if (pgid == "kakaopayT") {
		// īī��PAY �ŷ�����
		frm.mode.value = "getonpgdatakakaopayT";
	} else if (pgid == "kakaopayS") {
		// īī��PAY ���곻��
		frm.mode.value = "getonpgdatakakaopayS";
	} else if (pgid == "newkakaopayT") {
		// īī��PAY �ŷ�����
		frm.mode.value = "getonpgdatanewkakaopayT";
        alert('�ߺ�');
        return;
	} else if (pgid == "newkakaopayS") {
		// īī��PAY ���곻��
		frm.mode.value = "getonpgdatanewkakaopayS";
        alert('�ߺ�');
        return;
	} else if (pgid == "naverpay") {
		// ���̹�����
		frm.mode.value = "getonpgdatanaverpay";
	} else if (pgid == "gifticon") {
		frm.mode.value = "getonpgdatagifticon";
	} else if (pgid == "giftting") {
		frm.mode.value = "getonpgdatagiftting";
	} else if (pgid == "paycoT") {
		frm.mode.value = "getpaycoT";
	} else if (pgid == "paycoS") {
		frm.mode.value = "getpaycoS";
	} else if (pgid == "toss") {
		if (yyyymmdd.value.length == 10) {
			alert(yyyymmdd.value);
			popGetToss(yyyymmdd.value);
		} else {
			popGetToss("");
		}
		return;
	} else if (pgid == "tossdue") {
        // �佺�� �ŷ������� �������� ���ԵǾ� �ִ�.
		if (yyyymmdd.value.length == 10) {
			popGetTossDue(yyyymmdd.value);
		} else {
			popGetTossDue("");
		}
		return;
	} else if (pgid == "chaiT") {
		// �������� ���� �ŷ�����
		frm.mode.value = "getonpgdatachaipayT";
	} else if (pgid == "chaiS") {
		// �������� ���� �ŷ�����
		frm.mode.value = "getonpgdatachaipayS";
        alert('�ߺ�');
        return;
    } else if (pgid = 'appMethod6') {
        // �������Ա�
        frm.mode.value = "getonpgdatacappMethod6";
	} else {
		alert("ERROR");
		return;
	}

	if ((pgid == "paycoT") || (pgid == "paycoS") || (pgid == "kakaopayT") || (pgid == "kakaopayS") || (pgid == "newkakaopayT") || (pgid == "newkakaopayS") || (pgid == "uplus") || (pgid == "toss") || (pgid == "chaiT") || (pgid == "chaiS") || (pgid == "inicis") || (pgid == "appMethod6")) {
		if (yyyymmdd.value.length == 10) {
			alert(yyyymmdd.value);
			frm.yyyymmdd.value = yyyymmdd.value;
		} else {
			frm.yyyymmdd.value = "";
		}
	}

	if (pgid == "uplus") {
		var frmUplus = document.frm;
		if ((frmUplus.searchfield.value == "orderserial") && (frmUplus.searchtext.value != "")) {
			if (confirm("PG����Ÿ(ON " + pgid + ") �� �������� �Ͻðڽ��ϱ�?\n\n�ߺ��ֹ���ȣ(" + frmUplus.searchtext.value + ")") == true) {
				frm.orderserial.value = frmUplus.searchtext.value;
				frm.submit();
			}
		} else {
			if (confirm("PG����Ÿ(ON " + pgid + ") �� �������� �Ͻðڽ��ϱ�?") == true) {
				frm.submit();
			}
		}
	} else if (pgid == "newkakaopayT") {
		var frmUplus = document.frm;
		if ((frmUplus.searchfield.value == "orderserial") && (frmUplus.searchtext.value != "")) {
			if (confirm("PG����Ÿ(ON " + pgid + ") �� �������� �Ͻðڽ��ϱ�?\n\n�ֹ���ȣ(" + frmUplus.searchtext.value + ")") == true) {
				frm.orderserial.value = frmUplus.searchtext.value;
				frm.submit();
			}
		} else {
			if (confirm("PG����Ÿ(ON " + pgid + ") �� �������� �Ͻðڽ��ϱ�?") == true) {
				frm.submit();
			}
		}
	} else {
		if (confirm("PG����Ÿ(ON " + pgid + ") �� �������� �Ͻðڽ��ϱ�?") == true) {
			frm.submit();
		}
	}

}

function popGetToss(yyyymmdd) {
	var url = "http://wapi.10x10.co.kr/toss/api.asp?mode=settle&yyyymmdd=" + yyyymmdd;
	var popwin = window.open(url,"popGetToss","width=500 height=300 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popGetTossDue(yyyymmdd) {
	var url = "http://wapi.10x10.co.kr/toss/api.asp?mode=due&yyyymmdd=" + yyyymmdd;
	var popwin = window.open(url,"popGetTossDue","width=500 height=300 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsMatchPGData() {
	var frm = document.frmAct;

	frm.mode.value = "matchpgdata";

	if (confirm("�ڵ���Ī(10x10) �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsMatchEtcPayment() {
	var frm = document.frmAct;

	frm.mode.value = "matchetcpay";

	if (confirm("�ڵ���Ī(�������Ա�) �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

<% if (searchfield = "PGkey") and (searchtext <> "") then %>
function jsMatchPGDataOld() {
	var frm = document.frmAct;

	frm.mode.value = "matchpgdata6month";
	frm.PGKey.value = "<%= searchtext %>";

	if (confirm("�ڵ���Ī(10x10,6��������) �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}
<% end if %>

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

function popUploadNAVERPAYPGData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("popRegNAVERPAYPGDataFile_on.asp","popUploadNAVERPAYPGData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popUploadMobilPGData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("popRegKCPPGDataFile_on.asp?pgid=mobilians","popUploadKCPPGData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popUploadONPGData(pgid) {
    var window_width = 500;
    var window_height = 250;

	if (pgid == "gifticon") {
		// frm.mode.value = "getonpgdatagifticon";
	} else if (pgid == "giftting") {
		// frm.mode.value = "getonpgdatagiftting";
	} else {
		alert("ERROR");
		return;
	}

    var popwin = window.open("popRegPGDataFile_on.asp?pgid=" + pgid,"popUploadONPGData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popExtSiteJungsanData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("/admin/maechul/extjungsandata/popRegExtJungsanDataFile.asp","popExtSiteJungsanData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsPopInputOrderSerial(idx) {
	var v = "popMatchOrderSerial.asp?idx=" + idx;
	var popwin = window.open(v,"jsPopInputOrderSerial","width=400,height=300,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jsMatchCancel(logidx, datediff) {
	var frm = document.frmAct;

	frm.logidx.value = logidx;
	frm.mode.value = "matchcancelOnline";

	if (confirm("[���]���� ��Ī �Ͻðڽ��ϱ�?") == true) {
		if (datediff == true) {
			<% if (searchfield = "PGkey") and (searchtext <> "") then %>
			frm.PGKey.value = "<%= searchtext %>";
			<% end if %>
			frm.force.value = "Y";
		}
		frm.submit();
	}
}

function jsAddRefundLog(logidx) {
	var frm = document.frmAct;

	frm.logidx.value = logidx;
	frm.mode.value = "addActLog";

	if (confirm("���γ���(0��) �� �߰� �Ͻðڽ��ϱ�?") == true) {
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

function jsPopSumIpkum(idx) {
	var v = "popMatchSumIpkum.asp?idx=" + idx;
	var popwin = window.open(v,"jsPopSumIpkum","width=250,height=300,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jsPopRegReasonGubun(idx) {
	var v = "popRegReasonGubun.asp?idx=" + idx;
	var popwin = window.open(v,"jsPopRegReasonGubun","width=250,height=150,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jsPopModiAppDate(idx, gubun) {
	var v = "popModiAppDate.asp?idx=" + idx + '&gubun=' + gubun;
	var popwin = window.open(v,"jsPopModiAppDate","width=250,height=150,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function popCsList(csid){
    var window_width = 1280;
    var window_height = 960;
    //searchfield=asid&searchstring=2028907&divcd=&currstate=&delYN=N&periodYN=Y&yyyy1=2014&mm1=06&dd1=01&yyyy2=2014&mm2=09&dd2=01&extsitename=11stITS
	var popwin = window.open("/cscenter/action/cs_action.asp?searchfield=asid&searchstring=" + csid,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function jsRegReasonGubunarr() {
	if (frm.selectreasonGubun.value==""){
		alert("�󼼻����� ���� �ϼ���.");
		frm.selectreasonGubun.focus();
		return;
	}

	frm.mode.value = "RegReasonGubunarr";

	if (confirm("�ϰ����� �Է� �Ͻðڽ��ϱ�?") == true) {
		frm.action="/admin/maechul/pgdata/pgdata_process.asp";
		frm.submit();
	}
}
<% if (C_ADMIN_AUTH) then %>
function jsDelOne(idx) {
    var frm = document.frmAct;

	if (confirm("�ߺ� �ֹ��Ǹ� ���� ������ ����Դϴ�.\n������ ���� �Ͻðڽ��ϱ�?") == true) {
        frm.mode.value = "delapplog";
        frm.logidx.value = idx;
		frm.submit();
	}
}
<% end if %>
<% if (C_ADMIN_AUTH) then %>
function jsIniRentalCancel(pgkey) {
    var frm = document.frmAct;

	if (confirm("���⼭ ��� ó���� �ϱ����� �̴Ͻý� ���ο��� ��� ó�� �ؾ� �˴ϴ�.\n���ó�� �� cs���ó���� �������� �ؾ� �˴ϴ�.\n���ó�� �Ͻðڽ��ϱ�?") == true) {
        frm.mode.value = "inirentalcancel";
        frm.PGKey.value = pgkey;
		frm.submit();
	}
}
<% end if %>
function jsRegReasonGubun025() {
	<% if (chkSearchAppDate = "Y") and (appMethod = "77") and (pggubun = "bankrefund") and (PGuserid = "bankrefund_10x10") then %>
	frm.mode.value = "RegReasonGubun025";

	if (confirm("������(��ġ��ȯ��) �ϰ��Է� �Ͻðڽ��ϱ�?") == true) {
		frm.action="/admin/maechul/pgdata/pgdata_process.asp";
		frm.submit();
	}
	<% else %>
	alert('�Ʒ� �������� �˻��� ��츸 �Է°����մϴ�.\n\n - ����(���)���� üũ\n - ������� : ������ȯ��\n - PG�� : bankrefund\n - PG��id : bankrefund_10x10\n - �󼼻��� : �Է�����');
	return;
	<% end if %>
}

function popUploadHandData() {
	var popwin = window.open("popRegHand_on.asp","popUploadHandData","width=600 height=300 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popUploadIniRentalData() {
	var popwin = window.open("popRegIniRentalManualWrite_on.asp","popUploadHandData","width=600 height=300 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;
		<input type="checkbox" name="chkSearchAppDate"  value="Y" <% if (chkSearchAppDate = "Y") then %>checked<% end if %> > ����(���)����:
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<input type="checkbox" name="chkSearchIpkumDate"  value="Y" <% if (chkSearchIpkumDate = "Y") then %>checked<% end if %> > �Աݿ�����:
		<% DrawDateBoxdynamic yyyy3, "yyyy3", yyyy4, "yyyy4", mm3, "mm3", mm4, "mm4", dd3, "dd3", dd4, "dd4"  %>
		&nbsp;
		* ���α��� :
		<select class="select" name="appDivCode">
		<option value=""></option>
		<option value="A" <% if (appDivCode = "A") then %>selected<% end if %> >����</option>
		<option value="C" <% if (appDivCode = "C") then %>selected<% end if %> >���</option>
		<option value="R" <% if (appDivCode = "R") then %>selected<% end if %> >�κ����</option>
		<option value="">----</option>
		<option value="E" <% if (appDivCode = "E") then %>selected<% end if %> >����</option>
		</select>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		* �Աݿ����� :
		<input type="text" class="text" name="ipkumdate" value="<%= ipkumdate %>" size="10">
		&nbsp;
		* ǥ�ð��� :
		<select class="select" name="pagesize">
			<option value="100">100</option>
			<option value="500" <%= CHKIIF(pagesize="500", "selected", "")%> >500</option>
			<option value="1000" <%= CHKIIF(pagesize="1000", "selected", "")%> >1000</option>
			<option value="2500" <%= CHKIIF(pagesize="2500", "selected", "")%> >2500</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		* ������� :
		<select class="select" name="appMethod">
			<option value=""></option>
			<option value="7" <% if (appMethod = "7") then %>selected<% end if %> >������(����)</option>
			<option value="14" <% if (appMethod = "14") then %>selected<% end if %> >����������</option>
			<option value="100" <% if (appMethod = "100") then %>selected<% end if %> >�ſ�</option>
			<option value="20" <% if (appMethod = "20") then %>selected<% end if %> >�ǽð�</option>
			<option value="80" <% if (appMethod = "80") then %>selected<% end if %> >All@</option>
			<option value="110" <% if (appMethod = "110") then %>selected<% end if %> >OKĳ�ù�</option>
			<option value="400" <% if (appMethod = "400") then %>selected<% end if %> >�ڵ���</option>
			<option value="550" <% if (appMethod = "550") then %>selected<% end if %> >������</option>
			<option value="560" <% if (appMethod = "560") then %>selected<% end if %> >����Ƽ��</option>
            <option value="150" <% if (appMethod = "150") then %>selected<% end if %> >�̴Ϸ�Ż</option>
			<option value="">---------</option>
			<option value="77" <% if (appMethod = "77") then %>selected<% end if %> >������ȯ��</option>
			<option value="6" <% if (appMethod = "6") then %>selected<% end if %> >�������Ա�</option>
		</select>
		&nbsp;
		* PG�� :
		<select name="pggubun" class="select">
			<option value="">--����--</option>
			<%Call sbGetOptPGgubun(pggubun)%>
		</select>
		<% 'Call DrawSelectBoxPGGubun("pggubun", pggubun, "") %>
		&nbsp;
		* PG��id :
		<select name="PGuserid" class="select">
			<option value="">--����--</option>
			<%Call sbGetOptPGID(PGuserid)%>
		</select>
		<% 'Call DrawSelectBoxPGUserid("PGuserid", PGuserid, "") %>
		&nbsp;
		* �󼼻��� :
		<select class="select" name="reasonGubun">
		<option value=""></option>
		<option value="001" <% if (reasonGubun = "001") then %>selected<% end if %> >������(����)</option>
		<option value="002" <% if (reasonGubun = "002") then %>selected<% end if %> >������(���޻� ����)</option>
        <option value="003" <% if (reasonGubun = "003") then %>selected<% end if %> >������(�̴Ϸ�Ż)</option>
		<option value="020" <% if (reasonGubun = "020") then %>selected<% end if %> >������(��ġ��)</option>
		<option value="025" <% if (reasonGubun = "025") then %>selected<% end if %> >������(��ġ��ȯ��)</option>
		<option value="030" <% if (reasonGubun = "030") then %>selected<% end if %> >������(����Ʈ)</option>
		<option value="035" <% if (reasonGubun = "035") then %>selected<% end if %> >������(����Ʈȯ��)</option>
        <option value="004" <% if (reasonGubun = "004") then %>selected<% end if %> >������(B2B ����)</option>
		<option value="">---------------</option>
		<option value="040" <% if (reasonGubun = "040") then %>selected<% end if %> >CS����</option>
		<option value="">---------------</option>
		<option value="950" <% if (reasonGubun = "950") then %>selected<% end if %> >�������Ȯ��</option>
		<option value="999" <% if (reasonGubun = "999") then %>selected<% end if %> >��Ҹ�Ī</option>
		<option value="901" <% if (reasonGubun = "901") then %>selected<% end if %> >�ΰŽ����ݸ���</option>
		<option value="800" <% if (reasonGubun = "800") then %>selected<% end if %> >���ڼ���</option>
		<option value="900" <% if (reasonGubun = "900") then %>selected<% end if %> >��Ÿ</option>
		<option value="">---------------</option>
		<option value="XXX" <% if (reasonGubun = "XXX") then %>selected<% end if %> >�Է�����</option>
		</select>
		&nbsp;
		* ����Ʈ :
		<select class="select" name="sitename">
		<option value=""></option>
		<option value="10x10" <% if (sitename = "10x10") then %>selected<% end if %> >10x10(PC)</option>
		<option value="10x10mobile" <% if (sitename = "10x10mobile") then %>selected<% end if %> >10x10(MOBILE)</option>
		<option value="fingers" <% if (sitename = "fingers") then %>selected<% end if %> >�ΰŽ�</option>
		<option value="10x10gift" <% if (sitename = "10x10gift") then %>selected<% end if %> >10x10(GIFT)</option>
        <option value="wholesale" <% if (sitename = "wholesale") then %>selected<% end if %> >WHOLESALE</option>
		</select>
		&nbsp;
		* �˻����� :
		<select class="select" name="searchfield">
		<option value=""></option>
		<option value="PGkey" <% if (searchfield = "PGkey") then %>selected<% end if %> >PG��KEY</option>
		<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >�ֹ���ȣ</option>
		<option value="appPrice" <% if (searchfield = "appPrice") then %>selected<% end if %> >�ŷ��ݾ�</option>
		</select>
		<input type="text" class="text" name="searchtext" value="<%= searchtext %>" size="50">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		<input type="checkbox" name="excmatchfinish"  value="Y" <% if (excmatchfinish = "Y") then %>checked<% end if %> > ��Ī�Ϸ�(���ΰ� �ֹ���ȣ��Ī, ��Ұ� CS������Ī) ����
		&nbsp;
		<input type="checkbox" name="onlypricenotequal"  value="Y" <% if (onlypricenotequal = "Y") then %>checked<% end if %> > ���ֹ� ���αݾ� ���̳�����(30�� ��������)
		&nbsp;
		<input type="checkbox" name="showjumunlog"  value="Y" <% if (showjumunlog = "Y") then %>checked<% end if %> > �����α� ǥ��(30�� ��������)
		&nbsp;
		<input type="checkbox" name="showjumunlogNotMatch"  value="Y" <% if (showjumunlogNotMatch = "Y") then %>checked<% end if %> > <b>�����α� ��Ī�Ϸ� ����(30�� ��������, ������ �ٸ���� ǥ��)</b>
	</td>
</tr>
</table>
<!-- �˻� �� -->

<h5>�׽�Ʈ��...</h5>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;" border="0">
<tr>
	<td align="left" width="50%">
		- ��ſ����� ���� �ֹ��Է� ���� ���� �� ��� ��Ī<br>
		* �ǽð���ü(������ ���� ���)�� ��� ������ �ݾ��� ������� �ʴ´�.<br>
		* �̴Ͻý� �ָ� ���γ����� �Աݿ������� ������ �ݿ����̴�.<br>
		* <font color="red">�հ��Ա�</font>�� �켱 1���� �ֹ���ȣ�� ��Ī�� ���Ŀ� �߰��� �Է°����ϴ�.<br />
		* ���̹����� �ǽð���ü�� ����� ���, �������� ���γ����� �ٽ� �ٿ�޾ƾ� �Ѵ�.(http://wapi.10x10.co.kr/nPay/jungsanReceive.asp)<br />
		* UPLUS PK �ߺ� ���� �ִ� ���, <font color="red">�ش� �ֹ���ȣ �˻� �� ���γ��� ��������</font> ������ �������� �˴ϴ�.
	</td>
	<td align="left">
		* <font color="red"><b>�󼼻��� �ڵ��Է�</b></font> : �����ڷ� �ٿ�ε� -&gt; �ֹ���ȣ��Ī -&gt; �����α׸�Ī(30�а���ʿ�) -&gt; �̼������ۼ�
	</td>
</tr>
</table>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;" border="0">
<tr>
	<td align="left">
        <!--
		<input type="button" class="button" value="��������(ON INICIS)" onClick="jsGetOnPGData('inicis');" disabled>
		<input type="button" class="button" value="��������(ON INICIS HP)" onClick="jsGetOnPGData('inicishp');" disabled>
		<input type="button" class="button" value="��������(ON UPLUS)" onClick="jsGetOnPGData('uplus');" disabled>
        -->
        <!--
		<input type="button" class="button" value="����ϱ�(ON ���̹���������)" onClick="popUploadNAVERPAYPGData();">
        -->
        <input type="button" class="button" value="��������(ON ������)" onClick="jsGetOnPGData('appMethod6');">
		<br><br>
        <!--
		<input type="button" class="button" value="��������(ON NewKAKAO �ŷ�)" onClick="jsGetOnPGData('newkakaopayT');" disabled>
		<input type="button" class="button" value="��������(ON NewKAKAO ����)" onClick="jsGetOnPGData('newkakaopayS');" disabled>
        -->
        <!--
		<input type="button" class="button" value="��������(������ �ŷ�)" onClick="jsGetOnPGData('paycoT');">
		<input type="button" class="button" value="��������(������ ����)" onClick="jsGetOnPGData('paycoS');">
        -->
        <!--
		<input type="button" class="button" value="��������(�佺)" onClick="jsGetOnPGData('toss');">
        -->
        <!--
		<input type="button" class="button" value="��������(���� �ŷ�)" onClick="jsGetOnPGData('chaiT');">
        -->
        <!--
		<input type="button" class="button" value="��������(���� ����)" onClick="jsGetOnPGData('chaiS');" disabled>
        -->
		<input type="text" class="text" id="yyyymmdd" name="yyyymmdd" value="" size="12">
		<!--
		<br><br>
		<input type="button" class="button" value="��������(ON ���̹�����)" onClick="jsGetOnPGData111('naverpay');">
		-->
		<br><br>
		<input type="button" class="button" value="����ϱ�(����Ƽ��)" onClick="popUploadONPGData('gifticon');">
		<input type="button" class="button" value="����ϱ�(������)" onClick="popUploadONPGData('giftting');">
		<input type="button" class="button" value="����ϱ�(����)" onClick="popUploadHandData();">
		<% If session("ssBctId") = "thensi7" Then %>
			<input type="button" class="button" value="�̴Ϸ�Ż ������(����)" onClick="popUploadIniRentalData();">
		<% End If %>
	</td>
	<td align="right">
		<input type="button" class="button" value="�ڵ���Ī(�������Ա�)" onClick="jsMatchEtcPayment();">
        <input type="button" class="button" value="�ڵ���Ī(10x10)" onClick="jsMatchPGData();">
		<input type="button" class="button" value="�ڵ���Ī(�ΰŽ�)" onClick="jsMatchFingersPGData();">
		<input type="button" class="button" value="�ڵ���Ī(����Ʈ)" onClick="jsMatchGiftCardPGData();">
		<br /><br />
		<% if (searchfield = "PGkey") and (searchtext <> "") then %>
		<input type="button" class="button" value="�ڵ���Ī(10x10,6��������,<%= searchtext %>)" onClick="jsMatchPGDataOld();">
		<% end if %>

		<% if PGuserid <> "" then %>
		<br>
		<input type="button" class="button" value="������(��ġ��ȯ��) �ϰ��Է�" onClick="jsRegReasonGubun025();" style="width:180px;"> &nbsp;
			* �󼼻��� :
			<select class="select" name="selectreasonGubun">
			<option value=""></option>
			<option value="001" <% if (reasonGubun = "001") then %>selected<% end if %> >������(����)</option>
			<option value="002" <% if (reasonGubun = "002") then %>selected<% end if %> >������(���޻� ����)</option>
			<option value="020" <% if (reasonGubun = "020") then %>selected<% end if %> >������(��ġ��)</option>
			<option value="025" <% if (reasonGubun = "025") then %>selected<% end if %> >������(��ġ��ȯ��)</option>
			<option value="030" <% if (reasonGubun = "030") then %>selected<% end if %> >������(����Ʈ)</option>
			<option value="035" <% if (reasonGubun = "035") then %>selected<% end if %> >������(����Ʈȯ��)</option>
			<option value="">---------------</option>
			<option value="040" <% if (reasonGubun = "040") then %>selected<% end if %> >CS����</option>
			<option value="">---------------</option>
			<option value="950" <% if (reasonGubun = "950") then %>selected<% end if %> >�������Ȯ��</option>
			<option value="999" <% if (reasonGubun = "999") then %>selected<% end if %> >��Ҹ�Ī</option>
			<option value="901" <% if (reasonGubun = "901") then %>selected<% end if %> >�ΰŽ����ݸ���</option>
			<option value="800" <% if (reasonGubun = "800") then %>selected<% end if %> >���ڼ���</option>
			<option value="900" <% if (reasonGubun = "900") then %>selected<% end if %> >��Ÿ</option>
			<option value="">---------------</option>
			<option value="XXX" <% if (reasonGubun = "XXX") then %>selected<% end if %> >�Է�����</option>
			</select>
			<input type="button" class="button" value="�����ϰ��Է�" onClick="jsRegReasonGubunarr();" style="width:100px;">
		<% end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->
</form>
<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21">
		�˻���� : <b><%= oCPGData.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCPGData.FTotalPage %></b>
        &nbsp;
        �ŷ��Ѿ� : <b><%= FormatNumber(oCPGData.FTotalAppPrice, 0) %> ��</b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>PG��</td>
	<td>PG��id</td>
	<td width="80">�������</td>
	<td>PG��KEY</td>
	<td>PG��CSKEY</td>
	<td width="60">����</td>
	<td width="150">����(���)����</td>
	<td width="60">�ŷ���</td>
	<td width="60">������<br>(VAT����)</td>
	<td width="60">�Ա�<br>������</td>
	<td width="65">ī���<br>������</td>
	<td width="70">�Աݿ�����</td>
	<td>����Ʈ</td>
	<td>�ֹ���ȣ</td>
	<td width="60">CSIDX</td>
	<% if (showjumunlog = "Y") then %>
	<td>�����α�</td>
	<% end if %>
	<td>�󼼻���</td>
	<!--
	<td width="80">�����</td>
	-->
	<td>���</td>
</tr>

<% for i=0 to oCPGData.FresultCount -1 %>
<%
yyyy = Left(oCPGData.FItemList(i).FappDate, 4)
mm = Right(Left(oCPGData.FItemList(i).FappDate, 7), 2)
dd = Right(Left(oCPGData.FItemList(i).FappDate, 10), 2)

%>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCPGData.FItemList(i).FPGgubun %></td>
	<td><%= oCPGData.FItemList(i).FPGuserid %></td>
	<td><%= oCPGData.FItemList(i).GetAppMethodName %></td>
	<td><%= oCPGData.FItemList(i).FPGkey %></td>
	<td><%= oCPGData.FItemList(i).FPGCSkey %></td>
	<td>
		<font color="<%= oCPGData.FItemList(i).GetAppDivCodeColor %>"><%= oCPGData.FItemList(i).GetAppDivCodeName %></font>
	</td>
	<td>
        <a href="javascript:jsPopModiAppDate(<%= oCPGData.FItemList(i).Fidx %>, 'appDate')">
		<% if Not IsNull(oCPGData.FItemList(i).FcancelDate) then %>
			<%= oCPGData.FItemList(i).FcancelDate %>
		<% else %>
			<%= oCPGData.FItemList(i).FappDate %>
		<% end if %>
        </a>
	</td>
	<td align="right"><%= FormatNumber(oCPGData.FItemList(i).FappPrice, 0) %></td>
	<td align="right"><%= FormatNumber((oCPGData.FItemList(i).FcommPrice + oCPGData.FItemList(i).FcommVatPrice), 0) %></td>
	<td align="right"><%= FormatNumber(oCPGData.FItemList(i).FjungsanPrice, 0) %></td>
	<td><%= oCPGData.FItemList(i).Fpgmeachuldate %></td>
	<td>
        <a href="javascript:jsPopModiAppDate(<%= oCPGData.FItemList(i).Fidx %>, 'ipkumDate')">
            <%= oCPGData.FItemList(i).Fipkumdate %>
            <%= CHKIIF(IsNull(oCPGData.FItemList(i).Fipkumdate), "-", "") %>
        </a>
    </td>
	<td>
		<%= oCPGData.FItemList(i).Fsitename %>
	</td>
	<td>
		<% if IsNumeric(oCPGData.FItemList(i).Forderserial) then %>
		<a href="javascript:Cscenter_Action_List('<%= oCPGData.FItemList(i).FOrderSerial %>','','')"><%= oCPGData.FItemList(i).Forderserial %></a>
        <a href="javascript:jsPopInputOrderSerial(<%= oCPGData.FItemList(i).Fidx %>)">X</a>
        <% elseif IsNull(oCPGData.FItemList(i).Forderserial) or oCPGData.FItemList(i).Forderserial = "" then %>
        <input type="button" class="button" value="�Է�" onClick="jsPopInputOrderSerial(<%= oCPGData.FItemList(i).Fidx %>)">
		<% else %>
		<%= oCPGData.FItemList(i).Forderserial %>
        <a href="javascript:jsPopInputOrderSerial(<%= oCPGData.FItemList(i).Fidx %>)">X</a>
		<% end if %>
	</td>
	<td><a href="javascript:popCsList('<%= oCPGData.FItemList(i).Fcsasid %>');"><%= oCPGData.FItemList(i).Fcsasid %></a></td>
	<% if (showjumunlog = "Y") then %>
	<td><%= oCPGData.FItemList(i).GetFullLogOrderSerial %></td>
	<% end if %>
	<td><%= oCPGData.FItemList(i).GetReasonGubunName %></td>
	<!--
	<td><%= Left(oCPGData.FItemList(i).Fregdate, 10) %></td>
	-->
	<td>
		<% if IsNull(oCPGData.FItemList(i).Forderserial) and (oCPGData.FItemList(i).FappDivCode = "C") then %>
			<input type="button" class="button" value="��Ҹ�Ī" onClick="jsMatchCancel(<%= oCPGData.FItemList(i).Fidx %>, false);">
			<% if (searchfield = "PGkey") and (searchtext <> "") then %>
			<input type="button" class="button" value="��Ҹ�Ī(�ٸ���¥)" onClick="jsMatchCancel(<%= oCPGData.FItemList(i).Fidx %>, true);">
			<% end if %>
		<% elseif Not IsNull(oCPGData.FItemList(i).Forderserial) and (oCPGData.FItemList(i).FappDivCode = "C") and IsNull(oCPGData.FItemList(i).Fcsasid) then %>
			<input type="button" class="button" value="�ߺ�������Ҹ�Ī" onClick="jsDuplicateMatchCancel(<%= oCPGData.FItemList(i).Fidx %>);">
		<% end if %>
		<% if (oCPGData.FItemList(i).FPGgubun = "bankipkum") and (oCPGData.FItemList(i).FappDivCode <> "C") and (oCPGData.FItemList(i).FappPrice >= 1000) and (oCPGData.FItemList(i).Forderserial <> "") then %>
			<input type="button" class="button" value="�հ��Ա�" onClick="jsPopSumIpkum(<%= oCPGData.FItemList(i).Fidx %>)">
		<% end if %>
		<% if (oCPGData.FItemList(i).FPGgubun = "bankrefund") and (oCPGData.FItemList(i).FappDivCode <> "A") and (oCPGData.FItemList(i).FappPrice <> 0) then %>
			<input type="button" class="button" value="�����߰�(0��)" onClick="jsAddRefundLog(<%= oCPGData.FItemList(i).Fidx %>)">
		<% end if %>
		<% if IsNull(oCPGData.FItemList(i).FreasonGubun) or Not (InStr("001,020,030,950", oCPGData.FItemList(i).FreasonGubun) > 0) or C_ADMIN_AUTH or C_MngPart or C_PSMngPart then %>
			<input type="button" class="button" value="����" onClick="jsPopRegReasonGubun(<%= oCPGData.FItemList(i).Fidx %>)">
			<%' ������ �����ڸ� ���� �ǰ� �ߴ��ǵ� �α׸� ����Ƿ� ���� Ǯ���� %>
			<input type="button" class="button" value="����" onClick="jsDelOne(<%=oCPGData.FItemList(i).Fidx %>)">
			<%' if (C_ADMIN_AUTH or C_MngPart or C_PSMngPart) then %><!--[������]
            <input type="button" class="button" value="����" onClick="jsDelOne(<%'oCPGData.FItemList(i).Fidx %>)">-->
			<%' end if %>
			<%' �̴Ϸ�Ż ���� ���%>
			<% If (C_ADMIN_AUTH) Then %>
				<% If Trim(oCPGData.FItemList(i).FPGuserid) = "teenxteenr" Then %>
					<% If Trim(oCPGData.FItemList(i).GetAppDivCodeName) = "����" Then %>
						<input type="button" class="button" value="���" onClick="jsIniRentalCancel('<%=oCPGData.FItemList(i).FPGkey %>')">
					<% End If %>
				<% End If %>
			<% End If %>
		<% end if %>
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="21" align="center">
		<% if oCPGData.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCPGData.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCPGData.StartScrollPage to oCPGData.FScrollCount + oCPGData.StartScrollPage - 1 %>
			<% if i>oCPGData.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCPGData.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<%
set oCPGData = Nothing
%>

<form name="frmAct" method="post" action="/admin/maechul/pgdata/pgdata_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="logidx" value="">
<input type="hidden" name="yyyymmdd" value="">
<input type="hidden" name="PGKey" value="">
<input type="hidden" name="force" value="">
<input type="hidden" name="orderserial" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
