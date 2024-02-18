<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰
' Hieditor : 2011.04.22 이상구 생성
'			 2023.05.31 한용민 수정(검색조건 추가 / 결제방식 : 편의점결제, pg구분 : convinienspay)
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
        alert('중복');
        return;
	} else if (pgid == "inicishp") {
		frm.mode.value = "getonpgdatahp";
        alert('중복');
        return;
	} else if (pgid == "uplus") {
		frm.mode.value = "getonpgdatauplus";
        alert('중복');
        return;
	} else if (pgid == "kakaopayT") {
		// 카카오PAY 거래내역
		frm.mode.value = "getonpgdatakakaopayT";
	} else if (pgid == "kakaopayS") {
		// 카카오PAY 정산내역
		frm.mode.value = "getonpgdatakakaopayS";
	} else if (pgid == "newkakaopayT") {
		// 카카오PAY 거래내역
		frm.mode.value = "getonpgdatanewkakaopayT";
        alert('중복');
        return;
	} else if (pgid == "newkakaopayS") {
		// 카카오PAY 정산내역
		frm.mode.value = "getonpgdatanewkakaopayS";
        alert('중복');
        return;
	} else if (pgid == "naverpay") {
		// 네이버페이
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
        // 토스는 거래내역에 정산일자 포함되어 있다.
		if (yyyymmdd.value.length == 10) {
			popGetTossDue(yyyymmdd.value);
		} else {
			popGetTossDue("");
		}
		return;
	} else if (pgid == "chaiT") {
		// 차이페이 정산 거래내역
		frm.mode.value = "getonpgdatachaipayT";
	} else if (pgid == "chaiS") {
		// 차이페이 정산 거래내역
		frm.mode.value = "getonpgdatachaipayS";
        alert('중복');
        return;
    } else if (pgid = 'appMethod6') {
        // 무통장입금
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
			if (confirm("PG데이타(ON " + pgid + ") 를 가져오기 하시겠습니까?\n\n중복주문번호(" + frmUplus.searchtext.value + ")") == true) {
				frm.orderserial.value = frmUplus.searchtext.value;
				frm.submit();
			}
		} else {
			if (confirm("PG데이타(ON " + pgid + ") 를 가져오기 하시겠습니까?") == true) {
				frm.submit();
			}
		}
	} else if (pgid == "newkakaopayT") {
		var frmUplus = document.frm;
		if ((frmUplus.searchfield.value == "orderserial") && (frmUplus.searchtext.value != "")) {
			if (confirm("PG데이타(ON " + pgid + ") 를 가져오기 하시겠습니까?\n\n주문번호(" + frmUplus.searchtext.value + ")") == true) {
				frm.orderserial.value = frmUplus.searchtext.value;
				frm.submit();
			}
		} else {
			if (confirm("PG데이타(ON " + pgid + ") 를 가져오기 하시겠습니까?") == true) {
				frm.submit();
			}
		}
	} else {
		if (confirm("PG데이타(ON " + pgid + ") 를 가져오기 하시겠습니까?") == true) {
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

	if (confirm("자동매칭(10x10) 하시겠습니까?") == true) {
		frm.submit();
	}
}

function jsMatchEtcPayment() {
	var frm = document.frmAct;

	frm.mode.value = "matchetcpay";

	if (confirm("자동매칭(무통장입금) 하시겠습니까?") == true) {
		frm.submit();
	}
}

<% if (searchfield = "PGkey") and (searchtext <> "") then %>
function jsMatchPGDataOld() {
	var frm = document.frmAct;

	frm.mode.value = "matchpgdata6month";
	frm.PGKey.value = "<%= searchtext %>";

	if (confirm("자동매칭(10x10,6개월이전) 하시겠습니까?") == true) {
		frm.submit();
	}
}
<% end if %>

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

	if (confirm("[취소]내역 매칭 하시겠습니까?") == true) {
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

	if (confirm("승인내역(0원) 을 추가 하시겠습니까?") == true) {
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
		alert("상세사유를 선택 하세요.");
		frm.selectreasonGubun.focus();
		return;
	}

	frm.mode.value = "RegReasonGubunarr";

	if (confirm("일괄사유 입력 하시겠습니까?") == true) {
		frm.action="/admin/maechul/pgdata/pgdata_process.asp";
		frm.submit();
	}
}
<% if (C_ADMIN_AUTH) then %>
function jsDelOne(idx) {
    var frm = document.frmAct;

	if (confirm("중복 주문건만 삭제 가능한 기능입니다.\n정말로 삭제 하시겠습니까?") == true) {
        frm.mode.value = "delapplog";
        frm.logidx.value = idx;
		frm.submit();
	}
}
<% end if %>
<% if (C_ADMIN_AUTH) then %>
function jsIniRentalCancel(pgkey) {
    var frm = document.frmAct;

	if (confirm("여기서 취소 처리를 하기전에 이니시스 어드민에서 취소 처리 해야 됩니다.\n취소처리 후 cs취소처리도 수동으로 해야 됩니다.\n취소처리 하시겠습니까?") == true) {
        frm.mode.value = "inirentalcancel";
        frm.PGKey.value = pgkey;
		frm.submit();
	}
}
<% end if %>
function jsRegReasonGubun025() {
	<% if (chkSearchAppDate = "Y") and (appMethod = "77") and (pggubun = "bankrefund") and (PGuserid = "bankrefund_10x10") then %>
	frm.mode.value = "RegReasonGubun025";

	if (confirm("선수금(예치금환급) 일괄입력 하시겠습니까?") == true) {
		frm.action="/admin/maechul/pgdata/pgdata_process.asp";
		frm.submit();
	}
	<% else %>
	alert('아래 조건으로 검색한 경우만 입력가능합니다.\n\n - 승인(취소)일자 체크\n - 결제방식 : 무통장환불\n - PG사 : bankrefund\n - PG사id : bankrefund_10x10\n - 상세사유 : 입력이전');
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
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;
		<input type="checkbox" name="chkSearchAppDate"  value="Y" <% if (chkSearchAppDate = "Y") then %>checked<% end if %> > 승인(취소)일자:
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<input type="checkbox" name="chkSearchIpkumDate"  value="Y" <% if (chkSearchIpkumDate = "Y") then %>checked<% end if %> > 입금예정일:
		<% DrawDateBoxdynamic yyyy3, "yyyy3", yyyy4, "yyyy4", mm3, "mm3", mm4, "mm4", dd3, "dd3", dd4, "dd4"  %>
		&nbsp;
		* 승인구분 :
		<select class="select" name="appDivCode">
		<option value=""></option>
		<option value="A" <% if (appDivCode = "A") then %>selected<% end if %> >승인</option>
		<option value="C" <% if (appDivCode = "C") then %>selected<% end if %> >취소</option>
		<option value="R" <% if (appDivCode = "R") then %>selected<% end if %> >부분취소</option>
		<option value="">----</option>
		<option value="E" <% if (appDivCode = "E") then %>selected<% end if %> >에러</option>
		</select>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		* 입금예정일 :
		<input type="text" class="text" name="ipkumdate" value="<%= ipkumdate %>" size="10">
		&nbsp;
		* 표시갯수 :
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
		* 결제방식 :
		<select class="select" name="appMethod">
			<option value=""></option>
			<option value="7" <% if (appMethod = "7") then %>selected<% end if %> >무통장(가상)</option>
			<option value="14" <% if (appMethod = "14") then %>selected<% end if %> >편의점결제</option>
			<option value="100" <% if (appMethod = "100") then %>selected<% end if %> >신용</option>
			<option value="20" <% if (appMethod = "20") then %>selected<% end if %> >실시간</option>
			<option value="80" <% if (appMethod = "80") then %>selected<% end if %> >All@</option>
			<option value="110" <% if (appMethod = "110") then %>selected<% end if %> >OK캐시백</option>
			<option value="400" <% if (appMethod = "400") then %>selected<% end if %> >핸드폰</option>
			<option value="550" <% if (appMethod = "550") then %>selected<% end if %> >기프팅</option>
			<option value="560" <% if (appMethod = "560") then %>selected<% end if %> >기프티콘</option>
            <option value="150" <% if (appMethod = "150") then %>selected<% end if %> >이니렌탈</option>
			<option value="">---------</option>
			<option value="77" <% if (appMethod = "77") then %>selected<% end if %> >무통장환불</option>
			<option value="6" <% if (appMethod = "6") then %>selected<% end if %> >무통장입금</option>
		</select>
		&nbsp;
		* PG사 :
		<select name="pggubun" class="select">
			<option value="">--선택--</option>
			<%Call sbGetOptPGgubun(pggubun)%>
		</select>
		<% 'Call DrawSelectBoxPGGubun("pggubun", pggubun, "") %>
		&nbsp;
		* PG사id :
		<select name="PGuserid" class="select">
			<option value="">--선택--</option>
			<%Call sbGetOptPGID(PGuserid)%>
		</select>
		<% 'Call DrawSelectBoxPGUserid("PGuserid", PGuserid, "") %>
		&nbsp;
		* 상세사유 :
		<select class="select" name="reasonGubun">
		<option value=""></option>
		<option value="001" <% if (reasonGubun = "001") then %>selected<% end if %> >선수금(매출)</option>
		<option value="002" <% if (reasonGubun = "002") then %>selected<% end if %> >선수금(제휴사 매출)</option>
        <option value="003" <% if (reasonGubun = "003") then %>selected<% end if %> >선수금(이니랜탈)</option>
		<option value="020" <% if (reasonGubun = "020") then %>selected<% end if %> >선수금(예치금)</option>
		<option value="025" <% if (reasonGubun = "025") then %>selected<% end if %> >선수금(예치금환급)</option>
		<option value="030" <% if (reasonGubun = "030") then %>selected<% end if %> >선수금(기프트)</option>
		<option value="035" <% if (reasonGubun = "035") then %>selected<% end if %> >선수금(기프트환급)</option>
        <option value="004" <% if (reasonGubun = "004") then %>selected<% end if %> >선수금(B2B 매출)</option>
		<option value="">---------------</option>
		<option value="040" <% if (reasonGubun = "040") then %>selected<% end if %> >CS서비스</option>
		<option value="">---------------</option>
		<option value="950" <% if (reasonGubun = "950") then %>selected<% end if %> >무통장미확인</option>
		<option value="999" <% if (reasonGubun = "999") then %>selected<% end if %> >취소매칭</option>
		<option value="901" <% if (reasonGubun = "901") then %>selected<% end if %> >핑거스현금매출</option>
		<option value="800" <% if (reasonGubun = "800") then %>selected<% end if %> >이자수익</option>
		<option value="900" <% if (reasonGubun = "900") then %>selected<% end if %> >기타</option>
		<option value="">---------------</option>
		<option value="XXX" <% if (reasonGubun = "XXX") then %>selected<% end if %> >입력이전</option>
		</select>
		&nbsp;
		* 사이트 :
		<select class="select" name="sitename">
		<option value=""></option>
		<option value="10x10" <% if (sitename = "10x10") then %>selected<% end if %> >10x10(PC)</option>
		<option value="10x10mobile" <% if (sitename = "10x10mobile") then %>selected<% end if %> >10x10(MOBILE)</option>
		<option value="fingers" <% if (sitename = "fingers") then %>selected<% end if %> >핑거스</option>
		<option value="10x10gift" <% if (sitename = "10x10gift") then %>selected<% end if %> >10x10(GIFT)</option>
        <option value="wholesale" <% if (sitename = "wholesale") then %>selected<% end if %> >WHOLESALE</option>
		</select>
		&nbsp;
		* 검색조건 :
		<select class="select" name="searchfield">
		<option value=""></option>
		<option value="PGkey" <% if (searchfield = "PGkey") then %>selected<% end if %> >PG사KEY</option>
		<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >주문번호</option>
		<option value="appPrice" <% if (searchfield = "appPrice") then %>selected<% end if %> >거래금액</option>
		</select>
		<input type="text" class="text" name="searchtext" value="<%= searchtext %>" size="50">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		<input type="checkbox" name="excmatchfinish"  value="Y" <% if (excmatchfinish = "Y") then %>checked<% end if %> > 매칭완료(승인건 주문번호매칭, 취소건 CS내역매칭) 제외
		&nbsp;
		<input type="checkbox" name="onlypricenotequal"  value="Y" <% if (onlypricenotequal = "Y") then %>checked<% end if %> > 원주문 승인금액 상이내역만(30분 지연정보)
		&nbsp;
		<input type="checkbox" name="showjumunlog"  value="Y" <% if (showjumunlog = "Y") then %>checked<% end if %> > 결제로그 표시(30분 지연정보)
		&nbsp;
		<input type="checkbox" name="showjumunlogNotMatch"  value="Y" <% if (showjumunlogNotMatch = "Y") then %>checked<% end if %> > <b>결제로그 매칭완료 제외(30분 지연정보, 결제월 다른경우 표시)</b>
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<h5>테스트중...</h5>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;" border="0">
<tr>
	<td align="left" width="50%">
		- 통신오류로 인한 주문입력 실패 승인 및 취소 매칭<br>
		* 실시간이체(결제일 이후 취소)의 경우 수수료 금액을 취소하지 않는다.<br>
		* 이니시스 주말 승인내역의 입금예정일은 다음주 금요일이다.<br>
		* <font color="red">합계입금</font>은 우선 1개의 주문번호가 매칭된 이후에 추가로 입력가능하다.<br />
		* 네이버페이 실시간이체를 취소한 경우, 승인일자 승인내역을 다시 다운받아야 한다.(http://wapi.10x10.co.kr/nPay/jungsanReceive.asp)<br />
		* UPLUS PK 중복 오류 있는 경우, <font color="red">해당 주문번호 검색 후 승인내역 가져오기</font> 누르면 가져오기 됩니다.
	</td>
	<td align="left">
		* <font color="red"><b>상세사유 자동입력</b></font> : 승인자료 다운로드 -&gt; 주문번호매칭 -&gt; 결제로그매칭(30분경과필요) -&gt; 미수금재작성
	</td>
</tr>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;" border="0">
<tr>
	<td align="left">
        <!--
		<input type="button" class="button" value="가져오기(ON INICIS)" onClick="jsGetOnPGData('inicis');" disabled>
		<input type="button" class="button" value="가져오기(ON INICIS HP)" onClick="jsGetOnPGData('inicishp');" disabled>
		<input type="button" class="button" value="가져오기(ON UPLUS)" onClick="jsGetOnPGData('uplus');" disabled>
        -->
        <!--
		<input type="button" class="button" value="등록하기(ON 네이버페이정산)" onClick="popUploadNAVERPAYPGData();">
        -->
        <input type="button" class="button" value="가져오기(ON 무통장)" onClick="jsGetOnPGData('appMethod6');">
		<br><br>
        <!--
		<input type="button" class="button" value="가져오기(ON NewKAKAO 거래)" onClick="jsGetOnPGData('newkakaopayT');" disabled>
		<input type="button" class="button" value="가져오기(ON NewKAKAO 정산)" onClick="jsGetOnPGData('newkakaopayS');" disabled>
        -->
        <!--
		<input type="button" class="button" value="가져오기(페이코 거래)" onClick="jsGetOnPGData('paycoT');">
		<input type="button" class="button" value="가져오기(페이코 정산)" onClick="jsGetOnPGData('paycoS');">
        -->
        <!--
		<input type="button" class="button" value="가져오기(토스)" onClick="jsGetOnPGData('toss');">
        -->
        <!--
		<input type="button" class="button" value="가져오기(차이 거래)" onClick="jsGetOnPGData('chaiT');">
        -->
        <!--
		<input type="button" class="button" value="가져오기(차이 정산)" onClick="jsGetOnPGData('chaiS');" disabled>
        -->
		<input type="text" class="text" id="yyyymmdd" name="yyyymmdd" value="" size="12">
		<!--
		<br><br>
		<input type="button" class="button" value="가져오기(ON 네이버페이)" onClick="jsGetOnPGData111('naverpay');">
		-->
		<br><br>
		<input type="button" class="button" value="등록하기(기프티콘)" onClick="popUploadONPGData('gifticon');">
		<input type="button" class="button" value="등록하기(기프팅)" onClick="popUploadONPGData('giftting');">
		<input type="button" class="button" value="등록하기(수기)" onClick="popUploadHandData();">
		<% If session("ssBctId") = "thensi7" Then %>
			<input type="button" class="button" value="이니렌탈 수기등록(수기)" onClick="popUploadIniRentalData();">
		<% End If %>
	</td>
	<td align="right">
		<input type="button" class="button" value="자동매칭(무통장입금)" onClick="jsMatchEtcPayment();">
        <input type="button" class="button" value="자동매칭(10x10)" onClick="jsMatchPGData();">
		<input type="button" class="button" value="자동매칭(핑거스)" onClick="jsMatchFingersPGData();">
		<input type="button" class="button" value="자동매칭(기프트)" onClick="jsMatchGiftCardPGData();">
		<br /><br />
		<% if (searchfield = "PGkey") and (searchtext <> "") then %>
		<input type="button" class="button" value="자동매칭(10x10,6개월이전,<%= searchtext %>)" onClick="jsMatchPGDataOld();">
		<% end if %>

		<% if PGuserid <> "" then %>
		<br>
		<input type="button" class="button" value="선수금(예치금환급) 일괄입력" onClick="jsRegReasonGubun025();" style="width:180px;"> &nbsp;
			* 상세사유 :
			<select class="select" name="selectreasonGubun">
			<option value=""></option>
			<option value="001" <% if (reasonGubun = "001") then %>selected<% end if %> >선수금(매출)</option>
			<option value="002" <% if (reasonGubun = "002") then %>selected<% end if %> >선수금(제휴사 매출)</option>
			<option value="020" <% if (reasonGubun = "020") then %>selected<% end if %> >선수금(예치금)</option>
			<option value="025" <% if (reasonGubun = "025") then %>selected<% end if %> >선수금(예치금환급)</option>
			<option value="030" <% if (reasonGubun = "030") then %>selected<% end if %> >선수금(기프트)</option>
			<option value="035" <% if (reasonGubun = "035") then %>selected<% end if %> >선수금(기프트환급)</option>
			<option value="">---------------</option>
			<option value="040" <% if (reasonGubun = "040") then %>selected<% end if %> >CS서비스</option>
			<option value="">---------------</option>
			<option value="950" <% if (reasonGubun = "950") then %>selected<% end if %> >무통장미확인</option>
			<option value="999" <% if (reasonGubun = "999") then %>selected<% end if %> >취소매칭</option>
			<option value="901" <% if (reasonGubun = "901") then %>selected<% end if %> >핑거스현금매출</option>
			<option value="800" <% if (reasonGubun = "800") then %>selected<% end if %> >이자수익</option>
			<option value="900" <% if (reasonGubun = "900") then %>selected<% end if %> >기타</option>
			<option value="">---------------</option>
			<option value="XXX" <% if (reasonGubun = "XXX") then %>selected<% end if %> >입력이전</option>
			</select>
			<input type="button" class="button" value="사유일괄입력" onClick="jsRegReasonGubunarr();" style="width:100px;">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->
</form>
<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21">
		검색결과 : <b><%= oCPGData.FTotalcount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCPGData.FTotalPage %></b>
        &nbsp;
        거래총액 : <b><%= FormatNumber(oCPGData.FTotalAppPrice, 0) %> 원</b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>PG사</td>
	<td>PG사id</td>
	<td width="80">결제방식</td>
	<td>PG사KEY</td>
	<td>PG사CSKEY</td>
	<td width="60">구분</td>
	<td width="150">승인(취소)일자</td>
	<td width="60">거래액</td>
	<td width="60">수수료<br>(VAT포함)</td>
	<td width="60">입금<br>예정액</td>
	<td width="65">카드사<br>매입일</td>
	<td width="70">입금예정일</td>
	<td>사이트</td>
	<td>주문번호</td>
	<td width="60">CSIDX</td>
	<% if (showjumunlog = "Y") then %>
	<td>결제로그</td>
	<% end if %>
	<td>상세사유</td>
	<!--
	<td width="80">등록일</td>
	-->
	<td>비고</td>
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
        <input type="button" class="button" value="입력" onClick="jsPopInputOrderSerial(<%= oCPGData.FItemList(i).Fidx %>)">
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
			<input type="button" class="button" value="취소매칭" onClick="jsMatchCancel(<%= oCPGData.FItemList(i).Fidx %>, false);">
			<% if (searchfield = "PGkey") and (searchtext <> "") then %>
			<input type="button" class="button" value="취소매칭(다른날짜)" onClick="jsMatchCancel(<%= oCPGData.FItemList(i).Fidx %>, true);">
			<% end if %>
		<% elseif Not IsNull(oCPGData.FItemList(i).Forderserial) and (oCPGData.FItemList(i).FappDivCode = "C") and IsNull(oCPGData.FItemList(i).Fcsasid) then %>
			<input type="button" class="button" value="중복승인취소매칭" onClick="jsDuplicateMatchCancel(<%= oCPGData.FItemList(i).Fidx %>);">
		<% end if %>
		<% if (oCPGData.FItemList(i).FPGgubun = "bankipkum") and (oCPGData.FItemList(i).FappDivCode <> "C") and (oCPGData.FItemList(i).FappPrice >= 1000) and (oCPGData.FItemList(i).Forderserial <> "") then %>
			<input type="button" class="button" value="합계입금" onClick="jsPopSumIpkum(<%= oCPGData.FItemList(i).Fidx %>)">
		<% end if %>
		<% if (oCPGData.FItemList(i).FPGgubun = "bankrefund") and (oCPGData.FItemList(i).FappDivCode <> "A") and (oCPGData.FItemList(i).FappPrice <> 0) then %>
			<input type="button" class="button" value="내역추가(0원)" onClick="jsAddRefundLog(<%= oCPGData.FItemList(i).Fidx %>)">
		<% end if %>
		<% if IsNull(oCPGData.FItemList(i).FreasonGubun) or Not (InStr("001,020,030,950", oCPGData.FItemList(i).FreasonGubun) > 0) or C_ADMIN_AUTH or C_MngPart or C_PSMngPart then %>
			<input type="button" class="button" value="사유" onClick="jsPopRegReasonGubun(<%= oCPGData.FItemList(i).Fidx %>)">
			<%' 원래는 관리자만 삭제 되게 했던건데 로그를 남기므로 권한 풀어줌 %>
			<input type="button" class="button" value="삭제" onClick="jsDelOne(<%=oCPGData.FItemList(i).Fidx %>)">
			<%' if (C_ADMIN_AUTH or C_MngPart or C_PSMngPart) then %><!--[관리자]
            <input type="button" class="button" value="삭제" onClick="jsDelOne(<%'oCPGData.FItemList(i).Fidx %>)">-->
			<%' end if %>
			<%' 이니렌탈 전용 취소%>
			<% If (C_ADMIN_AUTH) Then %>
				<% If Trim(oCPGData.FItemList(i).FPGuserid) = "teenxteenr" Then %>
					<% If Trim(oCPGData.FItemList(i).GetAppDivCodeName) = "승인" Then %>
						<input type="button" class="button" value="취소" onClick="jsIniRentalCancel('<%=oCPGData.FItemList(i).FPGkey %>')">
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
