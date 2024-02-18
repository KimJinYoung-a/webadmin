<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : PG사 승인통계_ON
' Hieditor : 2011.04.22 이상구 생성
'			 2020.06.09 정태훈 수정(네이버포인트 추가, 차이페이 추가)
'			 2020.03.28 허진원 수정(삼성페이 추가)
'			 2023.05.31 한용민 수정(편의점결제 정산 승인내역 업로드 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/maechul/incMaechulFunction.asp"-->
<%
dim research, page
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim yyyy3,yyyy4,mm3,mm4,dd3,dd4
dim yyyy, mm, dd
dim fromDate ,toDate, tmpDate
dim fromDate2 ,toDate2
dim PGuserid, sitename, dategubun
dim appDivCode, cardReaderID, cardGubun, cardComp, cardAffiliateNo, ipkumdate
dim searchfield, searchtext
dim pggubun
dim chkSearchIpkumDate, chkSearchAppDate
dim sumgubun
dim reasonGubun
Dim i, j

	research = requestCheckvar(request("research"),10)
	page = requestCheckvar(request("page"),10)

	yyyy    = request("yyyy")
	mm      = request("mm")
	dd      = request("dd")

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
	dd4     = request("dd4")

	PGuserid 		= request("PGuserid")
	sitename 		= request("sitename")
	appDivCode 		= request("appDivCode")
	cardReaderID 	= request("cardReaderID")
	cardGubun 		= request("cardGubun")
	cardComp 		= request("cardComp")
	cardAffiliateNo = request("cardAffiliateNo")
	ipkumdate 		= request("ipkumdate")
	reasonGubun 	= request("reasonGubun")

	searchfield 	= request("searchfield")
	searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")

	dategubun 		= request("dategubun")

	pggubun 		= request("pggubun")

	chkSearchIpkumDate 	= request("chkSearchIpkumDate")
	chkSearchAppDate 	= request("chkSearchAppDate")

	sumgubun 	= request("sumgubun")

if (chkSearchIpkumDate="") then chkSearchAppDate = "Y"
if (page="") then page = 1

if (research="") then
	dategubun = "appdate"
end if

if (sumgubun = "") then
	sumgubun = "appMethod"
end if

if (yyyy="") then
	yyyy = Cstr(Year(Now()))
	mm = Cstr(Month(Now()))
	dd = Cstr(day(Now()))
end if

if (yyyy1="") then
	''fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 0), 1)  ''당월로 수정 //2016/03/31 by eastone
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) + 1), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))

	fromDate2 = fromDate
	toDate2 = toDate
	yyyy3 = yyyy1
	mm3 = mm1
	dd3 = dd1
	yyyy4 = yyyy2
	mm4 = mm2
	dd4 = dd2
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
	fromDate2 = DateSerial(yyyy3, mm3, dd3)
	toDate2 = DateSerial(yyyy4, mm4, dd4+1)
end if


Dim oCPGDataStatistics
set oCPGDataStatistics = new CPGData

	oCPGDataStatistics.FRectPGuserid = PGuserid
	oCPGDataStatistics.FRectSiteName = sitename

	oCPGDataStatistics.FRectDateGubun = dategubun
	oCPGDataStatistics.FRectReasonGubun = reasonGubun

	if (chkSearchAppDate = "Y") then
		oCPGDataStatistics.FRectStartdate = fromDate
		oCPGDataStatistics.FRectEndDate = toDate
		response.write "bbbb"
	end if

	if (chkSearchIpkumDate = "Y") then
		oCPGDataStatistics.FRectStartIpkumdate = fromDate2
		oCPGDataStatistics.FRectEndIpkumDate = toDate2
		response.write "aaaa"
	end if

	oCPGDataStatistics.FRectPGGubun = pggubun

	''oCPGDataStatistics.FRectSumGubun = sumgubun

    oCPGDataStatistics.getPGDataStatisticList_ON

dim totSumCardPrice, totSumBankPrice, totSumVBankPrice, totSumHPPrice, totSumPrice
dim totSumCardJungsanPrice, totSumBankJungsanPrice, totSumVBankJungsanPrice, totSumHPJungsanPrice, totSumJungsanPrice
dim totSumGifttingPrice, totSumGifticonPrice, totSumOKPrice, totSumAllAtPrice
dim totSumGifttingJungsanPrice, totSumGifticonJungsanPrice, totSumOKJungsanPrice, totSumAllAtJungsanPrice
dim totSumTenOutBankPrice, totSumTenInBankPrice
dim totSumTenOutBankJungsanPrice, totSumTenInBankJungsanPrice
dim totSumteenxteen3Price, totSumteenxteen4Price, totSumteenxteen5Price, totSumteenxteen6Price, totSumteenxteen8Price, totSumteenxteen9Price
dim totSumteenteen10Price, totSumtenbyten01Price, totSumtenbyten02Price, totSumteenxteehaPrice, totSumteenxteenrPrice, totSumteenteenspPrice
dim totSumteenteenapPrice, totSumKCTEN0001mPrice, totSumnaverpayPrice, totSumpaycoPrice, totSumbankipkumPrice, totSumbankipkum_10x10Price
dim totSumbankipkum_fingersPrice, totSumbankrefundPrice, totSumbankrefund_10x10Price, totSumbankrefund_fingersPrice, totSum10x10_2Price
dim totSumR5523Price, totSummobiliansPrice, totSumPGgifticonPrice, totSumPGgifttingPrice, totSumPGokcashbagPrice, totSumPGtossPrice, totSumPGchaiPrice
dim totSumnaverpayPoint
dim totSumteenxteen3JungsanPrice, totSumteenxteen4JungsanPrice, totSumteenxteen5JungsanPrice, totSumteenxteen6JungsanPrice, totSumteenxteen8JungsanPrice
dim totSumteenxteen9JungsanPrice, totSumteenteen10JungsanPrice, totSumtenbyten01JungsanPrice, totSumtenbyten02JungsanPrice, totSumteenxteehaJungsanPrice
dim totSumteenxteenrJungsanPrice, totSumteenteenspJungsanPrice, totSumteenteenapJungsanPrice, totSumKCTEN0001mJungsanPrice, totSumnaverpayJungsanPrice
dim totSumpaycoJungsanPrice, totSumbankipkumJungsanPrice, totSumbankipkum_10x10JungsanPrice, totSumbankipkum_fingersJungsanPrice, totSumbankrefundJungsanPrice
dim totSumbankrefund_10x10JungsanPrice, totSumbankrefund_fingersJungsanPrice, totSum10x10_2JungsanPrice, totSumR5523JungsanPrice, totSummobiliansJungsanPrice
dim totSumPGgifticonJungsanPrice, totSumPGgifttingJungsanPrice, totSumPGokcashbagJungsanPrice, totSumPGtossJungsanPrice
dim totsumKakaoJungsanPrice, totSumKakaopayPrice, totSumPGchaiJungsanPrice, totSumnaverpayJungsanPoint, totSumPGConvinienspayPrice, totSumConvinienspayJungsanPrice
%>

<script src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popPGDataList(yyyy1, mm1, dd1, PGuserid) {
	var popup = window.open("pgdata_on.asp?menupos=1567&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy1+"&mm2="+mm1+"&dd2="+dd1+"&PGuserid="+PGuserid,"popPGDataList","width=1024,height=768,scrollbars=yes,resizable=yes");
	popup.focus();
}

function jsReloadOnPGData(pgid, appdate) {
	var frm = document.frmAct;

	if (pgid == "inicis") {
		frm.mode.value = "getonpgdata";
		frm.yyyymmdd.value = appdate;
	} else if (pgid == "inicishp") {
		frm.mode.value = "getonpgdatahp";
		frm.yyyymmdd.value = appdate;
	} else if (pgid == "inicishppre") {
		frm.mode.value = "getonpgdatahppre";
		frm.yyyymmdd.value = appdate;
	} else if (pgid == "uplus") {
		frm.mode.value = "getonpgdatauplus";
		frm.yyyymmdd.value = appdate;
	} else if (pgid == "kakaopayT") {
		frm.mode.value = "getonpgdatanewkakaopayT";
		frm.yyyymmdd.value = appdate;
	} else if (pgid == "kakaopayS") {
		frm.mode.value = "getonpgdatanewkakaopayS";
		frm.yyyymmdd.value = appdate;
	} else if (pgid == "paycoT") {
		frm.mode.value = "getpaycoT";
		frm.yyyymmdd.value = appdate;
	} else if (pgid == "paycoS") {
		frm.mode.value = "getpaycoS";
		frm.yyyymmdd.value = appdate;
	} else if (pgid == "naverpayT") {
        jsCallNaverPay(appdate);
        return;
		// var popup = window.open("http://wapi.10x10.co.kr/nPay/jungsanReceive.asp?sDate=" + appdate + "&eDate=" + appdate,"popPGDataList","width=1024,height=768,scrollbars=yes,resizable=yes");
		// popup.focus();
	} else if (pgid == "chaiT") {
		frm.mode.value = "getonpgdatachaipayT";
		frm.yyyymmdd.value = appdate;
	} else if (pgid == "chai") {
		frm.mode.value = "getonpgdatachaipayS";
		frm.yyyymmdd.value = appdate;
	} else {
		alert("ERROR");
		return;
	}

	if (confirm("PG데이타(ON " + pgid + ") : " + appdate + "  를 새로고침(다시 가져오기) 하시겠습니까?") == true) {
		frm.submit();
	}
}

function jsReloadOnPGData2(pgid) {
	var frm = document.frmOneDate;

	//2016/08/22
	if (frm.mm.value.length<2) frm.mm.value='0'+frm.mm.value;
	if (frm.dd.value.length<2) frm.dd.value='0'+frm.dd.value;

	var yyyymmdd = frm.yyyy.value + "-" + frm.mm.value + "-" + frm.dd.value;

    if (frm.chkdate.checked != true) {
        yyyymmdd = '';
    }

	jsReloadOnPGData(pgid, yyyymmdd);
}

function popUploadNAVERPAYPGData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("popRegNAVERPAYPGDataFile_on.asp","popUploadNAVERPAYPGData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}
function popUploadconvinienspayPGData() {
    var window_width = 800;
    var window_height = 500;

    var popwin = window.open("/admin/maechul/pgdata/popRegconvinienspayDataFile_on.asp","popUploadconvinienspayPGData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popGetToss() {
	var frm = document.frmOneDate;

	//2016/08/22
	if (frm.mm.value.length<2) frm.mm.value='0'+frm.mm.value;
	if (frm.dd.value.length<2) frm.dd.value='0'+frm.dd.value;

	var yyyymmdd = frm.yyyy.value + "-" + frm.mm.value + "-" + frm.dd.value;

    if (frm.chkdate.checked != true) {
        yyyymmdd = '';
    }

	var url = "http://wapi.10x10.co.kr/toss/api.asp?mode=settle&yyyymmdd=" + yyyymmdd;
	var popwin = window.open(url,"popGetToss","width=500 height=300 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsRealCall(appdate, page) {
    var url;
    var data = '{}';

    url = window.location.protocol + "//wapi.10x10.co.kr/nPay/jungsanReceive.asp?sDate=" + appdate + "&eDate=" + appdate + "&page=" + page;
    console.log(url);

    $.ajax({
        type : 'GET',
        url : url,
        data : data,
        async : false,
        timeout : 100000,
        dataType: 'html',
        contentType: 'application/x-www-form-urlencoded; charset=utf-8',
        error:function(request, status, error) {
            alert("code:"+request.status+"\n"+"message:"+request.responseText+"\n"+"error:"+error);
        },
        success : function(data) {
            if (data.indexOf('S_OK') == -1) {
                alert("finished");
                return;
            }

            if (page*1 >= 50) { return; }

            jsRealCall(appdate, page*1 + 1);
        }
    });

}

function jsCallNaverPay(appdate) {
    var url;
    var data = '{}';

	if (confirm("PG데이타(ON 네이버페이) : " + appdate + "  를 새로고침(다시 가져오기) 하시겠습니까?\n\n3~10 분정도의 시간이 소요됩니다.") != true) {
		return;
	}

    jsRealCall(appdate, 1);
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;
		*합계구분:
		<select class="select" name="sumgubun">
			<option value="appMethod" <% if (sumgubun = "appMethod") then %>selected<% end if %> >결제수단별</option>
			<option value="PGuserid" <% if (sumgubun = "PGuserid") then %>selected<% end if %> >PG사 아이디별</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
		*기준일자:
		<select class="select" name="dategubun">
			<option value="appdate" <% if (dategubun = "appdate") then %>selected<% end if %> >승인(취소)일</option>
			<option value="ipkumdate" <% if (dategubun = "ipkumdate") then %>selected<% end if %> >입금예정일</option>
		</select>
		&nbsp;
		<input type="checkbox" name="chkSearchAppDate"  value="Y" <% if (chkSearchAppDate = "Y") then %>checked<% end if %> > *승인(취소)일자:
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<input type="checkbox" name="chkSearchIpkumDate"  value="Y" <% if (chkSearchIpkumDate = "Y") then %>checked<% end if %> > *입금예정일:
		<% DrawDateBoxdynamic yyyy3, "yyyy3", yyyy4, "yyyy4", mm3, "mm3", mm4, "mm4", dd3, "dd3", dd4, "dd4"  %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
		* PG사 :
		<select name="pggubun" class="select">
			<option value="">--선택--</option>
			<%Call sbGetOptPGgubun(pggubun)%>
		</select>
		&nbsp;
		* PG사id :
		<select name="PGuserid" class="select">
			<option value="">--선택--</option>
			<%Call sbGetOptPGID(PGuserid)%>
		</select>
		&nbsp;
		* 상세사유 :
		<select class="select" name="reasonGubun">
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
		&nbsp;
		* 사이트 :
		<select class="select" name="sitename">
		<option value=""></option>
		<option value="10x10all" <% if (sitename = "10x10all") then %>selected<% end if %> >10x10</option>
		<option value="10x10" <% if (sitename = "10x10") then %>selected<% end if %> >10x10(PC)</option>
		<option value="10x10mobile" <% if (sitename = "10x10mobile") then %>selected<% end if %> >10x10(MOBILE)</option>
		<option value="fingers" <% if (sitename = "fingers") then %>selected<% end if %> >핑거스</option>
		<option value="10x10gift" <% if (sitename = "10x10gift") then %>selected<% end if %> >10x10(GIFT)</option>
		</select>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<form name="frmOneDate" method="get" action="" style="margin:0px;">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* 이니시스(INICIS) 내역은 영업일 기준 5일뒤가 입금예정일이고, 입금일 기준으로 내역을 가져옵니다.<br />
		* 네이버페이 오류난 경우, 해당날짜(승인일자)를 다시 돌려준다.
		* 네이버페이 취소 데이터가 없거나 오류난 경우, 원승인일자를 다시 돌려준다.
	</td>
</tr>
<tr>
	<td align="right">
        <input type="checkbox" name="chkdate" value="Y">
		<% Call DrawOneDateBoxdynamic("yyyy", yyyy, "mm", mm, "dd", dd, "", "", "", "") %>
		<input type="button" class="button" value="새로고침(INICIS, 입금일자)" onClick="jsReloadOnPGData2('inicis');">
		<input type="button" class="button" value="새로고침(INICIS HP 01)" onClick="jsReloadOnPGData2('inicishppre');">
		<input type="button" class="button" value="새로고침(INICIS HP 02)" onClick="jsReloadOnPGData2('inicishp');">
        &nbsp;&nbsp;&nbsp;
		<input type="button" class="button" value="새로고침(네이버페이 거래)" onClick="jsReloadOnPGData2('naverpayT');">
        <input type="button" class="button" value="업로드(네이버페이 정산)" onClick="popUploadNAVERPAYPGData();">
        <br /><br />
		<input type="button" class="button" value="새로고침(KAKAO 거래)" onClick="jsReloadOnPGData2('kakaopayT');">
		<input type="button" class="button" value="새로고침(KAKAO 정산)" onClick="jsReloadOnPGData2('kakaopayS');">
        &nbsp;&nbsp;&nbsp;
        <input type="button" class="button" value="새로고침(토스)" onClick="popGetToss();">
		<input type="button" class="button" value="새로고침(tosspayments[구uplus])" onClick="jsReloadOnPGData2('uplus');">
		<br /><br />
        <input type="button" class="button" value="새로고침(페이코 거래)" onClick="jsReloadOnPGData2('paycoT');">
		<input type="button" class="button" value="새로고침(페이코 정산)" onClick="jsReloadOnPGData2('paycoS');">
        &nbsp;&nbsp;&nbsp;
		<input type="button" class="button" value="업로드(편의점결제 정산)" onClick="popUploadconvinienspayPGData();">
        <%' <input type="button" class="button" value="새로고침(CHAI 거래)" onClick="jsReloadOnPGData2('chaiT');"> %>
		<%' <input type="button" class="button" value="새로고침(CHAI 정산)" onClick="jsReloadOnPGData2('chai');"> %>
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80" rowspan="2">
		<% if (dategubun = "ipkumdate") then %>
			입금예정일
		<% else %>
			승인(취소)일
		<% end if %>
	</td>
	<td colspan="<% if (sumgubun = "appMethod") then %>11<% else %>34<% end if %>">승인(취소)액</td>
	<td colspan="<% if (sumgubun = "appMethod") then %>11<% else %>34<% end if %>">입금예정액</td>
	<td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (sumgubun = "appMethod") then %>
	<td width="90">신용카드</td>
	<td width="90">실시간이체</td>
	<td width="90">가상계좌</td>
	<td width="90">무통장환불</td>
	<td width="90">무통장입금</td>
	<td width="90">핸드폰</td>
	<td width="90">기프팅</td>
	<td width="90">기프티콘</td>
	<td width="90">OK캐시백</td>
	<td width="90">All@</td>
	<% else %>
	<td width="90">teenxteen3</td>
	<td width="90">teenxteen4</td>
	<td width="90">teenxteen5</td>
	<td width="90">teenxteen6</td>
	<td width="90">teenxteen8</td>
	<td width="90">teenxteen9</td>
	<td width="90">teenteen10</td>
	<td width="90">tenbyten01</td>
	<td width="90">tenbyten02</td>
	<td width="90">teenxteeha</td>
	<td width="90">teenxteenr</td>
	<td width="90">teenteensp</td>
	<td width="90">teenteenap</td>
	<td width="90">KCTEN0001m</td>
	<td width="90">newkakaopay</td>
	<td width="90">naverpay</td>
	<td width="90">naverPoint</td>
	<td width="90">payco</td>
	<td width="90">bankipkum</td>
	<td width="90">bankipkum_10x10</td>
	<td width="90">bankipkum_fingers</td>
	<td width="90">bankrefund</td>
	<td width="90">bankrefund_10x10</td>
	<td width="90">bankrefund_fingers</td>
	<td width="90">10x10_2</td>
	<td width="90">R5523</td>
	<td width="90">mobilians</td>
	<td width="90">gifticon</td>
	<td width="90">giftting</td>
	<td width="90">okcashbag</td>
	<td width="90">toss</td>
	<td width="90">chai</td>
	<td width="90">convinienspay</td>
	<% end if %>
	<td width="100">합계</td>

	<% if (sumgubun = "appMethod") then %>
	<td width="90">신용카드</td>
	<td width="90">실시간이체</td>
	<td width="90">가상계좌</td>
	<td width="90">무통장환불</td>
	<td width="90">무통장입금</td>
	<td width="90">핸드폰</td>
	<td width="90">기프팅</td>
	<td width="90">기프티콘</td>
	<td width="90">OK캐시백</td>
	<td width="90">All@</td>
	<% else %>
	<td width="90">teenxteen3</td>
	<td width="90">teenxteen4</td>
	<td width="90">teenxteen5</td>
	<td width="90">teenxteen6</td>
	<td width="90">teenxteen8</td>
	<td width="90">teenxteen9</td>
	<td width="90">teenteen10</td>
	<td width="90">tenbyten01</td>
	<td width="90">tenbyten02</td>
	<td width="90">teenxteeha</td>
	<td width="90">teenxteenr</td>
	<td width="90">teenteensp</td>
	<td width="90">teenteenap</td>
	<td width="90">KCTEN0001m</td>
	<td width="90">newkakaopay</td>
	<td width="90">naverpay</td>
	<td width="90">naverPoint</td>
	<td width="90">payco</td>
	<td width="90">bankipkum</td>
	<td width="90">bankipkum_10x10</td>
	<td width="90">bankipkum_fingers</td>
	<td width="90">bankrefund</td>
	<td width="90">bankrefund_10x10</td>
	<td width="90">bankrefund_fingers</td>
	<td width="90">10x10_2</td>
	<td width="90">R5523</td>
	<td width="90">mobilians</td>
	<td width="90">gifticon</td>
	<td width="90">giftting</td>
	<td width="90">okcashbag</td>
	<td width="90">toss</td>
	<td width="90">chai</td>
	<td width="90">convinienspay</td>
	<% end if %>

	<td width="100">합계</td>
</tr>

<% for i=0 to oCPGDataStatistics.FresultCount -1 %>
<%

totSumCardPrice = totSumCardPrice + oCPGDataStatistics.FItemList(i).FsumCardPrice
totSumBankPrice = totSumBankPrice + oCPGDataStatistics.FItemList(i).FsumBankPrice
totSumVBankPrice = totSumVBankPrice + oCPGDataStatistics.FItemList(i).FsumVBankPrice
totSumTenOutBankPrice = totSumTenOutBankPrice + oCPGDataStatistics.FItemList(i).FsumTenOutBankPrice
totSumTenInBankPrice = totSumTenInBankPrice + oCPGDataStatistics.FItemList(i).FsumTenInBankPrice
totSumHPPrice = totSumHPPrice + oCPGDataStatistics.FItemList(i).FsumHPPrice
totSumGifttingPrice = totSumGifttingPrice + oCPGDataStatistics.FItemList(i).FsumGifttingPrice
totSumGifticonPrice = totSumGifticonPrice + oCPGDataStatistics.FItemList(i).FsumGifticonPrice
totSumOKPrice = totSumOKPrice + oCPGDataStatistics.FItemList(i).FsumOKPrice
totSumAllAtPrice = totSumAllAtPrice + oCPGDataStatistics.FItemList(i).FsumAllAtPrice

totSumteenxteen3Price = totSumteenxteen3Price + oCPGDataStatistics.FItemList(i).Fsumteenxteen3Price
totSumteenxteen4Price = totSumteenxteen4Price + oCPGDataStatistics.FItemList(i).Fsumteenxteen4Price
totSumteenxteen5Price = totSumteenxteen5Price + oCPGDataStatistics.FItemList(i).Fsumteenxteen5Price
totSumteenxteen6Price = totSumteenxteen6Price + oCPGDataStatistics.FItemList(i).Fsumteenxteen6Price
totSumteenxteen8Price = totSumteenxteen8Price + oCPGDataStatistics.FItemList(i).Fsumteenxteen8Price
totSumteenxteen9Price = totSumteenxteen9Price + oCPGDataStatistics.FItemList(i).Fsumteenxteen9Price
totSumteenteen10Price = totSumteenteen10Price + oCPGDataStatistics.FItemList(i).Fsumteenteen10Price
totSumtenbyten01Price = totSumtenbyten01Price + oCPGDataStatistics.FItemList(i).Fsumtenbyten01Price
totSumtenbyten02Price = totSumtenbyten02Price + oCPGDataStatistics.FItemList(i).Fsumtenbyten02Price
totSumteenxteehaPrice = totSumteenxteehaPrice + oCPGDataStatistics.FItemList(i).FsumteenxteehaPrice
totSumteenxteenrPrice = totSumteenxteenrPrice + oCPGDataStatistics.FItemList(i).FsumteenxteenrPrice
totSumteenteenspPrice = totSumteenteenspPrice + oCPGDataStatistics.FItemList(i).FsumteenteenspPrice
totSumteenteenapPrice = totSumteenteenapPrice + oCPGDataStatistics.FItemList(i).FsumteenteenapPrice
totSumKCTEN0001mPrice = totSumKCTEN0001mPrice + oCPGDataStatistics.FItemList(i).FsumKCTEN0001mPrice
totSumKakaopayPrice = totSumKakaopayPrice + oCPGDataStatistics.FItemList(i).FsumKakaopayPrice
totSumnaverpayPrice = totSumnaverpayPrice + oCPGDataStatistics.FItemList(i).FsumnaverpayPrice
totSumnaverpayPoint = totSumnaverpayPoint + oCPGDataStatistics.FItemList(i).FsumnaverpayPoint
totSumpaycoPrice = totSumpaycoPrice + oCPGDataStatistics.FItemList(i).FsumpaycoPrice
totSumbankipkumPrice = totSumbankipkumPrice + oCPGDataStatistics.FItemList(i).FsumbankipkumPrice
totSumbankipkum_fingersPrice = totSumbankipkum_fingersPrice + oCPGDataStatistics.FItemList(i).Fsumbankipkum_fingersPrice
totSumbankipkum_10x10Price = totSumbankipkum_10x10Price + oCPGDataStatistics.FItemList(i).Fsumbankipkum_10x10Price
totSumbankrefundPrice = totSumbankrefundPrice + oCPGDataStatistics.FItemList(i).FsumbankrefundPrice
totSumbankrefund_10x10Price = totSumbankrefund_10x10Price + oCPGDataStatistics.FItemList(i).Fsumbankrefund_10x10Price
totSumbankrefund_fingersPrice = totSumbankrefund_fingersPrice + oCPGDataStatistics.FItemList(i).Fsumbankrefund_fingersPrice
totSum10x10_2Price = totSum10x10_2Price + oCPGDataStatistics.FItemList(i).Fsum10x10_2Price
totSumR5523Price = totSumR5523Price + oCPGDataStatistics.FItemList(i).FsumR5523Price
totSummobiliansPrice = totSummobiliansPrice + oCPGDataStatistics.FItemList(i).FsummobiliansPrice
totSumPGgifticonPrice = totSumPGgifticonPrice + oCPGDataStatistics.FItemList(i).FsumPGgifticonPrice
totSumPGgifttingPrice = totSumPGgifttingPrice + oCPGDataStatistics.FItemList(i).FsumPGgifttingPrice
totSumPGokcashbagPrice = totSumPGokcashbagPrice + oCPGDataStatistics.FItemList(i).FsumPGokcashbagPrice
totSumPGtossPrice = totSumPGtossPrice + oCPGDataStatistics.FItemList(i).FsumPGtossPrice
totSumPGConvinienspayPrice = totSumPGConvinienspayPrice + oCPGDataStatistics.FItemList(i).FsumPGConvinienspayPrice

totSumPrice = totSumPrice + oCPGDataStatistics.FItemList(i).FtotSumPrice

totSumCardJungsanPrice = totSumCardJungsanPrice + oCPGDataStatistics.FItemList(i).FsumCardJungsanPrice
totSumBankJungsanPrice = totSumBankJungsanPrice + oCPGDataStatistics.FItemList(i).FsumBankJungsanPrice
totSumVBankJungsanPrice = totSumVBankJungsanPrice + oCPGDataStatistics.FItemList(i).FsumVBankJungsanPrice
totSumTenOutBankJungsanPrice = totSumTenOutBankJungsanPrice + oCPGDataStatistics.FItemList(i).FsumTenOutBankJungsanPrice
totSumTenInBankJungsanPrice = totSumTenInBankJungsanPrice + oCPGDataStatistics.FItemList(i).FsumTenInBankJungsanPrice
totSumHPJungsanPrice = totSumHPJungsanPrice + oCPGDataStatistics.FItemList(i).FsumHPJungsanPrice
totSumGifttingJungsanPrice = totSumGifttingJungsanPrice + oCPGDataStatistics.FItemList(i).FsumGifttingJungsanPrice
totSumGifticonJungsanPrice = totSumGifticonJungsanPrice + oCPGDataStatistics.FItemList(i).FsumGifticonJungsanPrice
totSumOKJungsanPrice = totSumOKJungsanPrice + oCPGDataStatistics.FItemList(i).FsumOKJungsanPrice
totSumAllAtJungsanPrice = totSumAllAtJungsanPrice + oCPGDataStatistics.FItemList(i).FsumAllAtJungsanPrice

totSumteenxteen3JungsanPrice = totSumteenxteen3JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumteenxteen3JungsanPrice
totSumteenxteen4JungsanPrice = totSumteenxteen4JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumteenxteen4JungsanPrice
totSumteenxteen5JungsanPrice = totSumteenxteen5JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumteenxteen5JungsanPrice
totSumteenxteen6JungsanPrice = totSumteenxteen6JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumteenxteen6JungsanPrice
totSumteenxteen8JungsanPrice = totSumteenxteen8JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumteenxteen8JungsanPrice
totSumteenxteen9JungsanPrice = totSumteenxteen9JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumteenxteen9JungsanPrice
totSumteenteen10JungsanPrice = totSumteenteen10JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumteenteen10JungsanPrice
totSumtenbyten01JungsanPrice = totSumtenbyten01JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumtenbyten01JungsanPrice
totSumtenbyten02JungsanPrice = totSumtenbyten02JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumtenbyten02JungsanPrice
totSumteenxteehaJungsanPrice = totSumteenxteehaJungsanPrice + oCPGDataStatistics.FItemList(i).FsumteenxteehaJungsanPrice
totSumteenxteenrJungsanPrice = totSumteenxteenrJungsanPrice + oCPGDataStatistics.FItemList(i).FsumteenxteenrJungsanPrice
totSumteenteenspJungsanPrice = totSumteenteenspJungsanPrice + oCPGDataStatistics.FItemList(i).FsumteenteenspJungsanPrice
totSumteenteenapJungsanPrice = totSumteenteenapJungsanPrice + oCPGDataStatistics.FItemList(i).FsumteenteenapJungsanPrice
totSumKCTEN0001mJungsanPrice = totSumKCTEN0001mJungsanPrice + oCPGDataStatistics.FItemList(i).FsumKCTEN0001mJungsanPrice
totsumKakaoJungsanPrice = totsumKakaoJungsanPrice + oCPGDataStatistics.FItemList(i).FsumKakaoJungsanPrice
totSumnaverpayJungsanPrice = totSumnaverpayJungsanPrice + oCPGDataStatistics.FItemList(i).FsumnaverpayJungsanPrice
totSumnaverpayJungsanPoint = totSumnaverpayJungsanPoint + oCPGDataStatistics.FItemList(i).FsumnaverpayJungsanPoint
totSumpaycoJungsanPrice = totSumpaycoJungsanPrice + oCPGDataStatistics.FItemList(i).FsumpaycoJungsanPrice
totSumbankipkumJungsanPrice = totSumbankipkumJungsanPrice + oCPGDataStatistics.FItemList(i).FsumbankipkumJungsanPrice
totSumbankipkum_fingersJungsanPrice = totSumbankipkum_fingersJungsanPrice + oCPGDataStatistics.FItemList(i).Fsumbankipkum_fingersJungsanPrice
totSumbankipkum_10x10JungsanPrice = totSumbankipkum_10x10JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumbankipkum_10x10JungsanPrice
totSumbankrefundJungsanPrice = totSumbankrefundJungsanPrice + oCPGDataStatistics.FItemList(i).FsumbankrefundJungsanPrice
totSumbankrefund_10x10JungsanPrice = totSumbankrefund_10x10JungsanPrice + oCPGDataStatistics.FItemList(i).Fsumbankrefund_10x10JungsanPrice
totSumbankrefund_fingersJungsanPrice = totSumbankrefund_fingersJungsanPrice + oCPGDataStatistics.FItemList(i).Fsumbankrefund_fingersJungsanPrice
totSum10x10_2JungsanPrice = totSum10x10_2JungsanPrice + oCPGDataStatistics.FItemList(i).Fsum10x10_2JungsanPrice
totSumR5523JungsanPrice = totSumR5523JungsanPrice + oCPGDataStatistics.FItemList(i).FsumR5523JungsanPrice
totSummobiliansJungsanPrice = totSummobiliansJungsanPrice + oCPGDataStatistics.FItemList(i).FsummobiliansJungsanPrice
totSumPGgifticonJungsanPrice = totSumPGgifticonJungsanPrice + oCPGDataStatistics.FItemList(i).FsumPGgifticonJungsanPrice
totSumPGgifttingJungsanPrice = totSumPGgifttingJungsanPrice + oCPGDataStatistics.FItemList(i).FsumPGgifttingJungsanPrice
totSumPGokcashbagJungsanPrice = totSumPGokcashbagJungsanPrice + oCPGDataStatistics.FItemList(i).FsumPGokcashbagJungsanPrice
totSumPGtossJungsanPrice = totSumPGtossJungsanPrice + oCPGDataStatistics.FItemList(i).FsumPGtossJungsanPrice
totSumPGchaiJungsanPrice = totSumPGchaiJungsanPrice + oCPGDataStatistics.FItemList(i).FsumPGchaiJungsanPrice
totSumConvinienspayJungsanPrice = totSumConvinienspayJungsanPrice + oCPGDataStatistics.FItemList(i).FsumPGConvinienspayJungsanPrice

totSumJungsanPrice = totSumJungsanPrice + oCPGDataStatistics.FItemList(i).FtotSumJungsanPrice

yyyy = Left(oCPGDataStatistics.FItemList(i).Fyyyymmdd, 4)
mm = Right(Left(oCPGDataStatistics.FItemList(i).Fyyyymmdd, 7), 2)
dd = Right(Left(oCPGDataStatistics.FItemList(i).Fyyyymmdd, 10), 2)

%>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td>
		<a href="javascript:popPGDataList('<%= yyyy %>', '<%= mm %>', '<%= dd %>', '<%= PGuserid %>')">
			<%= oCPGDataStatistics.FItemList(i).Fyyyymmdd %>
		</a>
	</td>

	<% if (sumgubun = "appMethod") then %>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumCardPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumBankPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumVBankPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumTenOutBankPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumTenInBankPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumHPPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumGifttingPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumGifticonPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumOKPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumAllAtPrice, 0) %></td>
	<% else %>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen3Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen4Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen5Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen6Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen8Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen9Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenteen10Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumtenbyten01Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumtenbyten02Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumteenxteehaPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumteenxteenrPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumteenteenspPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumteenteenapPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumKCTEN0001mPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumKakaopayPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumnaverpayPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumnaverpayPoint, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumpaycoPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumbankipkumPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumbankipkum_10x10Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumbankipkum_fingersPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumbankrefundPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumbankrefund_10x10Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumbankrefund_fingersPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsum10x10_2Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumR5523Price, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsummobiliansPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGgifticonPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGgifttingPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGokcashbagPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGtossPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGchaiPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGConvinienspayPrice, 0) %></td>
	<% end if %>

	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FtotSumPrice, 0) %></td>

	<% if (sumgubun = "appMethod") then %>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumCardJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumBankJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumVBankJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumTenOutBankJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumTenInBankJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumHPJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumGifttingJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumGifticonJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumOKJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumAllAtJungsanPrice, 0) %></td>
	<% else %>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen3JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen4JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen5JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen6JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen8JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenxteen9JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumteenteen10JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumtenbyten01JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumtenbyten02JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumteenxteehaJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumteenxteenrJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumteenteenspJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumteenteenapJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumKCTEN0001mJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumKakaoJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumnaverpayJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumnaverpayJungsanPoint, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumpaycoJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumbankipkumJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumbankipkum_10x10JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumbankipkum_fingersJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumbankrefundJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumbankrefund_10x10JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsumbankrefund_fingersJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).Fsum10x10_2JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumR5523JungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsummobiliansJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGgifticonJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGgifttingJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGokcashbagJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGtossJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGchaiJungsanPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FsumPGConvinienspayJungsanPrice, 0) %></td>
	<% end if %>

	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FtotSumJungsanPrice, 0) %></td>

	<td>
		<!--
		<input type="button" class="button" value="새로고침(ON UPLUS)" onClick="jsReloadOnPGData('uplus', '<%= oCPGDataStatistics.FItemList(i).Fyyyymmdd %>');">
		<input type="button" class="button" value="새로고침(ON INICIS)" onClick="jsReloadOnPGData('inicis', '<%= oCPGDataStatistics.FItemList(i).Fyyyymmdd %>');">
		-->
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="FFFFFF">
	<td>합계</td>

	<% if (sumgubun = "appMethod") then %>
	<td align="right"><%= FormatNumber(totSumCardPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumBankPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumVBankPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumTenOutBankPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumTenInBankPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumHPPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumGifttingPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumGifticonPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumOKPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumAllAtPrice, 0) %>&nbsp;</td>
	<% else %>
	<td align="right"><%= FormatNumber(totSumteenxteen3Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen4Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen5Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen6Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen8Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen9Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenteen10Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumtenbyten01Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumtenbyten02Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteehaPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteenrPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenteenspPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenteenapPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumKCTEN0001mPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumKakaopayPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumnaverpayPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumnaverpayPoint, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumpaycoPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankipkumPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankipkum_10x10Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankipkum_fingersPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankrefundPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankrefund_10x10Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankrefund_fingersPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSum10x10_2Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumR5523Price, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSummobiliansPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGgifticonPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGgifttingPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGokcashbagPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGtossPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGchaiPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGConvinienspayPrice, 0) %>&nbsp;</td>
	<% end if %>
	<td align="right"><%= FormatNumber(totSumPrice, 0) %>&nbsp;</td>

	<% if (sumgubun = "appMethod") then %>
	<td align="right"><%= FormatNumber(totSumCardJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumBankJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumVBankJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumTenOutBankJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumTenInBankJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumHPJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumGifttingJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumGifticonJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumOKJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumAllAtJungsanPrice, 0) %>&nbsp;</td>
	<% else %>
	<td align="right"><%= FormatNumber(totSumteenxteen3JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen4JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen5JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen6JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen8JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteen9JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenteen10JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumtenbyten01JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumtenbyten02JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteehaJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenxteenrJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenteenspJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumteenteenapJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumKCTEN0001mJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totsumKakaoJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumnaverpayJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumnaverpayJungsanPoint, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumpaycoJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankipkumJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankipkum_10x10JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankipkum_fingersJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankrefundJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankrefund_10x10JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumbankrefund_fingersJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSum10x10_2JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumR5523JungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSummobiliansJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGgifticonJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGgifttingJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGokcashbagJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGtossJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumPGchaiJungsanPrice, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(totSumConvinienspayJungsanPrice, 0) %>&nbsp;</td>
	<% end if %>

	<td align="right"><%= FormatNumber(totSumJungsanPrice, 0) %>&nbsp;</td>

	<td align="right"></td>
</tr>
</table>

<form name="frmAct" method="post" action="/admin/maechul/pgdata/pgdata_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="yyyymmdd" value="">
</form>

<%
set oCPGDataStatistics = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
