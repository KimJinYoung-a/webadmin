<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : cs센터
' History	:  2007.06.01 이상구 생성
'              2023.11.15 한용민 수정(6개월이전 데이터도 처리가능하게 로직 변경)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_mifinishcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
dim research, page, divcd, makerid, vSiteName, itemid, Dtype, fromdate, todate, nexttodate
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, dplusOver, dplusLower, MifinishReason, MifinishState, sortby
dim exinmaychulgoday, exoldcs, exchangemindreturn, exregbycs, order6MonthBefore, csmifinish
dim OCSBrandMemo, OCSItemMemo, ix,iy
	research 	= requestCheckVar(request("research"),32)
	page 		= requestCheckVar(request("page"),32)
	divcd 		= requestCheckVar(request("divcd"),32)
	makerid 	= requestCheckVar(request("makerid"),32)
	vSiteName 	= requestCheckVar(request("vSiteName"),32)
	itemid 		= requestCheckVar(request("itemid"),32)
	Dtype 		= requestCheckVar(request("Dtype"),32)
	yyyy1   	= requestCheckVar(request("yyyy1"),4)
	mm1     	= requestCheckVar(request("mm1"),2)
	dd1     	= requestCheckVar(request("dd1"),2)
	yyyy2   	= requestCheckVar(request("yyyy2"),4)
	mm2     	= requestCheckVar(request("mm2"),2)
	dd2     	= requestCheckVar(request("dd2"),2)
	dplusOver   	= requestCheckVar(request("dplusOver"),10)
	dplusLower   	= requestCheckVar(request("dplusLower"),10)
	MifinishReason 	= requestCheckVar(request("MifinishReason"),2)
	MifinishState  	= requestCheckVar(request("MifinishState"),2)
	sortby			= requestCheckVar(request("sortby"),32)
	exinmaychulgoday	= requestCheckVar(request("exinmaychulgoday"),32)
	exoldcs				= requestCheckVar(request("exoldcs"),32)
	exchangemindreturn	= requestCheckVar(request("exchangemindreturn"),32)
	exregbycs	= requestCheckVar(request("exregbycs"),32)
	order6MonthBefore	= requestCheckVar(request("order6MonthBefore"),1)

if (page="") then page=1
if (Dtype="") then Dtype = "dday"

if (research = "") then
	if (dplusOver = "") then
		dplusOver = "7"
	end if

	exoldcs = "Y"
	''exchangemindreturn = "Y"
end if

if (yyyy1="") then
	todate = Left(CStr(now()),10)

	yyyy2 = Left(todate,4)
	mm2   = Mid(todate,6,2)
	dd2   = Mid(todate,9,2)

	fromdate = DateSerial(yyyy2,mm2-2, dd2+1)

	yyyy1 = Left(fromdate,4)
	mm1   = Mid(fromdate,6,2)
	dd1   = Mid(fromdate,9,2)
end if

nexttodate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

set csmifinish = new CCSMifinishMaster
	csmifinish.FRectDivCD = divcd
	csmifinish.FRectDesignerID = makerid
	csmifinish.FPageSize = 50
	csmifinish.FCurrPage = page
	csmifinish.FRectMifinishReason = MifinishReason
	csmifinish.FRectMifinishState  = MifinishState
	csmifinish.FRectItemID = itemid
	csmifinish.FRectSiteName = vSiteName
	csmifinish.FRectSortBy = sortby
	csmifinish.FRectExInMayChulgoDay = exinmaychulgoday
	csmifinish.FRectExOldCS = exoldcs
	csmifinish.FRectExChangeMindReturn = exchangemindreturn
	csmifinish.FRectExRegbyCS = exregbycs
	csmifinish.FRectorder6MonthBefore = order6MonthBefore

	if (Dtype = "topN") then
		todate = Left(CStr(now()),10)
		fromdate = Left(CStr(DateAdd("m", -2, now())),10)
		nexttodate = Left(CStr(DateAdd("d", 1, CDate(todate))),10)

		csmifinish.FRectRegStart = fromdate
		csmifinish.FRectRegEnd = nexttodate
		csmifinish.FPageSize = 300

		csmifinish.getUpcheMifinishList
	elseif (Dtype = "date") then
		csmifinish.FRectRegStart = LEft(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
		csmifinish.FRectRegEnd = nexttodate

		csmifinish.getUpcheMifinishList
	elseif (Dtype = "dday") then
		csmifinish.FRectdplusOver = dplusOver
		csmifinish.FRectdplusLower = dplusLower

		csmifinish.getUpcheMifinishList
	end if

set OCSBrandMemo = new CCSBrandMemo
	OCSBrandMemo.FRectMakerid = makerid

	if (makerid <> "") then
		OCSBrandMemo.GetBrandMemo
	end if

set OCSItemMemo = new CCSItemMemo
	OCSItemMemo.FRectItemId = itemid

	if (itemid <> "") then
		OCSItemMemo.GetItemidMemo
	end if

%>
<script type="text/javascript">

function chkSubmit(){
    var frm = document.frm;

    if ((frm.itemid.value.length>0)&&(!IsDigit(frm.itemid.value))){
        alert('상품번호는 숫자로 입력하세요.');
        frm.itemid.focus();
        return;
    }

    frm.yyyy1.disabled=false;
    frm.yyyy2.disabled=false;
    frm.mm1.disabled=false;
    frm.mm2.disabled=false;
    frm.dd1.disabled=false;
    frm.dd2.disabled=false;

    frm.submit();
}

function MifinishCSMaster(v){
	var popwin = window.open("/cscenter/mifinish/cs_mifinishmaster_main.asp?asid=" + v,"MifinishMaster","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ViewItem(itemid){
window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"sample");
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function chkComp(comp){
	var selectedval = comp.value;

	if (selectedval == "topN") {
	    comp.form.yyyy1.disabled = true;
	    comp.form.yyyy2.disabled = true;
	    comp.form.mm1.disabled = true;
	    comp.form.mm2.disabled = true;
	    comp.form.dd1.disabled = true;
	    comp.form.dd2.disabled = true;

	    comp.form.dplusOver.disabled = true;
	    comp.form.dplusLower.disabled = true;
	} else if (selectedval == "date") {
	    comp.form.yyyy1.disabled = false;
	    comp.form.yyyy2.disabled = false;
	    comp.form.mm1.disabled = false;
	    comp.form.mm2.disabled = false;
	    comp.form.dd1.disabled = false;
	    comp.form.dd2.disabled = false;

	    comp.form.dplusOver.disabled = true;
	    comp.form.dplusLower.disabled = true;
	} else if (selectedval == "dday") {
	    comp.form.yyyy1.disabled = true;
	    comp.form.yyyy2.disabled = true;
	    comp.form.mm1.disabled = true;
	    comp.form.mm2.disabled = true;
	    comp.form.dd1.disabled = true;
	    comp.form.dd2.disabled = true;

	    comp.form.dplusOver.disabled = false;
	    comp.form.dplusLower.disabled = false;
	}
}

function searchByMakerId(frm, makerid) {
	frm.makerid.value = makerid;
	frm.itemid.value = "";

	chkSubmit();
}

function searchByItemId(frm, itemid) {
	frm.itemid.value = itemid;

	chkSubmit();
}

function jsShowHideObject(id) {
	if (document.getElementById) {
		obj = document.getElementById(id);

		if (obj.style.display == "none") {
			obj.style.display = "";
		} else {
			obj.style.display = "none";
		}
	}
}

function jsUpcheBrandReturnMemo(makerid, tableobj)
{

	if (makerid == "") {
		alert("먼저 브랜드로 검색하세요.");
		return;
	}

	jsShowHideObject(tableobj);
}

function submitSaveBrandMemo(frm)
{

	if (frm.makerid.value == "") {
		alert("먼저 브랜드로 검색하세요.");
		return;
	}

	if (confirm("저장하시겠습니까?") == true) {
		frm.submit();
	}
}

function jsUpcheItemReturnMemo(itemid, tableobj)
{

	if (itemid == "") {
		alert("먼저 상품코드로 검색하세요.");
		return;
	}

	jsShowHideObject(tableobj);
}

function submitSaveItemMemo(frm)
{

	if (frm.itemid.value == "") {
		alert("먼저 상품코드로 검색하세요.");
		return;
	}

	if (confirm("저장하시겠습니까?") == true) {
		frm.submit();
	}
}

function jsMultiReturnReason(makerid, itemid, tableobj)
{

	if ((makerid == "") && (itemid == "")) {
		alert("브랜드 또는 상품코드로 검색하세요.");
		return;
	}

	jsShowHideObject(tableobj);
}

function jsSetReturnReason(frm) {
	if (frm.regReturnReason.value == "26") {
		frm.nextactday.value = "";
	}

	jsSetSMSMailText(frm);
}

function jsSetSMSMailText(frm) {
	jsSetSMSText(frm);
	jsSetMailText(frm);
}

function jsSetSMSText(frm) {
	var smsText;

	smsText = "";

	if (frm.regReturnReason.value == "25") {
		smsText = "[텐바이텐 업체반품안내] 고객님, 접수하신 상품 [상품명]([상품코드])은 업체로 반송 후";
		smsText = smsText + " 반송하신 운송장번호 알려주시면 확인 후 환불처리 진행해드리겠습니다.";
		smsText = smsText + " 아직 미반송하신 경우 업체로 반송 부탁 드립니다.";
	} else if (frm.regReturnReason.value == "26") {
		smsText = "[텐바이텐 반품철회안내] 고객님, 접수하신 상품 [상품명]([상품코드])은 제작상품으로 반품이 어려우십니다."
		smsText = smsText + " 도움드리지못해 죄송하며 접수철회됨에 깊은 양해바랍니다.."
		smsText = smsText + " 혹, 이미 반송하신 경우 고객센터로 연락 부탁드립니다.감사합니다."
	}

	frm.sendsmsmsg.value = smsText;
}

function jsSetMailText(frm) {
	var mailText;

	mailText = "안녕하세요. 고객님\n";
	mailText = mailText + "텐바이텐 고객행복센터입니다.\n\n";

	if (frm.regReturnReason.value == "25") {
		mailText = mailText + "고객님께서 반품접수하신 상품을 업체로 반송하셨는지요?\n"
		mailText = mailText + "아직 반품 이전이시면 수령하신 택배사 이용하여 업체로 반품 부탁 드리며,\n"
		mailText = mailText + "반품 후 반송장(반품 운송장)번호를\n\n"
		mailText = mailText + "텐바이텐 홈페이지(PC화면) > 마이텐바이텐 > 내가 신청한 서비스\n\n"
		mailText = mailText + "에서 반품 접수하신 내역에 입력해 주시면 보다 빠른 환불처리 가능함을 알려드립니다.\n\n"
		mailText = mailText + "감사합니다."
		mailText = mailText + ""
	} else if (frm.regReturnReason.value == "26") {
		mailText = mailText + "고객님.\n"
		mailText = mailText + "다른게 아니오라 고객님께서 반품접수하신 상품 [상품명]([상품코드])은 제작상품으로\n"
		mailText = mailText + "반품이 어려우십니다\n"
		mailText = mailText + "죄송하지만, 접수해주신 반품접수는 철회되었으니, 이점 깊은 양해부탁드리며\n"
		mailText = mailText + "도움드리지 못해 정말 죄송합니다\n"
		mailText = mailText + "추후 반품범위에 대한 상품을 더욱 더 넓힐수 있도록 노력하겠습니다\n"
		mailText = mailText + "아울러 이미 업체로 반송하신경우시면, 번거로우시더라도 반송장번호와 택배사를\n"
		mailText = mailText + "확인하시어 고객센터로 연락부탁드립니다.\n"
		mailText = mailText + "1:1 게시판 또는 메일회신해주셔도 되십니다\n\n\n"
		mailText = mailText + "저희는 더욱더 노력하고, 늘 변함없는 마음으로 고객님을 모시겠습니다.\n"
		mailText = mailText + "다른 더 궁금하신 사항은 언제든지 고객센터로 연락주시면 친절히 안내해드리겠으며\n"
		mailText = mailText + "언제나 행복한 날 되시길 바랍니다.~\n"
	}

	mailText = mailText + "\n\n고객센터 업무시간\n"
	mailText = mailText + "평일 AM 09:00 ~ PM 06:00\n"
	mailText = mailText + "점심시간 PM 12:00~01:00 토ㆍ일ㆍ공휴일 휴무\n"
	mailText = mailText + "☎ 1644-6030\n"
	mailText = mailText + "customer@10x10.co.kr\n"

	frm.sendmailmsg.value = mailText;
}

function CheckNcalendarOpen(returneason, nextactday) {
	if (returneason.value == "26") {
		// 반품불가
		alert("반품불가 안내일 경우 다음처리예정일을 입력할 수 없습니다.");
		return;
	}

	calendarOpen(nextactday);
}

function multiReturnInput(frm) {
	if (CheckSelected() != true) {
		alert("선택된 주문이 없습니다.");
		return;
	}

	if (frm.regReturnReason.value == "") {
		alert("미출고 사유를 선택하세요.");
		frm.regReturnReason.focus();
		return;
	}

	if ((frm.ckSendSMS.checked != true) && (frm.ckSendEmail.checked != true)) {
		alert("SMS 와 메일발송 둘중 하나는 체크해야 합니다.");
		return;
	}

	/*
	if ((frm.nextactday.value.length != 10) && (frm.regReturnReason.value != "26")) {
		alert("다음처리예정일을 입력하세요.");
		frm.nextactday.focus();
		return;
	}
	*/

	if (frm.sendsmsmsg.value == "") {
		alert("SMS발송문구를 입력하세요.");
		frm.sendsmsmsg.focus();
		return;
	}

	if (frm.sendmailmsg.value == "") {
		alert("MAIL발송문구를 입력하세요.");
		frm.sendmailmsg.focus();
		return;
	}

	if (confirm("일괄저장하시겠습니까?") == true) {

		for (var i=0;i<document.forms.length;i++){
			f = document.forms[i];
			if (f.name.substr(0,9)=="frmBuyPrc") {
				if (f.csdetailidx.checked) {
					frm.arrcsdetailidx.value = frm.arrcsdetailidx.value + "," + f.csdetailidx.value;
				}
			}
		}

		frm.submit();
	}
}

function CheckSelected(){
	var pass = false;
	var frm;

	for (var i = 0; i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.csdetailidx.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function getOnLoad(){
    var idx = 0;

    for (var i = 0; i < document.frm.Dtype.length; i++) {
    	if (document.frm.Dtype[i].value == "<%= Dtype %>") {
    		idx = i;
    		break;
    	}
    }

    chkComp(document.frm.Dtype[idx]);
}

window.onload=getOnLoad;

</script>


<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>" height="120">검색<br>조건</td>
		<td align="left">
			구분 :
			<select name="divcd" class="select">
				<option value="">-전체-</option>
				<option value="chulgocs" <%=CHKIIF(divcd="chulgocs","selected","")%>>업체 CS출고</option>
				<option value="returncs" <%=CHKIIF(divcd="returncs","selected","")%>>업체 반품</option>
			</select>
			&nbsp;
			브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			&nbsp;
			사이트 :
            <% call drawSelectBoxXSiteOrderInputPartnerCS("vSiteName", vSiteName) %>
			&nbsp;
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">
		</td>

		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="chkSubmit();">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td>
			<input type="radio" name="Dtype" value="topN" <%= cHKIIF(Dtype="topN","checked","") %> onClick="chkComp(this);" >TOP <%= CHKIIF(Dtype = "topN",csmifinish.FPageSize,100) %>개(최근2달)
			&nbsp;
			<input type="radio" name="Dtype" value="date" <%= cHKIIF(Dtype="date","checked","") %>  onClick="chkComp(this);" >검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			<input type="radio" name="Dtype" value="dday" <%= cHKIIF(Dtype="dday","checked","") %>  onClick="chkComp(this);" >소요일수 :
			<select class="select" name="dplusOver">
				<option value="" >전체</option>
				<option value="below3day" <%= CHKIIF(dplusOver="below3day","selected","") %> >D+3 미만전체</option>
				<option value="3" <%= CHKIIF(dplusOver="3","selected","") %> >D+3 이상</option>
				<option value="4" <%= CHKIIF(dplusOver="4","selected","") %> >D+4이상</option>
				<option value="7" <%= CHKIIF(dplusOver="7","selected","") %> >D+7이상</option>
				<option value="14" <%= CHKIIF(dplusOver="14","selected","") %> >D+14이상</option>
			</select>
			~
			<select class="select" name="dplusLower">
				<option value="" >전체</option>
				<option value="7" <%= CHKIIF(dplusLower="7","selected","") %> >D+7미만</option>
				<option value="30" <%= CHKIIF(dplusLower="30","selected","") %> >D+30이하</option>
				<option value="60" <%= CHKIIF(dplusLower="60","selected","") %> >D+60이하</option>
				<option value="90" <%= CHKIIF(dplusLower="90","selected","") %> >D+90이하</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td>
			미처리사유 :
			<select class="select" name="MifinishReason">
				<option value="">전체</option>
				<option value=""> [출고] --------</option>
				<option value="00" <%= CHKIIF((MifinishReason="00" and (divcd = "" or divcd = "chulgocs")),"selected","") %> >&nbsp;&nbsp;입력이전</option>
				<option value="03" <%= CHKIIF(MifinishReason="03","selected","") %> >&nbsp;&nbsp;출고지연</option>
				<option value="05" <%= CHKIIF(MifinishReason="05","selected","") %> >&nbsp;&nbsp;품절출고불가</option>
				<option value="02" <%= CHKIIF(MifinishReason="02","selected","") %> >&nbsp;&nbsp;주문제작</option>
				<option value="04" <%= CHKIIF(MifinishReason="04","selected","") %> >&nbsp;&nbsp;예약상품</option>
				<option value=""> [반품] --------</option>
				<option value="00" <%= CHKIIF((MifinishReason="00" and divcd = "returncs"),"selected","") %> >&nbsp;&nbsp;입력이전</option>
				<option value="25" <%= CHKIIF(MifinishReason="25","selected","") %> >&nbsp;&nbsp;송장입력 안내</option>
				<option value="26" <%= CHKIIF(MifinishReason="26","selected","") %> >&nbsp;&nbsp;반품불가 안내</option>
				<option value="21" <%= CHKIIF(MifinishReason="21","selected","") %> >&nbsp;&nbsp;고객 부재</option>
				<option value="22" <%= CHKIIF(MifinishReason="22","selected","") %> >&nbsp;&nbsp;고객 반품예정</option>
				<option value="23" <%= CHKIIF(MifinishReason="23","selected","") %> >&nbsp;&nbsp;CS택배접수</option>
				<option value="12" <%= CHKIIF(MifinishReason="12","selected","") %> >&nbsp;&nbsp;업체지연</option>
				<option value="41" <%= CHKIIF(MifinishReason="41","selected","") %> >&nbsp;&nbsp;택배사 수거지연</option>
			</select>
			&nbsp;
			처리구분 :
			<select class="select" name="MifinishState">
				<option value="">전체</option>
				<option value="0" <%= CHKIIF(MifinishState="0","selected","") %> >CS(CALL)미처리</option>
				<option value="4" <%= CHKIIF(MifinishState="4","selected","") %> >고객안내</option>
				<option value="6" <%= CHKIIF(MifinishState="6","selected","") %> >CS처리완료</option>
			</select>
			&nbsp;
			정렬순서 :
			<select class="select" name="sortby">
				<option value="">소요일수</option>
				<option value="makerid" <%= CHKIIF(sortby="makerid","selected","") %> >브랜드</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td>
			<input type="checkbox" class="checkbox" name="exinmaychulgoday" value="Y" <%= CHKIIF(exinmaychulgoday="Y","checked","") %>> 처리예정일 이전 미처리CS 제외
			<input type="checkbox" class="checkbox" name="exoldcs" value="Y" <%= CHKIIF(exoldcs="Y","checked","") %>> 장기간(3개월) 미처리CS 제외
			<input type="checkbox" class="checkbox" name="exchangemindreturn" value="Y" <%= CHKIIF(exchangemindreturn="Y","checked","") %>> 변심반품 제외
			<input type="checkbox" class="checkbox" name="exregbycs" value="Y" <%= CHKIIF(exregbycs="Y","checked","") %> > 고객직접접수 내역만
			<input type="checkbox" class="checkbox" name="exchangemindreturn11" value="Y" disabled> 송장미입력 내역만
			<input type="checkbox" class="checkbox" name="order6MonthBefore" value="Y" <% if order6MonthBefore="Y" then response.write "checked" %>>6개월이전주문
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>
* D+4, D+7, D+14 일은 <font color=red>근무일수 기준</font>입니다.<br>
* 교환출고의 경우 교환상품 미회수인 CS 는 제외합니다.
<br>
<input type="button" class="button" value="브랜드 반품관련 메모" onClick="jsUpcheBrandReturnMemo('<%= makerid %>', 'brandmemo');" <% if (divcd <> "returncs") then %>disabled<% end if %> >
<input type="button" class="button" value="상품 반품관련 메모" onClick="jsUpcheItemReturnMemo('<%= itemid %>', 'itemmemo');" <% if (divcd <> "returncs") then %>disabled<% end if %> >
<input type="button" class="button" value="반품관련 안내 일괄입력" onClick="jsMultiReturnReason('<%= makerid %>', '<%= itemid %>', 'regallreturnreason');" <% if (divcd <> "returncs") then %>disabled<% end if %> >
<br>

<form name="frmBrandMemo" method="post" action="cs_mifinishlist_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="modifybrandmemo">
<input type="hidden" name="makerid" value="<%= makerid %>">
<div id="brandmemo" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30"><b>브랜드ID</b></td>
		<td width="20%" bgcolor="FFFFFF"><%= makerid %></td>
		<td width="10%"></td>
		<td width="25%" bgcolor="FFFFFF"></td>
		<td width="10%">최종수정일</td>
		<td bgcolor="FFFFFF"><%= OCSBrandMemo.Freturn_modifyday %></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30"></td>
		<td width="20%" bgcolor="FFFFFF"></td>
		<td width="10%"></td>
		<td width="25%" bgcolor="FFFFFF"></td>
		<td width="10%">작성자</td>
		<td bgcolor="FFFFFF"><%= OCSBrandMemo.Freturn_reguserid %></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">브랜드 반품관련 메모</td>
		<td colspan="5" bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="return_comment" cols="100" rows="7"><%= OCSBrandMemo.Freturn_comment %></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="FFFFFF" colspan = "6" height="35">
			<input type="button" class="button_s" value=" 저장하기 " onClick="submitSaveBrandMemo(frmBrandMemo)">
		</td>
	</tr>
</table>
<br>
</div>
</form>

<form name="frmItemMemo" method="post" action="cs_mifinishlist_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="modifyitemmemo">
<input type="hidden" name="itemid" value="<%= itemid %>">
<div id="itemmemo" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30"><b>상품코드</b></td>
		<td width="20%" bgcolor="FFFFFF" align="left"><%= itemid %></td>
		<td width="10%">상품구분</td>
		<td width="25%" bgcolor="FFFFFF" align="left">
			<input type="radio" name="return_changemindyn" value="Y" <%= CHKIIF((OCSItemMemo.Freturn_changemindyn = "Y" or OCSItemMemo.Freturn_changemindyn = ""),"checked","") %> > 일반
			<input type="radio" name="return_changemindyn" value="N" <%= CHKIIF(OCSItemMemo.Freturn_changemindyn = "N","checked","") %> > 변심반품 불가
		</td>
		<td width="10%"></td>
		<td bgcolor="FFFFFF" align="left">

		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30">최종수정일</td>
		<td width="20%" bgcolor="FFFFFF"  align="left">
			<%= OCSItemMemo.Freturn_modifyday %>
		</td>
		<td width="10%">작성자</td>
		<td width="25%" bgcolor="FFFFFF" align="left">
			<%= OCSItemMemo.Freturn_reguserid %>
		</td>
		<td width="10%"></td>
		<td bgcolor="FFFFFF" align="left">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">상품 반품관련 메모</td>
		<td colspan="5" bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="return_comment" cols="100" rows="7"><%= OCSItemMemo.Freturn_comment %></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="FFFFFF" colspan = "6" height="35">
			<input type="button" class="button_s" value=" 저장하기 " onClick="submitSaveItemMemo(frmItemMemo)">
		</td>
	</tr>
</table>
</form>
<br>
</div>

<form name="frmReturnInput" method="post" action="cs_mifinishlist_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="regallreturnreason">
<input type="hidden" name="arrcsdetailidx" value="">
<div id="regallreturnreason" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="2" height="30"><b>반품 안내 일괄전송</b></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="10%" height="30">반품안내 문구</td>
		<td bgcolor="FFFFFF" align="left">
			<select class="select" name="regReturnReason" onChange="jsSetReturnReason(frmReturnInput)">
				<option value=""></option>
				<option value="25">송장입력 안내</option>
				<option value="26">반품 불가</option>
			</select>
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="10%">다음처리예정일</td>
		<td bgcolor="FFFFFF" align="left">
		    <input class="text" type="text_ro" name="nextactday" value="" size="10" maxlength="10" readonly>
		    <a href="javascript:CheckNcalendarOpen(frmReturnInput.regReturnReason, frmReturnInput.nextactday);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="7%">고객안내</td>
		<td bgcolor="FFFFFF" align="left">
			<input name="ckSendSMS" type="checkbox" checked  >SMS발송&nbsp;
			<input name="ckSendEmail" type="checkbox" checked  >MAIL발송&nbsp;
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">SMS<br>발송내용</td>
		<td bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="sendsmsmsg" cols="52" rows="5"></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">MAIL<br>발송내용</td>
		<td bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="sendmailmsg" cols="90" rows="7"></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="FFFFFF" colspan = "2" height="35">
			<input type="button" class="button" value="반품 안내 일괄전송" onclick="multiReturnInput(frmReturnInput);">
		</td>
	</tr>
</table>
</form>
<br>
</div>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="19">
		<% if Dtype="topN" then %>
		검색결과 : <b><% = csmifinish.FTotalCount %></b> (최대 <%= csmifinish.FPageSize %>건 까지 검색됩니다.)
		<% else %>
			검색결과 : <b><% = csmifinish.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= csmifinish.FTotalpage %></b>
		<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"></td>
		<td width="30">구분</td>
		<td>브랜드ID</td>
		<td width="70">주문번호</td>
		<td width="50">ASID</td>
		<td width="55">주문자</td>
		<td width="55">수령인</td>
		<td width="50">상품코드</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="30">수량</td>
		<td width="60">CS등록일<br>(기준일)</td>
		<td width="35">소요<br>일수</td>
		<td width="105">미처리사유</td>
		<td width="25">송장<br>입력</td>
		<td width="60">처리예정일</td>
		<td width="65">처리구분</td>
		<td width="65">최종수정</td>
		<td width="35">상세<br>정보</td>
	</tr>
	<% if csmifinish.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="19" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for ix=0 to csmifinish.FresultCount-1 %>
	<form name="frmBuyPrc_<%= ix %>" method="post" style="margin:0px;">
	<input type="hidden" name="orderserial" value="<%= csmifinish.FItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<% if csmifinish.FItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="DDDDDD">
	<% end if %>
		<td>
			<input type="checkbox" name="csdetailidx" value="<%= csmifinish.FItemList(ix).Fcsdetailidx %>" <% if csmifinish.FItemList(ix).FMifinishReason<>"00" or csmifinish.FItemList(ix).Fsongjangyn = "Y" then %>disabled<%end if %>>
		</td>
		<td>
			<font color="<%= csmifinish.FItemList(ix).getDivcdColor %>"><%= csmifinish.FItemList(ix).getDivcdStr %></font>
		</td>
		<td>
			<a href="javascript:searchByMakerId(frm, '<%= csmifinish.FItemList(ix).FMakerid %>')">
				<%= csmifinish.FItemList(ix).FMakerid %>
			</a>
		</td>
		<td>
			<a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= csmifinish.FItemList(ix).FOrderSerial %>')" class="zzz">
			<%= csmifinish.FItemList(ix).FOrderSerial %></a>
		</td>
		<td><a href="javascript:PopCSActionEdit(<%= csmifinish.FItemList(ix).Fasid %>,'editreginfo')"><%= csmifinish.FItemList(ix).Fasid %></a></td>
		<td>
			<%= csmifinish.FItemList(ix).FBuyname %><%'= printUserId(csmifinish.FItemList(ix).FBuyname, 1, "*") %>
		</td>
		<td>
			<%= csmifinish.FItemList(ix).FReqname %><%'= printUserId(csmifinish.FItemList(ix).FReqname, 1, "*") %>
		</td>
		<td>
			<a href="javascript:searchByItemId(frm, <%= csmifinish.FItemList(ix).FItemid %>)">
				<%= csmifinish.FItemList(ix).FItemid %>
			</a>
		</td>
		<td align="left">
			<a href="javascript:ViewItem(<% =csmifinish.FItemList(ix).FItemid  %>)"><%= csmifinish.FItemList(ix).FItemname %></a>
			<% if (csmifinish.FItemList(ix).FItemoption<>"") then %>
				<font color="blue">[<%= csmifinish.FItemList(ix).FItemoption %>]</font>
			<% end if %>
		</td>
		<td><%= csmifinish.FItemList(ix).FItemcnt %></td>
		<td><%= Left(csmifinish.FItemList(ix).Fregdate,10) %></td>
		<td><%= csmifinish.FItemList(ix).getDPlusDateStr %></td>
		<td>
		    <%= csmifinish.FItemList(ix).getMifinishText %>

		    <% if not IsNULL(csmifinish.FItemList(ix).FMifinishregdate) then %>
			    <br>(<%= Left(csmifinish.FItemList(ix).FMifinishregdate,10) %>)
		    <% end if %>
		</td>
		<td>
			<% if (csmifinish.FItemList(ix).Fsongjangyn = "Y") then %>Y<% end if %>
		</td>
		<td><%= csmifinish.FItemList(ix).FMifinishipgodate %></td>
		<td><%= csmifinish.FItemList(ix).getMifinishStateText %></td>
		<td>
			<% if Not IsNull(csmifinish.FItemList(ix).Flastupdate) then %>
				<acronym title="<%= csmifinish.FItemList(ix).Flastupdate %>"><%= Left(csmifinish.FItemList(ix).Flastupdate,10) %></acronym><br>
			<% end if %>

			<% if Not IsNull(csmifinish.FItemList(ix).Freguserid) then %>
				<%= csmifinish.FItemList(ix).Freguserid %>
			<% end if %>
		</td>
		<td>
			<a href="javascript:MifinishCSMaster('<%= csmifinish.FItemList(ix).Fasid %>');"><img src="/images/icon_search.jpg" border="0"></a>
		</td>
	</tr>
	</form>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="19" align="center">
		<% if Dtype="topN" then %>
		최대 <%= csmifinish.FPageSize %>건 까지 검색됩니다.
		<% else %>
    		<% if csmifinish.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= csmifinish.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>
    		<% for ix=0 + csmifinish.StartScrollPage to csmifinish.FScrollCount + csmifinish.StartScrollPage - 1 %>
    			<% if ix>csmifinish.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(ix) then %>
    			<font color="red">[<%= ix %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
    			<% end if %>
    		<% next %>

    		<% if csmifinish.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
    	<% end if %>
		</td>
	</tr>
<% end if %>
</table>

<%
set csmifinish = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
