<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 클래스
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim research

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, regextMeachulDate
dim ayyyy1,ayyyy2,amm1,amm2,add1,add2
dim yyyy, mm, dd
dim fromDate ,toDate, tmpDate, tmpDate2
dim afromDate ,atoDate, atmpDate, atmpDate2
dim grpgubun, sellsite
Dim i, j, ipkumdateChk

research = requestCheckvar(request("research"),10)
sellsite = requestCheckvar(request("sellsite"),32)
regextMeachulDate = requestCheckvar(request("regextMeachulDate"),10)

if (regextMeachulDate="") then regextMeachulDate=LEFT(dateadd("d",-1,now()),10)

yyyy1   = requestCheckvar(request("yyyy1"),4)
mm1     = requestCheckvar(request("mm1"),2)
dd1     = requestCheckvar(request("dd1"),2)
yyyy2   = requestCheckvar(request("yyyy2"),4)
mm2     = requestCheckvar(request("mm2"),2)
dd2     = requestCheckvar(request("dd2"),2)

ayyyy1   = requestCheckvar(request("ayyyy1"),4)
amm1     = requestCheckvar(request("amm1"),2)
add1     = requestCheckvar(request("add1"),2)
ayyyy2   = requestCheckvar(request("ayyyy2"),4)
amm2     = requestCheckvar(request("amm2"),2)
add2     = requestCheckvar(request("add2"),2)

grpgubun = requestCheckvar(request("grpgubun"),32)
if (grpgubun = "") then
	grpgubun = "sellsite"
end if

ipkumdateChk = requestCheckvar(request("ipkumdateChk"),1)

dim extjdate : extjdate = requestCheckvar(request("extjdate"),8) ''YYYYMMDD
if (extjdate="") then
    extjdate = replace(LEFT(dateAdd("d",-1,now()),10),"-","")
end if

if (yyyy1="") then
	'// 전달 데이터까지 보여준다.
	tmpDate2 = LEFT(dateadd("d",-1,now()),10)

	fromDate = DateSerial(Cstr(Year(tmpDate2)), Cstr(Month(tmpDate2)), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, DateAdd("m",1,toDate))
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

if (ayyyy1="") then
	'// 전달 데이터까지 보여준다.
	atmpDate2 = LEFT(dateadd("d",-1,now()),10)

	afromDate = DateSerial(Cstr(Year(atmpDate2)), Cstr(Month(atmpDate2)), 1)
	atoDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	ayyyy1 = Cstr(Year(afromDate))
	amm1 = Cstr(Month(afromDate))
	add1 = Cstr(day(afromDate))

	atmpDate = DateAdd("d", -1, DateAdd("m",1,atoDate))
	ayyyy2 = Cstr(Year(atmpDate))
	amm2 = Cstr(Month(atmpDate))
	add2 = Cstr(day(atmpDate))
else
	afromDate = DateSerial(ayyyy1, amm1, add1)
	atoDate = DateSerial(ayyyy2, amm2, add2+1)
end if

Dim oCExtJungsan
set oCExtJungsan = new CExtJungsan
	oCExtJungsan.FPageSize = 366
	oCExtJungsan.FCurrPage = 1

	oCExtJungsan.FRectStartdate = fromDate
	oCExtJungsan.FRectEndDate = toDate

	oCExtJungsan.FRectGroupGubun = grpgubun
	oCExtJungsan.FRectSellSite = sellsite

	oCExtJungsan.FRectIpkumdateChk = ipkumdateChk
	oCExtJungsan.FRectAStartdate = afromDate
	oCExtJungsan.FRectAEndDate = atoDate


    oCExtJungsan.GetExtJungsanStatistic

dim totExtTenMeachulPriceProduct, totExtCommPriceProduct, totExtTenJungsanPriceProduct
dim totExtTenMeachulPriceDeliver, totExtCommPriceDeliver, totExtTenJungsanPriceDeliver
dim totExtTenMeachulPriceEtc, totExtCommPriceEtc, totExtTenJungsanPriceEtc
dim totExtTenMeachulPrice, totExtCommPrice, totExtTenJungsanPrice, totExtTenMiMapping, totExtTenCount
dim totExtTenMiMapping_C, totExtTenCount_C
dim totMiMappOrder, totMiMappOrder_C

dim totExtitemCostProduct, totExtitemCostDeliver, totExtitemCost
dim totExtReducedPriceProduct, totExtOwnCouponPriceProduct, totExtTenCouponPriceProduct
dim totExtReducedPriceDeliver, totExtOwnCouponPriceDeliver, totExtTenCouponPriceDeliver
dim totExtReducedPrice, totExtOwnCouponPrice, totExtTenCouponPrice


%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

/*
function popPGDataList(yyyy1, mm1, dd1, shopid, cardComp) {
	var popup = window.open("pgdata_off.asp?menupos=1562&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy1+"&mm2="+mm1+"&dd2="+dd1+"&shopid="+shopid + "&cardComp=" + cardComp,"popPGDataList","width=1024,height=768,scrollbars=yes,resizable=yes");
	popup.focus();
}

function popjumundetail(yyyy1, mm1, dd1, shopid) {
	var popjumundetail = window.open("popOffShopOrderList.asp?menupos=648&oldlist=&datefg=jumun&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy1+"&mm2="+mm1+"&dd2="+dd1+"&shopid="+shopid+"&buyergubun=","popjumundetail","width=1024,height=768,scrollbars=yes,resizable=yes");
	popjumundetail.focus();
}

function popPGDataListNotMatch(yyyy1, mm1, dd1, shopid) {
	<% 'if (dategubun = "ipkumdate") then %>
		alert("거래일자로 검색한 후 조회가능합니다.");
		return;
	<% 'end if %>

	var popup = window.open("pgdata_off.asp?menupos=1562&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy1+"&mm2="+mm1+"&dd2="+dd1+"&shopid="+shopid + "&excmatchfinish=Y","popPGDataListNotMatch","width=1024,height=768,scrollbars=yes,resizable=yes");
	popup.focus();
}
 */

function changeBg(idName, onoff) {
	var objA = document.getElementById(idName + "a");
	var objB = document.getElementById(idName + "b");

	if (onoff == "on") {
		objA.style.background="F1F1F1";
		objB.style.background="F1F1F1";
	} else {
		objA.style.background="FFFFFF";
		objB.style.background="FFFFFF";
	}
}

function jsGrpAuto(s){
	if(s != ""){
		$("#grpgubun").val("extMeachulDate");
	}
}

function popExtSiteJungsanData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("/admin/maechul/extjungsandata/popRegExtJungsanDataFile.asp","popExtSiteJungsanData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popExtJungsanOrderCheck() {
	var window_width = 1000;
    var window_height = 800;

    var popwin = window.open("/admin/maechul/extjungsandata/popExtJungsanOrderCheck.asp?sellsite=<%=sellsite%>","popExtJungsanOrderCheck","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();

}

function popExtOrderJungsanCheck(){
	var window_width = 1200;
    var window_height = 800;

    var popwin = window.open("/admin/maechul/extjungsandata/popExtOrderJungsanCheck.asp?sellsite=<%=sellsite%>","popExtOrderJungsanCheck","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popSongjangChangeCheck(){
	var window_width = 1200;
    var window_height = 800;

    var popwin = window.open("/admin/maechul/extjungsandata/popSongjangChangeLog.asp?sellsite=<%=sellsite%>&sitescope=50","popSongjangChangeLog","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popExtOrderJungsanFixdate(){
	var window_width = 1200;
    var window_height = 800;

    var popwin = window.open("/admin/maechul/extjungsandata/popJungsanfixdate.asp?sellsite=<%=sellsite%>&sitescope=50","popExtOrderJungsanFixdate","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsMapping(site,jdate,num){
	if(confirm("선택하신 "+site+", "+jdate+" 매핑하시겠습니까?") == true) {
		$("#mappingbtn"+num).hide();
		$("#mappinging"+num).show();

		var str = $.ajax({
			type: "POST",
			url: "/admin/maechul/extjungsandata/extJungsan_mimapping_proc.asp",
			data: "sellsite="+site+"&jdate="+jdate+"",
			dataType: "text",
			async: false
		}).responseText;
		if(str != ""){
			var strArray;
			strArray = str.split('||');

			if(strArray[0]=="0"){
				alert("처리되었습니다.\n\n"+strArray[1]+"");
				$("#mappinging"+num).hide();
			}else{
				alert("RETURN : "+strArray[0]+"\n"+"retMsg : "+strArray[1]+"");
				$("#mappinging"+num).hide();
				$("#mappingbtn"+num).show();
			}
		}else{
			alert("Error [0]");
			$("#mappinging"+num).hide();
			$("#mappingbtn"+num).show();
		}
	}
}

function jsDelMeachul(site,jdate){
	if(confirm("선택하신 "+site+", "+jdate+" 삭제하시겠습니까?") == true) {
		var iurl="/admin/maechul/extjungsandata/extJungsan_process.asp?mode=delmeachulbyday&sellsite="+site+"&yyyymmdd="+jdate;
		var popwin = window.open(iurl,"popExtJungsanDel","width=200, height=200 left=0 top=0 scrollbars=yes resizable=yes status=yes");

		popwin.focus();
	}
}

function addLotteAddCommission(comp,yyyymm){
	var addcom = document.getElementById("lotteaddcommission");

	if (addcom.value.length<1){
		alert('수수료를 입력해주세요.');
		addcom.focus();
		return;
	}

	if (!IsDigit(addcom.value)){
		alert('숫자로 입력해주세요.');
		addcom.focus();
		return;
	}

	if (confirm(yyyymm+' 변동수수료를 입력(수정)하시겠습니까?')){
		var iurl="/admin/maechul/extjungsandata/extJungsan_process.asp?mode=addcommission&sellsite=lotteCom&addval="+addcom.value+"&yyyymm="+yyyymm;
		var popwin = window.open(iurl,"popExtJungsanAdd","width=200, height=200 left=0 top=0 scrollbars=yes resizable=yes status=yes");

		popwin.focus();
	}
}

function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function fnAPIExtJungsanReg(isitename){
    var iregyyyymmdd = document.getElementById("regextMeachulDate").value;
    if (confirm(isitename+ ' (' + iregyyyymmdd + ') 등록하시겠습니까?')){

    }
}

function popExtSiteJungsanByAPI(sitename,comp){
    var yyyymmdd = comp.form.extjdate.value;

    <% If application("Svr_Info")="Dev" Then %>
        var idomain = "http://testwapi.10x10.co.kr";
    <% else %>
        var idomain = "<%=apiURL%>";
    <% end if %>

    if (sitename=="lottecom"){
        var popwin = window.open(idomain+"/outmall/proc/LotteCom_JungsanRsvProc.asp?yyyymmdd="+yyyymmdd,"LotteCom_JungsanRsvProc","width=400,height=200,crollbars=yes,resizable=yes,status=yes");
    }

    if (sitename=="interpark"){
        var popwin = window.open(idomain+"/outmall/proc/Interpark_JungsanRsvProc.asp?yyyymmdd="+yyyymmdd,"Interpark_JungsanRsvProc","width=400,height=200,crollbars=yes,resizable=yes,status=yes");
    }

    if (sitename=="ssg"){
		<% if (sellsite="ssg6006") or (sellsite="ssg6007") then %>
		var popwin = window.open(idomain+"/outmall/ssg/xSiteJungsan_ssg_Process.asp?yyyymmdd="+yyyymmdd+"&sellsite=<%=sellsite%>","ssg_JungsanRsvProc","width=400,height=200,crollbars=yes,resizable=yes,status=yes");
		<% else %>
        var popwin = window.open(idomain+"/outmall/ssg/xSiteJungsan_ssg_Process.asp?yyyymmdd="+yyyymmdd,"ssg_JungsanRsvProc","width=400,height=200,crollbars=yes,resizable=yes,status=yes");
		<% end if %>
    }

    popwin.focus();
}

function popAPIJungsanData() {
    var window_width = 600;
    var window_height = 500;

    var popwin2 = window.open("/admin/maechul/extjungsandata/popApiJungsanData.asp","popAPIJungsanData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin2.focus();
}
</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 제휴몰:
		<select class="select" name="sellsite" onChange="jsGrpAuto(this.value);">
			<option value="">- 전체 -</option>
			<option value="interpark" <% if (sellsite = "interpark") then %>selected<% end if %> >인터파크</option>
			<option value="lotteimall" <% if (sellsite = "lotteimall") then %>selected<% end if %> >롯데아이몰</option>
			<option value="lotteCom" <% if (sellsite = "lotteCom") then %>selected<% end if %> >롯데닷컴</option>
			<option value="11st1010" <% if (sellsite = "11st1010") then %>selected<% end if %> >11번가</option>
			<option value="auction1010" <% if (sellsite = "auction1010") then %>selected<% end if %> >옥션</option>
			<option value="gmarket1010" <% if (sellsite = "gmarket1010") then %>selected<% end if %> >지마켓(NEW)</option>
			<option value="gseshop" <% if (sellsite = "gseshop") then %>selected<% end if %> >GS샵</option>
			<option value="cjmall" <% if (sellsite = "cjmall") then %>selected<% end if %> >CJ몰</option>
			<option value="homeplus" <% if (sellsite = "homeplus") then %>selected<% end if %> >홈플러스</option>
			<option value="ssg" <% if (sellsite = "ssg") then %>selected<% end if %> >SSG</option>
			<option value="ssg6006" <% if (sellsite = "ssg6006") then %>selected<% end if %> >SSG-이마트</option>
			<option value="ssg6007" <% if (sellsite = "ssg6007") then %>selected<% end if %> >SSG-ssg</option>
			<option value="shintvshopping" <% if (sellsite = "shintvshopping") then %>selected<% end if %> >신세계TV쇼핑</option>
			<option value="skstoa" <% if (sellsite = "skstoa") then %>selected<% end if %> >SKSTOA</option>
			<option value="wetoo1300k" <% if (sellsite = "wetoo1300k") then %>selected<% end if %> >1300k</option>
			<option value="nvstorefarm" <% if (sellsite = "nvstorefarm") then %>selected<% end if %> >스토어팜</option>
			<option value="Mylittlewhoopee" <% if (sellsite = "Mylittlewhoopee") then %>selected<% end if %> >스토어팜 캣앤독</option>
<!--
			<option value="nvstorefarmclass" <% if (sellsite = "nvstorefarmclass") then %>selected<% end if %> >스토어팜-클래스</option>
			<option value="nvstoremoonbangu" <% if (sellsite = "nvstoremoonbangu") then %>selected<% end if %> >스토어팜 문방구</option>
-->
			<option value="nvstoregift" <% if (sellsite = "nvstoregift") then %>selected<% end if %> >스토어팜 선물하기</option>
			<option value="ezwel" <% if (sellsite = "ezwel") then %>selected<% end if %> >이지웰페어</option>
			<option value="kakaogift" <% if (sellsite = "kakaogift") then %>selected<% end if %> >카카오기프트</option>
			<option value="kakaostore" <% if (sellsite = "kakaostore") then %>selected<% end if %> >카카오톡스토어</option>
			<option value="boribori1010" <% if (sellsite = "boribori1010") then %>selected<% end if %> >보리보리</option>
			<option value="GS25" <% if (sellsite = "GS25") then %>selected<% end if %> >GS25카달로그</option>
			<option value="coupang" <% if (sellsite = "coupang") then %>selected<% end if %> >쿠팡</option>
			<option value="halfclub" <% if (sellsite = "halfclub") then %>selected<% end if %> >하프클럽</option>
			<option value="hmall1010" <% if (sellsite = "hmall1010") then %>selected<% end if %> >Hmall</option>
			<option value="WMP" <% if (sellsite = "WMP") then %>selected<% end if %> >WMP</option>
			<option value="wmpfashion" <% if (sellsite = "wmpfashion") then %>selected<% end if %> >WMPW패션</option>
			<option value="LFmall" <% if (sellsite = "LFmall") then %>selected<% end if %> >LFmall</option>
			<option value="lotteon" <% if (sellsite = "lotteon") then %>selected<% end if %> >롯데On</option>
			<option value="wconcept1010" <% if (sellsite = "wconcept1010") then %>selected<% end if %> >W컨셉</option>
			<option value="goodshop1010" <% if (sellsite = "goodshop1010") then %>selected<% end if %> >굿샵</option>
			<option value="withnature1010" <% if (sellsite = "withnature1010") then %>selected<% end if %> >자연이랑</option>
			<option value="yes24" <% if (sellsite = "yes24") then %>selected<% end if %> >yes24</option>
			<option value="alphamall" <% if (sellsite = "alphamall") then %>selected<% end if %> >알파몰</option>
			<option value="ohou1010" <% if (sellsite = "ohou1010") then %>selected<% end if %> >오늘의집</option>
			<option value="wadsmartstore" <% if (sellsite = "wadsmartstore") then %>selected<% end if %> >와드스마트스토어</option>
			<option value="casamia_good_com" <% if (sellsite = "casamia_good_com") then %>selected<% end if %> >까사미아</option>
			<option value="cookatmall" <% if (sellsite = "cookatmall") then %>selected<% end if %> >쿠캣</option>
			<option value="aboutpet" <% if (sellsite = "aboutpet") then %>selected<% end if %> >어바웃펫</option>
		</select>
		&nbsp;
		* 합계구분:
		<select class="select" name="grpgubun" id="grpgubun">
			<option value="sellsite" <% if (grpgubun = "sellsite") then %>selected<% end if %> >제휴몰</option>
			<option value="extMeachulDate" <% if (grpgubun = "extMeachulDate") then %>selected<% end if %> >매출일자</option>
		</select>
		&nbsp;
		* 제휴몰정산일자:
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		* 결제일자:
		<input type="checkbox" class="checkbox" name="ipkumdateChk" <%= chkiif(ipkumdateChk="Y", "checked", "") %> value="Y">
		<% DrawDateBoxdynamic ayyyy1, "ayyyy1", ayyyy2, "ayyyy2", amm1, "amm1", amm2, "amm2", add1, "add1", add2, "add2" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF" >
	<td>
	* 쿠폰금액이 제휴정산 수신 후 반영되는 제휴몰 : SSG, Hmall, WMP, LotteiMall, LotteOn, LFMall, coupang
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<p>
<!-- 액션 시작 -->
<form name="frmAct">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="등록하기" onClick="popExtSiteJungsanData();">
		&nbsp;
		<input type="button" class="button" value="정산vs주문입력검토" onClick="popExtJungsanOrderCheck();">
		&nbsp;
		<input type="button" class="button" value="주문입력vs정산검토" onClick="popExtOrderJungsanCheck();">
		&nbsp;
		<input type="button" class="button" value="송장변경로그검토" onClick="popSongjangChangeCheck();">
		<!--
		&nbsp;
		<input type="button" class="button" value="정산확정일검토" onClick="popExtOrderJungsanFixdate();">
		-->
	</td>
	<td align="right">
		<% if (sellsite="lotteCom") and (RIGHT(fromDate,2)="01") then %>
			<% if (dateDiff("m",fromDate,now())<2) and (LEFT(fromDate,7)<>LEFT(now(),7)) then %>
			롯데닷컴 <%=LEFT(fromDate,7)%>월 변동수수료 <input type="text" name="lotteaddcommission" id="lotteaddcommission" value="" size="10" style="text-align:right"> <input type="button" value="등록" onClick="addLotteAddCommission(this,'<%=LEFT(fromDate,7)%>');">
			<% end if %>
		<% end if %>
	</td>
	<td align="right" width="300">
		<input type="button" class="button" value="API정산팝업" onClick="popAPIJungsanData();">
		<!--
	    <input type="text" name="extjdate" value="<%=extjdate%>" size="10" maxlength="8">
	    &nbsp;
	    <input type="button" class="button" value="SSG API 등록" onClick="popExtSiteJungsanByAPI('ssg',this);">

	    <input type="button" class="button" value="인터파크 API 등록" onClick="popExtSiteJungsanByAPI('interpark',this);">

	    <input type="button" class="button" value="lotteCom API 등록" onClick="popExtSiteJungsanByAPI('lottecom',this);">
	    &nbsp;
	    <input type="button" class="button" value="lotteiMall API 등록" onClick="popExtSiteJungsanByAPI('lotteimall',this);">
	    -->
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->
<p>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100" rowspan="2">
		<% if (grpgubun = "sellsite") then %>
			제휴몰
		<% else %>
			매출일자
		<% end if %>
	</td>
	<td colspan="6">상품</td>
	<td colspan="6">배송비</td>
	<td colspan="3">기타</td>
	<td colspan="6">합계</td>
	<td rowspan="2">주문 매핑<br>(총 / 상품)</td>
	<td rowspan="2">쿠폰 반영<br>(총 / 상품)</td>
	<td rowspan="2">주문수<br>(총 / 상품)</td>
	<td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>판매금액</td>
    <td>제휴부담<br>쿠폰</td>
    <td>텐텐부담<br>쿠폰</td>
	<td>매출금액</td>
	<td>수수료</td>
	<td>정산금액</td>

	<td>판매금액</td>
  	<td>제휴부담<br>쿠폰</td>
  	<td>텐텐부담<br>쿠폰</td>
	<td>매출금액</td>
	<td>수수료</td>
	<td>정산금액</td>

	<td>매출금액</td>
	<td>수수료</td>
	<td>정산금액</td>

	<td>판매금액</td>
  	<td>제휴부담<br>쿠폰</td>
  	<td>텐텐부담<br>쿠폰</td>
	<td>매출금액</td>
	<td>수수료</td>
	<td>정산금액</td>
</tr>

<% for i=0 to oCExtJungsan.FresultCount -1 %>
<%
totExtTenMeachulPriceProduct = totExtTenMeachulPriceProduct + oCExtJungsan.FItemList(i).FtotExtTenMeachulPriceProduct
totExtCommPriceProduct = totExtCommPriceProduct + oCExtJungsan.FItemList(i).FtotExtCommPriceProduct
totExtTenJungsanPriceProduct = totExtTenJungsanPriceProduct + oCExtJungsan.FItemList(i).FtotExtTenJungsanPriceProduct

totExtTenMeachulPriceDeliver = totExtTenMeachulPriceDeliver + oCExtJungsan.FItemList(i).FtotExtTenMeachulPriceDeliver
totExtCommPriceDeliver = totExtCommPriceDeliver + oCExtJungsan.FItemList(i).FtotExtCommPriceDeliver
totExtTenJungsanPriceDeliver = totExtTenJungsanPriceDeliver + oCExtJungsan.FItemList(i).FtotExtTenJungsanPriceDeliver

totExtTenMeachulPriceEtc = totExtTenMeachulPriceEtc + oCExtJungsan.FItemList(i).FtotExtTenMeachulPriceEtc
totExtCommPriceEtc = totExtCommPriceEtc + oCExtJungsan.FItemList(i).FtotExtCommPriceEtc
totExtTenJungsanPriceEtc = totExtTenJungsanPriceEtc + oCExtJungsan.FItemList(i).FtotExtTenJungsanPriceEtc

totExtTenMeachulPrice = totExtTenMeachulPrice + oCExtJungsan.FItemList(i).FtotExtTenMeachulPrice
totExtCommPrice = totExtCommPrice + oCExtJungsan.FItemList(i).FtotExtCommPrice
totExtTenJungsanPrice = totExtTenJungsanPrice + oCExtJungsan.FItemList(i).FtotExtTenJungsanPrice
totExtTenMiMapping = totExtTenMiMapping + oCExtJungsan.FItemList(i).FextMiMapping
totExtTenCount = totExtTenCount + oCExtJungsan.FItemList(i).FextRowCount

totExtTenMiMapping_C = totExtTenMiMapping_C + oCExtJungsan.FItemList(i).FextMiMapping_C
totExtTenCount_C = totExtTenCount_C + oCExtJungsan.FItemList(i).FextRowCount_C

totMiMappOrder    = totMiMappOrder + oCExtJungsan.FItemList(i).FMiMappOrder
totMiMappOrder_C  = totMiMappOrder_C + oCExtJungsan.FItemList(i).FMiMappOrder_C

totExtitemCostProduct   = totExtitemCostProduct + oCExtJungsan.FItemList(i).FtotExtitemCostProduct
totExtReducedPriceProduct   = totExtReducedPriceProduct + oCExtJungsan.FItemList(i).FtotExtReducedPriceProduct
totExtOwnCouponPriceProduct = totExtOwnCouponPriceProduct + oCExtJungsan.FItemList(i).FtotExtOwnCouponPriceProduct
totExtTenCouponPriceProduct = totExtTenCouponPriceProduct + oCExtJungsan.FItemList(i).FtotExtTenCouponPriceProduct

totExtitemCostDeliver   = totExtitemCostDeliver + oCExtJungsan.FItemList(i).FtotExtitemCostDeliver
totExtReducedPriceDeliver   = totExtReducedPriceDeliver + oCExtJungsan.FItemList(i).FtotExtReducedPriceDeliver
totExtOwnCouponPriceDeliver = totExtOwnCouponPriceDeliver + oCExtJungsan.FItemList(i).FtotExtOwnCouponPriceDeliver
totExtTenCouponPriceDeliver = totExtTenCouponPriceDeliver + oCExtJungsan.FItemList(i).FtotExtTenCouponPriceDeliver

totExtitemCost         = totExtitemCost + oCExtJungsan.FItemList(i).FtotExtitemCost
totExtReducedPrice         = totExtReducedPrice + oCExtJungsan.FItemList(i).FtotExtReducedPrice
totExtOwnCouponPrice       = totExtOwnCouponPrice + oCExtJungsan.FItemList(i).FtotExtOwnCouponPrice
totExtTenCouponPrice       = totExtTenCouponPrice + oCExtJungsan.FItemList(i).FtotExtTenCouponPrice
%>
<tr id="obj<%= i %>a" align="center" bgcolor="FFFFFF"  onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td>
		<% if (grpgubun = "sellsite") then %>
			<%= oCExtJungsan.FItemList(i).GetSellSiteName %>
		<% else %>
			<%= oCExtJungsan.FItemList(i).FextMeachulDate %>
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtitemCostProduct, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtOwnCouponPriceProduct, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenCouponPriceProduct, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenMeachulPriceProduct, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtCommPriceProduct, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenJungsanPriceProduct, 0) %></td>

	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtitemCostDeliver, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtOwnCouponPriceDeliver, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenCouponPriceDeliver, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenMeachulPriceDeliver, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtCommPriceDeliver, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenJungsanPriceDeliver, 0) %></td>

	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenMeachulPriceEtc, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtCommPriceEtc, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenJungsanPriceEtc, 0) %></td>

	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtitemCost, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtOwnCouponPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenCouponPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenMeachulPrice, 0) %>
	<% if (oCExtJungsan.FItemList(i).GetDiffMeachulPrice<>0) then %>
		<font color="red"><br>(<%= FormatNumber(oCExtJungsan.FItemList(i).GetDiffMeachulPrice, 0) %>)</font>
	<% end if %>
	</td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtCommPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FtotExtTenJungsanPrice, 0) %></td>
	<td align="right">
	    <%= FormatNumber(oCExtJungsan.FItemList(i).FMiMappOrder, 0) %>
	    / <%= FormatNumber(oCExtJungsan.FItemList(i).FMiMappOrder_C, 0) %>
	</td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextMiMapping, 0) %>
	    / <%= FormatNumber(oCExtJungsan.FItemList(i).FextMiMapping_C, 0) %>
	</td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextRowCount, 0) %>
	    / <%= FormatNumber(oCExtJungsan.FItemList(i).FextRowCount_C, 0) %>
	</td>
	<td>
		<% if (oCExtJungsan.FItemList(i).GetDiffJungsanPrice<>0) then %>
		<font color="red"><%= FormatNumber(oCExtJungsan.FItemList(i).GetDiffJungsanPrice, 0) %></font>
		<br />
		<% end if %>
		<% If grpgubun = "extMeachulDate" AND sellsite <> "" AND oCExtJungsan.FItemList(i).FextMiMapping > 0 Then %>
			<% If oCExtJungsan.FItemList(i).FextMeachulDate > "2017-02-28" Then %>
			<span id="mappingbtn<%= i %>"><input type="button" value="쿠폰반영" onClick="jsMapping('<%=CHKIIF(sellsite<>"",sellsite,oCExtJungsan.FItemList(i).GetSellSiteName)%>','<%= oCExtJungsan.FItemList(i).FextMeachulDate %>','<%= i %>');"></span>
			<span id="mappinging<%= i %>" style="display:none;"><img src="http://fiximage.10x10.co.kr/icons/loading16.gif" style="width:16px;height:16px;" /></span>
			<% End If %>


		<% End If %>

		<% if grpgubun = "extMeachulDate" AND sellsite <> "" AND dateDiff("D",oCExtJungsan.FItemList(i).FextMeachulDate,now())<5 then %>
			<input type="button" value="삭제" onClick="jsDelMeachul('<%=CHKIIF(sellsite<>"",sellsite,oCExtJungsan.FItemList(i).GetSellSiteName)%>','<%= oCExtJungsan.FItemList(i).FextMeachulDate %>')">
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="FFFFFF">
	<td></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtitemCostProduct, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtOwnCouponPriceProduct, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenCouponPriceProduct, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenMeachulPriceProduct, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtCommPriceProduct, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenJungsanPriceProduct, 0) %></strong></td>

	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtitemCostDeliver, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtOwnCouponPriceDeliver, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenCouponPriceDeliver, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenMeachulPriceDeliver, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtCommPriceDeliver, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenJungsanPriceDeliver, 0) %></strong></td>

	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenMeachulPriceEtc, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtCommPriceEtc, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenJungsanPriceEtc, 0) %></strong></td>

	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtitemCost, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtOwnCouponPrice, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenCouponPrice, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenMeachulPrice, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtCommPrice, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenJungsanPrice, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totMiMappOrder, 0) %> / <%= FormatNumber(totMiMappOrder_C, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenMiMapping, 0) %> / <%= FormatNumber(totExtTenMiMapping_C, 0) %></strong></td>
	<td align="right" bgcolor="#E6B9B8"><strong><%= FormatNumber(totExtTenCount, 0) %> / <%= FormatNumber(totExtTenCount_C, 0) %></strong></td>
	<td></td>
</tr>
</table>

<%
set oCExtJungsan = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
