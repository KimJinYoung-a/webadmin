<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// 즐겨찾기
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script language='javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;

	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "즐겨찾기에서 제외하시겠습니까?";
	} else {
		msg = "즐겨찾기에 추가하시겠습니까?";
	}

	ret = confirm(msg);

	if (ret) {
		frm.submit();
	}
}
</script>
<% if session("sslgnMethod")<>"S" then %>
<!-- USB키 처리 시작 (2008.06.23;허진원) -->
<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
<script language="javascript" src="/js/check_USBToken.js"></script>
<!-- USB키 처리 끝 -->
<% end if %>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function goTab(a){
	$('input[name="tab"]').val(a);
	frm1.submit();
}

function goExcelDown(){
	frm1.action = "exceldown.asp";
	frm1.submit();
	
	frm1.action = "";
}
</script>
</head>
<body <% if session("sslgnMethod")<>"S" then %>onload="checkUSBKey()"<% end if %>>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftStatisticCls.asp" -->
<%
	Dim i, sDate, eDate, cStat, vTab, vArrTalk, vArrDay, vArrShop
	Dim vTotPC, vTotMob, vTotApp, vTotPC1, vTotPC2, vTotMob1, vTotMob2, vTotApp1, vTotApp2, vTotTotalPC, vTotTotalMob, vTotTotalApp

	sDate = NullFillWith(request("sDate"),DateAdd("d",-10,date()))
	eDate = NullFillWith(request("eDate"),date())
	vTab = NullFillWith(request("tab"),1)
	
	
	SET cStat = New CgiftStat_list
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		cStat.sbPojangStatDaily
%>

<div class="contSectFix scrl">
	<div class="contHead">
		<div class="simpleDesp">
			- 주문 결제 상태 및 배송 상황에 따라 약간의 수치 차이가 있을 수 있습니다.
		</div>
	</div>

	<!-- 상단 검색폼 시작 -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="tab" value="<%=vTab%>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li>
					<label class="formTit" for="term1">기간 :</label>
					<input type="text" class="formTxt" id="sDate" name="sDate" value="<%=sDate%>" style="width:100px" placeholder="시작일" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="sDate_trigger" alt="달력으로 검색" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "sDate", trigger    : "sDate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					~
					<input type="text" class="formTxt" id="eDate" name="eDate" value="<%=eDate%>" style="width:100px" placeholder="종료일" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="eDate_trigger" alt="달력으로 검색" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "eDate", trigger    : "eDate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="검색" />
	</div>
	</form>

	<% If vTab = "1" Then %>
	<div class="pad20">
		<div class="overHidden">
			<div class="ftRt">
				<p class="btn2 cBk1 ftLt"><a href="javascript:goExcelDown();"><span class="eIcon down"><em class="fIcon xls">데이터저장</em></span></a></p>
			</div>
		</div>

		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th rowspan="2"><div>날짜</div></th>
					<th rowspan="2"><div>채널</div></th>
					<th colspan="2"><div>포장 수량</div></th>
					<th rowspan="2"><div>포장수량 합계</div></th>
					<th rowspan="2"><div>가격</div></th>
				</tr>
				<tr>
					<th><div>상품 1개</div></th>
					<th><div>상품 2개 이상</div></th>
				</tr>
				</thead>
				<tbody>
				<%
					for i=0 to cStat.FResultCount - 1
					
					vTotPC1		= vTotPC1 + cStat.FItemList(i).FPW1
					vTotPC2 	= vTotPC2 + cStat.FItemList(i).FPW2
					vTotMob1	= vTotMob1 + cStat.FItemList(i).FPM1
					vTotMob2	= vTotMob2 + cStat.FItemList(i).FPM2
					vTotApp1	= vTotApp1 + cStat.FItemList(i).FPA1
					vTotApp2	= vTotApp2 + cStat.FItemList(i).FPA2
					
					vTotPC		= cStat.FItemList(i).FPW1 + cStat.FItemList(i).FPW2
					vTotMob 	= cStat.FItemList(i).FPM1 + cStat.FItemList(i).FPM2
					vTotApp 	= cStat.FItemList(i).FPA1 + cStat.FItemList(i).FPA2
					
					vTotTotalPC		= vTotTotalPC + vTotPC
					vTotTotalMob	= vTotTotalMob + vTotMob
					vTotTotalApp	= vTotTotalApp + vTotApp
				%>
					<tr>
						<td rowspan="3"><%= cStat.FItemList(i).FDate %></td>
						<td>PC(W)</td>
						<td><%= cStat.FItemList(i).FPW1 %></td>
						<td><%= cStat.FItemList(i).FPW2 %></td>
						<td><%= vTotPC %></td>
						<td><%= FormatNumber((vTotPC*2000),0) %></td>
					</tr>
					<tr>
						<td>모바일웹(M)</td>
						<td><%= cStat.FItemList(i).FPM1 %></td>
						<td><%= cStat.FItemList(i).FPM2 %></td>
						<td><%= vTotMob %></td>
						<td><%= FormatNumber((vTotMob*2000),0) %></td>
					</tr>
					<tr>
						<td>모바일앱(A)</td>
						<td><%= cStat.FItemList(i).FPA1 %></td>
						<td><%= cStat.FItemList(i).FPA2 %></td>
						<td><%= vTotApp %></td>
						<td><%= FormatNumber((vTotApp*2000),0) %></td>
					</tr>
				</tbody>
				<%
					next
				%>
				<tfoot>
					<tr>
						<td colspan="2" class="bgGy1"><strong>합계</strong></td>
						<td class="bgGy1"><strong><%= FormatNumber((vTotPC1 + vTotMob1 + vTotApp1),0) %></strong></td>
						<td class="bgGy1"><strong><%= FormatNumber((vTotPC2 + vTotMob2 + vTotApp2),0) %></strong></td>
						<td class="bgGy1"><strong><%= FormatNumber((vTotTotalPC + vTotTotalMob + vTotTotalApp),0) %></strong></td>
						<td class="bgGy1"><strong><%= FormatNumber(((vTotTotalPC + vTotTotalMob + vTotTotalApp)*2000),0) %></strong></td>
					</tr>
				</tfoot>
			</table>
		</div>
	</div>
</div>

<% End If %>

<% SET cStat = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->