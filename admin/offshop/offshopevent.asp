<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/eventAppReport.asp"-->
<%
Dim sSdate, sEdate, mode, offEvent, eventNo, arrList, i, appRunUser, buyprice, appRunDay
Dim TotalCouponCnt, TotalAlreayUserCnt, TotalNewJoinCnt, TotalDelUserCnt, TermCouponCnt, TermAlreayUserCnt, TermJoinCnt, TermDelUserCnt, TotalOnBuyUserCnt, TermOnBuyUserCnt, TotalOnAvgDays, TermOnAvgDays, TotalActiveUserTotalSum, TermActiveUserTotalSum
Dim TotalactiveUserCnt, TermActiveUserCnt
Dim TotalVisitCount, TotalJumunCount, TermVisitCount, TermJumunCount
sSdate			= requestCheckVar(request("iSD"),10)
sEdate			= requestCheckVar(request("iED"),10)
eventNo			= requestCheckVar(request("eventNo"),10)
appRunUser		= requestCheckVar(request("appRunUser"),1)
buyprice		= request("buyprice")
appRunDay		= request("appRunDay")

If sSdate = "" Then sSdate = DateSerial(Year(Now()), Month(Now()), 1)
If sEdate = "" Then sEdate = Date()
If eventNo = "" Then eventNo = 3
If appRunUser = "" Then appRunUser = 0
If buyprice = "" Then buyprice = 100
If appRunDay = "" Then appRunDay = 1
	
SET offEvent = new COffEvent
	offEvent.FRectSdate			= sSdate
	offEvent.FRectEdate			= sEdate
	offEvent.FRectEventNo		= eventNo
	offEvent.FRectBuyprice		= buyprice
	offEvent.FRectAppRunUser	= appRunUser
	offEvent.FRectAppRunDay		= appRunDay
	arrList = offEvent.fnOffEventReport

	TotalCouponCnt			= offEvent.FTotalCouponCnt
	TotalAlreayUserCnt		= offEvent.FAlReadyUserCnt
	TotalNewJoinCnt 		= offEvent.FUserJoinCnt
	TotalDelUserCnt 		= offEvent.FDelUserCnt
	TotalOnBuyUserCnt	 	= offEvent.FOnBuyUserCnt
	TotalOnAvgDays	 		= offEvent.FOnAvgDays
	TotalactiveUserCnt		= offEvent.FActiveUserCnt
	TotalActiveUserTotalSum	= offEvent.FActiveUserTotalSum

	TermCouponCnt			= offEvent.FTermCouponCnt
	TermAlreayUserCnt		= offEvent.FTermAlreayUserCnt
	TermJoinCnt				= offEvent.FTermJoinCnt
	TermDelUserCnt 			= offEvent.FTermDelUserCnt
	TermOnBuyUserCnt 		= offEvent.FTermOnBuyUserCnt
	TermOnAvgDays 			= offEvent.FTermOnAvgDays
	TermActiveUserCnt 		= offEvent.FTermActiveUserCnt
	TermActiveUserTotalSum	= offEvent.FTermActiveUserTotalSum

	TotalVisitCount 	= offEvent.FTotalVisitCount
	TotalJumunCount	 	= offEvent.FTotalJumunCount
	TermVisitCount		= offEvent.FTermVisitCount
	TermJumunCount		= offEvent.FTermJumunCount
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
function pop_setup(){
	var pCM = window.open("/admin/offshop/popSearchsetup.asp?buyprice=<%=buyprice%>&appRunUser=<%=appRunUser%>&appRunDay=<%=appRunDay%>","popsetup","width=600,height=300,scrollbars=yes,resizable=yes");
	pCM.focus();
}

function pop_UserReport(){
	var pCM2 = window.open("/admin/offshop/popUserReport.asp?eventNo=<%=eventNo%>&sDate=<%=sSdate%>&eDate=<%=sEdate%>","popUserReport","width=1000,height=680,scrollbars=yes,resizable=yes");
	pCM2.focus();
}

function fnAddShop(v){
	$("#shopid").val(v);
	$.ajax({
		type: "POST",
		url: "/admin/offshop/offshopevent_ajax.asp",
		data: $("#frm").serialize(),
		dataType: "text",
		async: false,
		success : function(result){
		    $("#shopTBL").empty().html(result);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}
function fnAddSumTerm(){
	$.ajax({
		type: "POST",
		url: "/admin/offshop/offshopevent_sumajax.asp",
		data: $("#frm").serialize(),
		dataType: "text",
		async: false,
		success : function(result){
		    $("#shopTBL").empty().html(result);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" id="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="shopid" id="shopid">
<input type="hidden" name="appRunUser" id="appRunUser" value="<%=appRunUser%>">
<input type="hidden" name="buyprice" id="buyprice" value="<%=buyprice%>">
<input type="hidden" name="appRunDay" id="appRunDay" value="<%=appRunDay%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		기간 : 
		<input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "iSD", trigger    : "iSD_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "iED", trigger    : "iED_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
<!--
		&nbsp;
		<select class="select" name="dayMode">
			<option>일별</option>
			<option>주별</option>
			<option>월별</option>
			<option>연별</option>	
		</select>
-->
		<br /><br />
		이벤트 : 
		<select class="select" name="eventNo">
			<option value="1" <%= chkiif(eventNo=1, "selected", "") %> >앱 설치 이벤트</option>
			<option value="2" <%= chkiif(eventNo=2, "selected", "") %> >아트토이컬쳐</option>
			<option value="3" <%= chkiif(eventNo=3, "selected", "") %> >사은품 증정 이벤트</option>
			<option value="4" <%= chkiif(eventNo=4, "selected", "") %> >신촌물총축제</option>
			<option value="5" <%= chkiif(eventNo=5, "selected", "") %> >2018 BML</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" bgcolor="#FFFFFF">
	<td colspan="15">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<strong>매장별 통계</strong>
			</td>
			<td align="right">
				<input type="button" class="button" value="회원통계 조회" onclick="pop_UserReport();">
				&nbsp;&nbsp;
				<input type="button" class="button" value="통계 설정" onclick="pop_setup();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="6.5%"></td>
	<td width="6.5%">매장유입객수(명)</td>
	<td width="6.5%">매장구매(건)</td>
	<td width="6.5%">쿠폰사용(건)</td>
	<td width="6.5%">유입객수<br>대비(%)</td>
	<td width="6.5%">구매건수<br>대비(%)</td>
	<td width="6.5%">온라인 구매<br>전환총금액(원)</td>
	<td width="6.5%">기존회원(건)</td>
	<td width="6.5%">회원가입(건)</td>
	<td width="6.5%">탈퇴(건)</td>
	<td width="6.5%">회원 전환률(%)</td>
	<td width="6.5%">온라인 구매<br />전환(명)</td>
	<td width="6.5%">온라인 구매<br />평균 소요기간(일)</td>
	<td width="6.5%">유저 활성화(건)</td>
	<td width="6.5%">유저 활성화률(%)</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td>누적 전체</td>
	<td><%= FormatNumber(TotalVisitCount, 0) %></td>
	<td><%= FormatNumber(TotalJumunCount, 0) %></td>
	<td><%= FormatNumber(TotalCouponCnt, 0) %></td>
	<td>
	<%
		If TotalVisitCount <> 0 Then
			response.write Round(TotalCouponCnt / TotalVisitCount * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td>
	<%
		If TotalJumunCount <> 0 Then
			response.write Round(TotalCouponCnt / TotalJumunCount * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(TotalActiveUserTotalSum,0) %></td>
	<td><%= FormatNumber(TotalAlreayUserCnt,0) %></td>
	<td><%= FormatNumber(TotalNewJoinCnt,0) %></td>
	<td><%= FormatNumber(TotalDelUserCnt,0) %></td>
	<td>
	<%
		If TotalCouponCnt <> 0 Then
			response.write Round(TotalNewJoinCnt / TotalCouponCnt * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(TotalOnBuyUserCnt,0) %></td>
	<td>
	<%
		If TotalOnBuyUserCnt <> 0 Then
			response.write Round(TotalOnAvgDays / TotalOnBuyUserCnt, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(TotalactiveUserCnt,0) %></td>
	<td>
	<%
		If TotalCouponCnt <> 0 Then
			response.write Round((TotalactiveUserCnt) / TotalCouponCnt * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer;" onclick="fnAddSumTerm();">
	<td>기간 전체</td>
	<td><%= FormatNumber(TermVisitCount, 0) %></td>
	<td><%= FormatNumber(TermJumunCount, 0) %></td>
	<td><%= FormatNumber(TermCouponCnt, 0) %></td>
	<td>
	<%
		If TermVisitCount <> 0 Then
			response.write Round(TermCouponCnt / TermVisitCount * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td>
	<%
		If TermJumunCount <> 0 Then
			response.write Round(TermCouponCnt / TermJumunCount * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(TermActiveUserTotalSum, 0) %></td>
	<td><%= FormatNumber(TermAlreayUserCnt, 0) %></td>
	<td><%= FormatNumber(TermJoinCnt, 0) %></td>
	<td><%= FormatNumber(TermDelUserCnt, 0) %></td>
	<td>
	<%
		If TermCouponCnt <> 0 Then
			response.write Round(TermJoinCnt / TermCouponCnt * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(TermOnBuyUserCnt, 0) %></td>
	<td>
	<%
		If TermOnBuyUserCnt <> 0 Then
			response.write Round(TermOnAvgDays / TermOnBuyUserCnt, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(TermActiveUserCnt, 0) %></td>
	<td>
	<%
		If TermCouponCnt <> 0 Then
			response.write Round((TermActiveUserCnt) / TermCouponCnt * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
</tr>
<tr align="center" bgcolor="#D2D2D2">
	<td colspan="15"></td>
</tr>
<% If IsArray(arrList) Then %>
<% For i=0 To Ubound(arrList, 2) %>
<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer;" onclick="fnAddShop('<%= arrList(0, i) %>');">
	<td><%= arrList(1, i) %></td>
	<td><%= FormatNumber(arrList(9, i), 0) %></td>
	<td><%= FormatNumber(arrList(10, i), 0) %></td>
	<td><%= FormatNumber(arrList(2, i), 0) %></td>
	<td>
	<%
		If arrList(9, i) <> 0 Then
			response.write Round(arrList(2, i) / arrList(9, i) * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td>
	<%
		If arrList(10, i) <> 0 Then
			response.write Round(arrList(2, i) / arrList(10, i) * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(arrList(11, i), 0) %></td>
	<td><%= FormatNumber(arrList(4, i), 0) %></td>
	<td><%= FormatNumber(arrList(3, i), 0) %></td>
	<td><%= FormatNumber(arrList(5, i), 0) %></td>
	<td>
	<%
		If arrList(2, i) <> 0 Then
			response.write Round(arrList(3, i) / arrList(2, i) * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(arrList(6, i), 0) %></td>
	<td>
	<%
		If arrList(6, i) <> 0 Then
			response.write Round(arrList(7, i) / arrList(6, i), 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(arrList(8, i), 0) %></td>
	<td>
	<%
		If arrList(2, i) <> 0 Then
			response.write Round((arrList(8, i)) / arrList(2, i) * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
</tr>
<% Next %>
<% End If %>
</table>
<div id="shopTBL"></div>
<% SET offEvent = nothing %>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->