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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftStatisticCls.asp" -->
<%
	Dim i, sDate, eDate, cStat, vTab, vArrTalk, vArrDay, vArrShop
	Dim vTotPC, vTotMob, vTotTalkPC, vTotTalkMob, vTotDayPC, vTotDayMob, vTotShopPC, vTotShopMob
	Dim vTW5, vTW0, vTW1, vTW2, vTW3, vTW4, vTW7, vTM5, vTM0, vTM1, vTM2, vTM3, vTM4, vTM7
	Dim vDW5, vDW0, vDW1, vDW2, vDW3, vDW4, vDW7, vDM5, vDM0, vDM1, vDM2, vDM3, vDM4, vDM7
	Dim vSW5, vSW0, vSW1, vSW2, vSW3, vSW4, vSW7, vTot5, vTot0, vTot1, vTot2, vTot3, vTot4, vTot7
	sDate = NullFillWith(request("sDate"),DateAdd("d",-10,date()))
	eDate = NullFillWith(request("eDate"),date())
	vTab = NullFillWith(request("tab"),1)
	
	
	SET cStat = New CgiftStat_list
	If vTab = "1" Then
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		cStat.sbStatDaily
	ElseIf vTab = "2" Then
		cStat.FRectGubun = "talk"
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		vArrTalk = cStat.fnStatUserLevel
		
		cStat.FRectGubun = "day"
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		vArrDay = cStat.fnStatUserLevel
		
		cStat.FRectGubun = "shop"
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		vArrShop = cStat.fnStatUserLevel
	End If
%>
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
<table height="70">
<tr>
	<td width="90px"><a href="javascript:goTab('1');"><span style="font-size:11pt;<%=CHKIIF(vTab="1","font-weight:bold;text-decoration:underline;","")%>">[일별 조회]</span></a></td>
	<td style="padding:10px;"><a href="javascript:goTab('2');"><span style="font-size:11pt;<%=CHKIIF(vTab="2","font-weight:bold;text-decoration:underline;","")%>">[등급별 조회]</span></a></td>
</tr>
<tr>
	<td colspan="2">
		<span style="color:blue;font-size:9pt;">
			※ <strong>등급별 조회 데이터</strong>는 회원DB를 연결하는 데이터로 <strong>탈퇴한 회원의 데이터는 조회하지 않습니다.</strong> 그러므로 <strong>약간의 수치 차이</strong>가 있을 수 있습니다.
		</span>
	</td>
</tr>
</table>

<!-- 상단 검색폼 시작 -->
<form name="frm1" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="tab" value="<%=vTab%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="700" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="60">
	<td width="60" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		&nbsp;&nbsp;
        <input id="sDate" name="sDate" value="<%=sDate%>" class="text" size="10" maxlength="10" readonly /> <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "sDate", trigger    : "sDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		&nbsp;&nbsp;~&nbsp;&nbsp;
        <input id="eDate" name="eDate" value="<%=eDate%>" class="text" size="10" maxlength="10" readonly /> <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "eDate", trigger    : "eDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="80" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검 색" style="width:70px;height:50px;">
	</td>
</tr>
</table>
</form>
<br />
<input type="button" value="엑셀다운" onClick="goExcelDown()">
<br /><br />
<% If vTab = "1" Then %>
<table width="600" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td height="20">기간</td>
    <td>구분</td>
    <td>기프트 톡</td>
    <td>기프트 데이</td>
    <td>기프트 샵</td>
    <td>합계</td>
</tr>
<%
	for i=0 to cStat.FResultCount - 1
	
	vTotTalkPC	= vTotTalkPC + cStat.FItemList(i).FTalkWeb
	vTotTalkMob = vTotTalkMob + cStat.FItemList(i).FTalkMob
	vTotDayPC	= vTotDayPC + cStat.FItemList(i).FDayWeb
	vTotDayMob	= vTotDayMob + cStat.FItemList(i).FDayMob
	vTotShopPC	= vTotShopPC + cStat.FItemList(i).FShopWeb
	vTotShopMob	= ""
	
	vTotPC = cStat.FItemList(i).FTalkWeb + cStat.FItemList(i).FDayWeb + cStat.FItemList(i).FShopWeb
	vTotMob = cStat.FItemList(i).FTalkMob + cStat.FItemList(i).FDayMob
%>
	<tr bgcolor="#FFFFFF">
		<td align="center" rowspan="2"><%= cStat.FItemList(i).FDate %></td>
		<td height="20" style="padding:5px;">PC</td>
		<td style="padding:5px;"><%= cStat.FItemList(i).FTalkWeb %></td>
		<td style="padding:5px;"><%= cStat.FItemList(i).FDayWeb %></td>
		<td style="padding:5px;"><%= cStat.FItemList(i).FShopWeb %></td>
		<td style="padding:5px;"><%= vTotPC %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td height="20" style="padding:5px;">모바일</td>
		<td style="padding:5px;"><%= cStat.FItemList(i).FTalkMob %></td>
		<td style="padding:5px;"><%= cStat.FItemList(i).FDayMob %></td>
		<td style="padding:5px;"><%= cStat.FItemList(i).FShopMob %></td>
		<td style="padding:5px;"><%= vTotMob %></td>
	</tr>
<%
	
	
	next
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td height="20"></td>
    <td>구분</td>
    <td>기프트 톡</td>
    <td>기프트 데이</td>
    <td>기프트 샵</td>
    <td>합계</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" rowspan="2" bgcolor="#E6B9B8">합계</td>
	<td height="20" style="padding:5px;" bgcolor="#E6B9B8">PC</td>
	<td style="padding:5px;" bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotTalkPC,0) %></strong></td>
	<td style="padding:5px;" bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotDayPC,0) %></strong></td>
	<td style="padding:5px;" bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotShopPC,0) %></strong></td>
	<td style="padding:5px;" bgcolor="#E6B9B8"><strong><%= FormatNumber((vTotTalkPC + vTotDayPC + vTotShopPC),0) %></strong></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="20" style="padding:5px;" bgcolor="#E6B9B8">모바일</td>
	<td style="padding:5px;" bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotTalkMob,0) %></strong></td>
	<td style="padding:5px;" bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotDayMob,0) %></strong></td>
	<td style="padding:5px;" bgcolor="#E6B9B8"></td>
	<td style="padding:5px;" bgcolor="#E6B9B8"><strong><%= FormatNumber((vTotTalkMob + vTotDayMob),0) %></strong></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">총참여자</td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber((vTotTalkPC + vTotTalkMob),0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber((vTotDayPC + vTotDayMob),0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTotShopPC,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber((vTotTalkPC + vTotTalkMob + vTotDayPC + vTotDayMob + vTotShopPC),0) %></strong></td>
</tr>
</table>
<% Else
	vTW5 = fnArrCount(vArrTalk,"w",5)
	vTW0 = fnArrCount(vArrTalk,"w",0)
	vTW1 = fnArrCount(vArrTalk,"w",1)
	vTW2 = fnArrCount(vArrTalk,"w",2)
	vTW3 = fnArrCount(vArrTalk,"w",3)
	vTW4 = fnArrCount(vArrTalk,"w",4)
	vTW7 = fnArrCount(vArrTalk,"w",7)
	vTM5 = fnArrCount(vArrTalk,"m",5)
	vTM0 = fnArrCount(vArrTalk,"m",0)
	vTM1 = fnArrCount(vArrTalk,"m",1)
	vTM2 = fnArrCount(vArrTalk,"m",2)
	vTM3 = fnArrCount(vArrTalk,"m",3)
	vTM4 = fnArrCount(vArrTalk,"m",4)
	vTM7 = fnArrCount(vArrTalk,"m",7)
	vDW5 = fnArrCount(vArrDay,"W",5)
	vDW0 = fnArrCount(vArrDay,"W",0)
	vDW1 = fnArrCount(vArrDay,"W",1)
	vDW2 = fnArrCount(vArrDay,"W",2)
	vDW3 = fnArrCount(vArrDay,"W",3)
	vDW4 = fnArrCount(vArrDay,"W",4)
	vDW7 = fnArrCount(vArrDay,"W",7)
	vDM5 = fnArrCount(vArrDay,"M",5)
	vDM0 = fnArrCount(vArrDay,"M",0)
	vDM1 = fnArrCount(vArrDay,"M",1)
	vDM2 = fnArrCount(vArrDay,"M",2)
	vDM3 = fnArrCount(vArrDay,"M",3)
	vDM4 = fnArrCount(vArrDay,"M",4)
	vDM7 = fnArrCount(vArrDay,"M",7)
	vSW5 = fnArrCount(vArrShop,"w",5)
	vSW0 = fnArrCount(vArrShop,"w",0)
	vSW1 = fnArrCount(vArrShop,"w",1)
	vSW2 = fnArrCount(vArrShop,"w",2)
	vSW3 = fnArrCount(vArrShop,"w",3)
	vSW4 = fnArrCount(vArrShop,"w",4)
	vSW7 = fnArrCount(vArrShop,"w",7)
	
	vTot5 = vTW5 + vTM5 + vDW5 + vDM5 + vSW5
	vTot0 = vTW0 + vTM0 + vDW0 + vDM0 + vSW0
	vTot1 = vTW1 + vTM1 + vDW1 + vDM1 + vSW1
	vTot2 = vTW2 + vTM2 + vDW2 + vDM2 + vSW2
	vTot3 = vTW3 + vTM3 + vDW3 + vDM3 + vSW3
	vTot4 = vTW4 + vTM4 + vDW4 + vDM4 + vSW4
	vTot7 = vTW7 + vTM7 + vDW7 + vDM7 + vSW7

	vTotTalkPC	= vTW5 + vTW0 + vTW1 + vTW2 + vTW3 + vTW4 + vTW7
	vTotTalkMob	= vTM5 + vTM0 + vTM1 + vTM2 + vTM3 + vTM4 + vTM7
	vTotDayPC	= vDW5 + vDW0 + vDW1 + vDW2 + vDW3 + vDW4 + vDW7
	vTotDayMob	= vDM5 + vDM0 + vDM1 + vDM2 + vDM3 + vDM4 + vDM7
	vTotShopPC	= vSW5 + vSW0 + vSW1 + vSW2 + vSW3 + vSW4 + vSW7
%>
<table width="600" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td height="50">기간</td>
    <td>구분</td>
    <td class="<%=getUserLevelCSS(5)%>"><strong>오렌지</strong></td>
    <td class="<%=getUserLevelCSS(0)%>"><strong>옐로우</strong></td>
    <td class="<%=getUserLevelCSS(1)%>"><strong>그린</strong></td>
    <td class="<%=getUserLevelCSS(2)%>"><strong>블루</strong></td>
    <td class="<%=getUserLevelCSS(3)%>"><strong>VIP<br />실버</strong></td>
    <td class="<%=getUserLevelCSS(4)%>"><strong>VIP<br />골드</strong></td>
    <td class="<%=getUserLevelCSS(7)%>"><strong>스텝</strong></td>
    <td>합계</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" rowspan="2">기프트톡</td>
	<td height="20" style="padding:5px;">PC</td>
	<td style="padding:5px;"><%= FormatNumber(vTW5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW7,0) %></td>
	<td style="padding:5px;"><%= vTotTalkPC %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="20" style="padding:5px;">모바일</td>
	<td style="padding:5px;"><%= FormatNumber(vTM5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM7,0) %></td>
	<td style="padding:5px;"><%= vTotTalkMob %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" rowspan="2">기프트데이</td>
	<td height="20" style="padding:5px;">PC</td>
	<td style="padding:5px;"><%= FormatNumber(vDW5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW7,0) %></td>
	<td style="padding:5px;"><%= vTotDayPC %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="20" style="padding:5px;">모바일</td>
	<td style="padding:5px;"><%= FormatNumber(vDM5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM7,0) %></td>
	<td style="padding:5px;"><%= vTotDayMob %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">기프트샵</td>
	<td height="20" style="padding:5px;">PC</td>
	<td style="padding:5px;"><%= FormatNumber(vSW5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW7,0) %></td>
	<td style="padding:5px;"><%= vTotShopPC %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">총참여자</td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTot5,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTot0,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTot1,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTot2,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTot3,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTot4,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTot7,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber((vTotTalkPC + vTotTalkMob + vTotDayPC + vTotDayMob + vTotShopPC),0) %></strong></td>
</tr>
</table>
<% End If %>

<% SET cStat = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->