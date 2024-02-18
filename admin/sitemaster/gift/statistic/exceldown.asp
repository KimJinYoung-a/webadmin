<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 기프트 통계
' Hieditor : 2015.05.27 강준구 생성
'			 2016.07.19 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftStatisticCls.asp" -->
<%
	Dim i, sDate, eDate, cStat, vTab, vArrTalk, vArrDay, vArrShop
	Dim vTotPC, vTotMob, vTotTalkPC, vTotTalkMob, vTotDayPC, vTotDayMob, vTotShopPC, vTotShopMob
	Dim vTW5, vTW0, vTW1, vTW2, vTW3, vTW4, vTW6, vTW7, vTM5, vTM0, vTM1, vTM2, vTM3, vTM4, vTM6, vTM7
	Dim vDW5, vDW0, vDW1, vDW2, vDW3, vDW4, vDW6, vDW7, vDM5, vDM0, vDM1, vDM2, vDM3, vDM4, vDM6, vDM7
	Dim vSW5, vSW0, vSW1, vSW2, vSW3, vSW4, vSW6, vSW7, vTot5, vTot0, vTot1, vTot2, vTot3, vTot4, vTot6, vTot7
	sDate = NullFillWith(request("sDate"),DateAdd("d",-10,date()))
	eDate = NullFillWith(request("eDate"),date())
	vTab = NullFillWith(request("tab"),1)
	

	Response.ContentType = "application/x-msexcel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition", "attachment;filename=gift_" & CHKIIF(vTab="1","daily","userlevel") & "_" & sDate & "to" & eDate & ".xls"

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

<html>
<body>
<% If vTab = "1" Then %>
<table border="1">
	<tr>
	    <td>기간</td>
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
		<tr>
			<td rowspan="2"><%= cStat.FItemList(i).FDate %></td>
			<td>PC</td>
			<td><%= cStat.FItemList(i).FTalkWeb %></td>
			<td><%= cStat.FItemList(i).FDayWeb %></td>
			<td><%= cStat.FItemList(i).FShopWeb %></td>
			<td><%= vTotPC %></td>
		</tr>
		<tr>
			<td>모바일</td>
			<td><%= cStat.FItemList(i).FTalkMob %></td>
			<td><%= cStat.FItemList(i).FDayMob %></td>
			<td><%= cStat.FItemList(i).FShopMob %></td>
			<td><%= vTotMob %></td>
		</tr>
	<%
		next
	%>
		<tr>
			<td rowspan="2" bgcolor="#E6B9B8">합계</td>
			<td bgcolor="#E6B9B8">PC</td>
			<td bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotTalkPC,0) %></strong></td>
			<td bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotDayPC,0) %></strong></td>
			<td bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotShopPC,0) %></strong></td>
			<td bgcolor="#E6B9B8"><strong><%= FormatNumber((vTotTalkPC + vTotDayPC + vTotShopPC),0) %></strong></td>
		</tr>
		<tr>
			<td bgcolor="#E6B9B8">모바일</td>
			<td bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotTalkMob,0) %></strong></td>
			<td bgcolor="#E6B9B8"><strong><%= FormatNumber(vTotDayMob,0) %></strong></td>
			<td bgcolor="#E6B9B8"></td>
			<td bgcolor="#E6B9B8"><strong><%= FormatNumber((vTotTalkMob + vTotDayMob),0) %></strong></td>
		</tr>
		<tr>
			<td colspan="2" bgcolor="<%= adminColor("tabletop") %>">총참여자</td>
			<td bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber((vTotTalkPC + vTotTalkMob),0) %></strong></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber((vTotDayPC + vTotDayMob),0) %></strong></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTotShopPC,0) %></strong></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber((vTotTalkPC + vTotTalkMob + vTotDayPC + vTotDayMob + vTotShopPC),0) %></strong></td>
		</tr>
</table>

<% Else
	vTW5 = fnArrCount(vArrTalk,"w",5)
	vTW0 = fnArrCount(vArrTalk,"w",0)
	vTW1 = fnArrCount(vArrTalk,"w",1)
	vTW2 = fnArrCount(vArrTalk,"w",2)
	vTW3 = fnArrCount(vArrTalk,"w",3)
	vTW4 = fnArrCount(vArrTalk,"w",4)
	vTW6 = fnArrCount(vArrTalk,"w",6)
	vTW7 = fnArrCount(vArrTalk,"w",7)
	vTM5 = fnArrCount(vArrTalk,"m",5)
	vTM0 = fnArrCount(vArrTalk,"m",0)
	vTM1 = fnArrCount(vArrTalk,"m",1)
	vTM2 = fnArrCount(vArrTalk,"m",2)
	vTM3 = fnArrCount(vArrTalk,"m",3)
	vTM4 = fnArrCount(vArrTalk,"m",4)
	vTM6 = fnArrCount(vArrTalk,"m",6)
	vTM7 = fnArrCount(vArrTalk,"m",7)
	vDW5 = fnArrCount(vArrDay,"W",5)
	vDW0 = fnArrCount(vArrDay,"W",0)
	vDW1 = fnArrCount(vArrDay,"W",1)
	vDW2 = fnArrCount(vArrDay,"W",2)
	vDW3 = fnArrCount(vArrDay,"W",3)
	vDW4 = fnArrCount(vArrDay,"W",4)
	vDW6 = fnArrCount(vArrDay,"W",6)
	vDW7 = fnArrCount(vArrDay,"W",7)
	vDM5 = fnArrCount(vArrDay,"M",5)
	vDM0 = fnArrCount(vArrDay,"M",0)
	vDM1 = fnArrCount(vArrDay,"M",1)
	vDM2 = fnArrCount(vArrDay,"M",2)
	vDM3 = fnArrCount(vArrDay,"M",3)
	vDM4 = fnArrCount(vArrDay,"M",4)
	vDM6 = fnArrCount(vArrDay,"M",6)
	vDM7 = fnArrCount(vArrDay,"M",7)
	vSW5 = fnArrCount(vArrShop,"w",5)
	vSW0 = fnArrCount(vArrShop,"w",0)
	vSW1 = fnArrCount(vArrShop,"w",1)
	vSW2 = fnArrCount(vArrShop,"w",2)
	vSW3 = fnArrCount(vArrShop,"w",3)
	vSW4 = fnArrCount(vArrShop,"w",4)
	vSW6 = fnArrCount(vArrShop,"w",6)
	vSW7 = fnArrCount(vArrShop,"w",7)
	
	vTot5 = vTW5 + vTM5 + vDW5 + vDM5 + vSW5
	vTot0 = vTW0 + vTM0 + vDW0 + vDM0 + vSW0
	vTot1 = vTW1 + vTM1 + vDW1 + vDM1 + vSW1
	vTot2 = vTW2 + vTM2 + vDW2 + vDM2 + vSW2
	vTot3 = vTW3 + vTM3 + vDW3 + vDM3 + vSW3
	vTot4 = vTW4 + vTM4 + vDW4 + vDM4 + vSW4
	vTot6 = vTW6 + vTM6 + vDW6 + vDM6 + vSW6
	vTot7 = vTW7 + vTM7 + vDW7 + vDM7 + vSW7

	vTotTalkPC	= vTW5 + vTW0 + vTW1 + vTW2 + vTW3 + vTW4 + vTW6 + vTW7
	vTotTalkMob	= vTM5 + vTM0 + vTM1 + vTM2 + vTM3 + vTM4 + vTM6 + vTM7
	vTotDayPC	= vDW5 + vDW0 + vDW1 + vDW2 + vDW3 + vDW4 + vDW6 + vDW7
	vTotDayMob	= vDM5 + vDM0 + vDM1 + vDM2 + vDM3 + vDM4 + vDM6 + vDM7
	vTotShopPC	= vSW5 + vSW0 + vSW1 + vSW2 + vSW3 + vSW4 + vSW6 + vSW7
%>
<table border="1">
<tr align="center">
    <td bgcolor="<%= adminColor("tabletop") %>">기간</td>
    <td bgcolor="<%= adminColor("tabletop") %>">구분</td>
    <td bgcolor="<%= adminColor("tabletop") %>"><strong>오렌지</strong></td>
    <td bgcolor="<%= adminColor("tabletop") %>"><strong>옐로우</strong></td>
    <td bgcolor="<%= adminColor("tabletop") %>"><strong>그린</strong></td>
    <td bgcolor="<%= adminColor("tabletop") %>"><strong>블루</strong></td>
    <td bgcolor="<%= adminColor("tabletop") %>"><strong>VIP<br />실버</strong></td>
    <td bgcolor="<%= adminColor("tabletop") %>"><strong>VIP<br />골드</strong></td>
    <td bgcolor="<%= adminColor("tabletop") %>"><strong>VVIP</strong></td>
    <td bgcolor="<%= adminColor("tabletop") %>"><strong>스텝</strong></td>
    <td bgcolor="<%= adminColor("tabletop") %>">합계</td>
</tr>
<tr>
	<td align="center" rowspan="2">기프트톡</td>
	<td style="padding:5px;">PC</td>
	<td style="padding:5px;"><%= FormatNumber(vTW5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW6,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTW7,0) %></td>
	<td style="padding:5px;"><%= vTotTalkPC %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:5px;">모바일</td>
	<td style="padding:5px;"><%= FormatNumber(vTM5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM6,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vTM7,0) %></td>
	<td style="padding:5px;"><%= vTotTalkMob %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" rowspan="2">기프트데이</td>
	<td style="padding:5px;">PC</td>
	<td style="padding:5px;"><%= FormatNumber(vDW5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW6,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDW7,0) %></td>
	<td style="padding:5px;"><%= vTotDayPC %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:5px;">모바일</td>
	<td style="padding:5px;"><%= FormatNumber(vDM5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM6,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vDM7,0) %></td>
	<td style="padding:5px;"><%= vTotDayMob %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">기프트샵</td>
	<td style="padding:5px;">PC</td>
	<td style="padding:5px;"><%= FormatNumber(vSW5,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW0,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW1,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW2,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW3,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW4,0) %></td>
	<td style="padding:5px;"><%= FormatNumber(vSW6,0) %></td>
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
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTot6,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber(vTot7,0) %></strong></td>
	<td style="padding:5px;" bgcolor="<%= adminColor("tabletop") %>"><strong><%= FormatNumber((vTotTalkPC + vTotTalkMob + vTotDayPC + vTotDayMob + vTotShopPC),0) %></strong></td>
</tr>
</table>
<% End If %>
<% SET cStat = Nothing %>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->