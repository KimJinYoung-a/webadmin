<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 회원등급별통계 엑셀다운로드
' Hieditor : 2019.03.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim yyyy, mm, i
	yyyy = requestcheckvar(Request("yyyy"),4)
	mm = requestcheckvar(Request("mm"),2)
dim tot_userlevelcount, tot_iOSexistscount, tot_ANDPushexistscount, tot_ANDALLY, tot_ANDALLN, tot_iOSALLY, tot_iOSALLN, tot_ANDPushY, tot_ANDPushN
dim tot_iOSPushY, tot_iOSPushN, tot_emailokY, tot_emailokN, tot_smsokY, tot_smsokN
	tot_userlevelcount=0
	tot_ANDPushexistscount=0
	tot_iOSexistscount=0
	tot_ANDALLY=0
	tot_ANDALLN=0
	tot_iOSALLY=0
	tot_iOSALLN=0
	tot_ANDPushY=0
	tot_ANDPushN=0
	tot_iOSPushY=0
	tot_iOSPushN=0
	tot_emailokY=0
	tot_emailokN=0
	tot_smsokY=0
	tot_smsokN=0

if yyyy="" then yyyy=year(date)
if mm="" then mm=Format00(2,month(date))

dim oreport
set oreport = new CUserLevelMonth
	oreport.FRectyyyymm = yyyy & "-" & mm
	oreport.GetLevelList

dim oagreeY
set oagreeY = new CUserLevelMonth
	oagreeY.FRectyyyymm = yyyy & "-" & mm
	oagreeY.GetLevelagreeList

dim oHOLD
set oHOLD = new CUserLevelMonth
	oHOLD.GetUserHOLD_count

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_회원등급별통계.xls"
Response.CacheControl = "public"
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		검색결과 : <b><%= oagreeY.FresultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>날짜<br>(개편이후)</td>
	<td>회원등급</td>
	<td>고객수</td>
	<td>안드보유수</td>
	<td>ios보유수</td>
	<td>안드ALL수신Y</td>
	<td>안드ALL수신N</td>
	<td>iosALL수신Y</td>
	<td>iosALL수신N</td>
	<td>안드푸시수신Y</td>
	<td>안드푸시수신N</td>
	<td>ios푸시수신Y</td>
	<td>ios푸시수신N</td>
	<td>이메일수신Y</td>
	<td>이메일수신N</td>
	<td>문자수신Y</td>
	<td>문자수신N</td>
</tr>

<% if oagreeY.FresultCount>0 then %>
	<% for i=0 to oagreeY.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td class='txt'><%= oagreeY.FItemList(i).fyyyymm %></td>
		<td><%= oagreeY.FItemList(i).fuserlevelname %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fuserlevelcount,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fANDPushexistscount,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSexistscount,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fANDALLY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fANDALLN,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSALLY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSALLN,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fANDPushY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fANDPushN,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSPushY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSPushN,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).femailokY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).femailokN,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fsmsokY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fsmsokN,0) %></td>
	</tr>
	<%
	tot_userlevelcount = tot_userlevelcount + oagreeY.FItemList(i).fuserlevelcount
	tot_ANDPushexistscount = tot_ANDPushexistscount + oagreeY.FItemList(i).fANDPushexistscount
	tot_iOSexistscount = tot_iOSexistscount + oagreeY.FItemList(i).fiOSexistscount
	tot_ANDALLY = tot_ANDALLY + oagreeY.FItemList(i).fANDALLY
	tot_ANDALLN = tot_ANDALLN + oagreeY.FItemList(i).fANDALLN
	tot_iOSALLY = tot_iOSALLY + oagreeY.FItemList(i).fiOSALLY
	tot_iOSALLN = tot_iOSALLN + oagreeY.FItemList(i).fiOSALLN
	tot_ANDPushY = tot_ANDPushY + oagreeY.FItemList(i).fANDPushY
	tot_ANDPushN = tot_ANDPushN + oagreeY.FItemList(i).fANDPushN
	tot_iOSPushY = tot_iOSPushY + oagreeY.FItemList(i).fiOSPushY
	tot_iOSPushN = tot_iOSPushN + oagreeY.FItemList(i).fiOSPushN
	tot_emailokY = tot_emailokY + oagreeY.FItemList(i).femailokY
	tot_emailokN = tot_emailokN + oagreeY.FItemList(i).femailokN
	tot_smsokY = tot_smsokY + oagreeY.FItemList(i).fsmsokY
	tot_smsokN = tot_smsokN + oagreeY.FItemList(i).fsmsokN
	next
	%>

	<tr align="center" bgcolor="#FFFFFF">
		<td colspan=2>합계</td>
		<td><%= FormatNumber(tot_userlevelcount,0) %></td>
		<td><%= FormatNumber(tot_ANDPushexistscount,0) %></td>
		<td><%= FormatNumber(tot_iOSexistscount,0) %></td>
		<td><%= FormatNumber(tot_ANDALLY,0) %></td>
		<td><%= FormatNumber(tot_ANDALLN,0) %></td>
		<td><%= FormatNumber(tot_iOSALLY,0) %></td>
		<td><%= FormatNumber(tot_iOSALLN,0) %></td>
		<td><%= FormatNumber(tot_ANDPushY,0) %></td>
		<td><%= FormatNumber(tot_ANDPushN,0) %></td>
		<td><%= FormatNumber(tot_iOSPushY,0) %></td>
		<td><%= FormatNumber(tot_iOSPushN,0) %></td>
		<td><%= FormatNumber(tot_emailokY,0) %></td>
		<td><%= FormatNumber(tot_emailokN,0) %></td>
		<td><%= FormatNumber(tot_smsokY,0) %></td>
		<td><%= FormatNumber(tot_smsokN,0) %></td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=2>휴면계정 합계</td>
		<td>
			<% if oHOLD.ftotalcount>0 then %>
				<%= FormatNumber(oHOLD.FOneItem.fUserHOLD_count,0) %>
			<% else %>
				0
			<% end if %>
		</td>
		<td colspan=14></td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="17" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<% 'if false then %>
<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="12">
		검색결과 : <b><%= oreport.FresultCount %></b>
		&nbsp;&nbsp;&nbsp;&nbsp; ※ <%= year(dateadd("m",-1,dateserial(yyyy,mm,"01"))) %>년 <%= month(dateadd("m",-1,dateserial(yyyy,mm,"01"))) %>월말 기준
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">년/월<br />(개편이전)</td>
	<td width="90"><br />(Orange)</td>
	<td width="90">White<br />(Yellow)</td>
	<td width="90">Red<br />(Green)</td>
	<td width="90">Blue<br />(VIP)</td>
	<td width="90">VIP Gold<br />(VIP Silver)</td>
	<td width="90">VVIP<br />(VIP Gold)</td>
	<td width="90"><br />(VVIP)</td>
	<td width="90">Staff</td>
	<td width="90">FAMILY</td>
	<td width="90">BIZ</td>
	<td width="93" bgcolor="#E0E0E0">소계</td>
</tr>

<% if oreport.FresultCount>0 then %>
	<% for i=0 to oreport.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td class='txt'><%=oreport.FItemList(i).FAxisDate%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FOrange,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FYellow,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FGreen,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FBlue,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FSilver,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FGold,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FVVIP,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FStaff,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).fFAMILY,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).fBIZ,0)%></td>
		<td bgcolor="#FAFAFA"><%=FormatNumber(oreport.FItemList(i).FTotal,0)%></td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<% 'end if %>
</body>
</html>

<%
set oreport = nothing
set oagreeY = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->