<%@ language=vbscript %>
<% option explicit %>
<%

%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/test/tempdata/classes/eventcntcls.asp"-->

<%
dim oeventuserlist , i

	set oeventuserlist = new Ceventuserlist
	oeventuserlist.Feventuserlist3()
%>

<%
Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"event_40245.xls"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="25" valign="top">
		<td align="center" colspan="13">
			<strong>카카오톡_다이어트대작전(40245) 이벤트 응모건수 데이터</strong>
		</td>
	</tr>
</table>
<br>
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">

			<tr align="center" bgcolor="grey">
				<td>날짜</td>
				<td align="center">2월13일</td>
				<td align="center">2월14일</td>
				<td align="center">2월15일</td>
				<td align="center">2월16일</td>
				<td align="center">2월17일</td>
				<td align="center">2월18일</td>
				<td align="center">2월19일</td>
				<td align="center">2월20일</td>
				<td align="center">2월21일</td>
				<td align="center">2월22일</td>
				<td align="center">2월23일</td>
				<td align="center">2월24일</td>
		    </tr>

			<tr bgcolor=#FFFFFF>
				<td align="center">응모건수</td>
				<% for i= 0 to 11 %>
				<td align="center"><%= oeventuserlist.flist(i) %></td>
				<% next %>
			</tr>
		</table>
<br>
<table>
	<tr>
		<td>
			※ 총 기간내 순수 참여자 수: <%= oeventuserlist.Ftotalcount %>
		</td>
	</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

