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
			<strong>īī����_���̾�Ʈ������(40245) �̺�Ʈ ����Ǽ� ������</strong>
		</td>
	</tr>
</table>
<br>
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">

			<tr align="center" bgcolor="grey">
				<td>��¥</td>
				<td align="center">2��13��</td>
				<td align="center">2��14��</td>
				<td align="center">2��15��</td>
				<td align="center">2��16��</td>
				<td align="center">2��17��</td>
				<td align="center">2��18��</td>
				<td align="center">2��19��</td>
				<td align="center">2��20��</td>
				<td align="center">2��21��</td>
				<td align="center">2��22��</td>
				<td align="center">2��23��</td>
				<td align="center">2��24��</td>
		    </tr>

			<tr bgcolor=#FFFFFF>
				<td align="center">����Ǽ�</td>
				<% for i= 0 to 11 %>
				<td align="center"><%= oeventuserlist.flist(i) %></td>
				<% next %>
			</tr>
		</table>
<br>
<table>
	<tr>
		<td>
			�� �� �Ⱓ�� ���� ������ ��: <%= oeventuserlist.Ftotalcount %>
		</td>
	</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

