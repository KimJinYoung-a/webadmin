<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �̺�Ʈ ������ ����Ʈ
' History : 2007.09.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventuserclass.asp"-->

<% 
dim seachbox,eventbox ,page , i
	seachbox = request("seachbox")
	eventbox = request("eventbox")
	page = Request("page")

If page="" Then page = 1

dim oeventuserlist
set oeventuserlist = new Ceventuserlist
	if eventbox = "3" then
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist3()
	end if 
	if eventbox = "5" then
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist5()
	end if 
	if eventbox = "7" then
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist7()
	end if 
	if eventbox = "9" then
		set oeventuserlist = new Ceventuserlist
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectinvaliduseryn="N"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist9()
	end if 

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=event_userlist" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body>
<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25" valign="top">
	<td>
		<font color="red"><strong><%= seachbox %> �̺�Ʈ ������ ����Ʈ</strong></font>
	</td>
</tr>
</table>
<!--ǥ ��峡-->
	
<table width="100%" border="0" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">
<tr bgcolor=#DDDDFF>
	<td align="center">id</td>
	<!--<td align="center">�ֹι�ȣ���ڸ�</td>-->
	<td align="center">����</td>
	<td align="center">�̸�</td>
	<td align="center">�̸���</td>
	<td align="center">��ȭ��ȣ</td>
	<td align="center">�ڵ�����ȣ</td>
	<td align="center">�ּ�</td>
	<td align="center">����</td>
	<td align="center">�ڸ�Ʈ</td>

	<% If eventbox = "9" Then %>
		<td align="center">��÷Ƚ��</td>
		<td align="center">�ֱٴ�÷��</td>
	<% End If %>
</tr>
<% if oeventuserlist.ftotalcount >0 then %>
<% for i = 0 to oeventuserlist.ftotalcount - 1 %>
<tr bgcolor=#FFFFFF>
	<td><%= oeventuserlist.flist(i).fuserid %></td>
	<!--<td><%'= left(oeventuserlist.flist(i).fjuminno,6) %>-->
	</td>
	<td>
		<%
		if mid(oeventuserlist.flist(i).fjuminno,8,1) = "1" then
		response.write "����"
		else
		response.write "����"
		end if
		%>
	</td>
	<td><%= oeventuserlist.flist(i).fusername %></td>
	<td><%= oeventuserlist.flist(i).fusermail %></td>
	<td><%= oeventuserlist.flist(i).fuserphone %></td>
	<td><%= oeventuserlist.flist(i).fusercell %></td>
	<td align="left">
		[<%= oeventuserlist.flist(i).fzipcode %>]
		&nbsp;
		<%= oeventuserlist.flist(i).faddress1 %>
		&nbsp;
		<%= oeventuserlist.flist(i).fuseraddr %>
	</td>
	<td><%= getUserLevelStr(oeventuserlist.flist(i).fLevel) %></td>
	<td><%= replace(oeventuserlist.flist(i).fevtcom_txt,"<","&lt;") %></td>

	<% If eventbox = "9" Then %>
		<td><%= oeventuserlist.flist(i).fWcnt %></td>
		<td><%= oeventuserlist.flist(i).fWdate %></td>
	<% End If %>
</tr>
<% next %>

<% else %>
<tr align="center" bgcolor="#DDDDFF">
	<td align=center bgcolor="#FFFFFF" colspan=15>�˻� ����� �����ϴ�.</td>
</tr>
<% end if %>	
</table>

</body>
</html>

<%
set oeventuserlist=nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

