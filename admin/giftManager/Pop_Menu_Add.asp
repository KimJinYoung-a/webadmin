<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/giftManager/GiftManagerCls.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body>

<%

dim Depth,cdL,cdM,cdS
Depth = request("depth")
cdL= request("cdL")
cdM= request("cdM")
cdS= request("cdS")



dim objView

set objView = new giftManagerView
objView.getMenuView cdL,cdM,cdS

%>
<table width="400" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="UpdateFRM" action="Menu_Process.asp" target="">
	<input type="hidden" name="Depth" value="<%= Depth %>">
	<tr>
		<td width="130" bgcolor="#FFFFFF"></td>
		<td bgcolor="#FFFFFF"></td>
	</tr>
<% IF objView.LCode <>"" then %>
	<input type="text" name="LCode" size="4" value="<%= objView.LCode %>" />
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�� ī�װ�</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.LCode %></font>] <%= objView.LCodeNm %>
	</tr>
<% END IF %>

<% IF objView.MCode <>"" then %>
	<input type="text" name="MCode" size="4" value="<%= objView.MCode %>" />
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�� ī�װ�</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.MCode %></font>] <%= objView.MCodeNm %>
	</tr>
<% END IF %>

<% IF objView.SCode <>"" then %>
	<input type="text" name="SCode" size="4" value="<%= objView.SCode %>" />
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�� ī�װ�</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.SCode %></font>]<%= objView.SCodeNm %>
	</tr>
<% END IF %>


<% SELECT CASE DEPTH %>
	<% CASE "L" %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�� ī�װ�</td>
		<td bgcolor="#FFFFFF"><input type="text" size="4" name="LCode" value="">(1 ~ 99)</td>
	</tr>	
	<% CASE "M" %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�� ī�װ�</td>
		<td bgcolor="#FFFFFF"><input type="text" size="4" name="MCode" value="">(1 ~ 99)</td>
	</tr>		
	<% CASE "S" %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�� ī�װ�</td>
		<td bgcolor="#FFFFFF"><input type="text" size="4" name="SCode" value="">(1 ~ 99)</td>
	</tr>		
<% END SELECT %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
		<td bgcolor="#FFFFFF"><input type="text" size="4" name="OrderNo" value="0">(1 ~ 99)</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">ī�װ���</td>
		<td bgcolor="#FFFFFF"><input type="text" size="16" name="CodeNm" value=""></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">���ļ���</td>
		<td bgcolor="#FFFFFF">
			<select name="SortMethod">
				<option value="cashHigh">���ݼ�(������)</option>
				<option value="cashLow">���ݼ�(������)</option>
				<option value="itemidHigh">��ǰ��ȣ��(������)</option>
				<option value="itemidLow">��ǰ��ȣ��(������)</option>
				<option value="OrderNo">������ȣ��</option>
			</select>
		</td>
	</tr>
<% IF DEPTH = "L" THEN %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">ǥ������</td>
		<td bgcolor="#FFFFFF">
			<select name="ListType">
				<option value="list">��ǰ����Ʈ</option>
				<option value="wish">���ø���Ʈ</option>
				<option value="mania">�ŴϾ� ���̵�</option>
				<option value="event">�̺�Ʈ</option>
			</select>
		</td>
<% END IF %>
	<tr>
		<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="submit" class="button" value="����"></td>
	</tr>
	</form>
</table>
<% set objView = nothing %>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->