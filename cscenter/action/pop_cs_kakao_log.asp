<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_smscertcls.asp" -->
<%
Dim i
Dim usercell
usercell = requestCheckVar(request("usercell"), 32)

Dim occssmscert
SET occssmscert = New CCSSMSCert
	occssmscert.FCurrPage = 1
	occssmscert.FPageSize = 100
	occssmscert.FRectUserCell = usercell
If (usercell <> "") Then
    occssmscert.GetCSKakaoLogList
End If
Dim occssmscert_cs
SET occssmscert_cs = New CCSSMSCert
	occssmscert_cs.FCurrPage = 1
	occssmscert_cs.FPageSize = 100
	occssmscert_cs.FRectUserCell = usercell
If (usercell <> "") Then
    occssmscert_cs.GetCSKakaoLogList_cs
End If
Dim occssmscert_mkt
SET occssmscert_mkt = New CCSSMSCert
	occssmscert_mkt.FCurrPage = 1
	occssmscert_mkt.FPageSize = 100
	occssmscert_mkt.FRectUserCell = usercell
If (usercell <> "") Then
    occssmscert_mkt.GetCSKakaoLogList_mkt
End If
%>
<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�޴��� : <input type="text" class="text" name="usercell" value="<%= usercell %>" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
	</td>

	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onclick="document.frm.submit()">
	</td>
</tr>
</table>
</form>

<br />

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* �˸��� ä�� : �ٹ�����. �ֱ� 100�Ǳ��� ���� �˴ϴ�.
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="50">IDX</td>
	<td width="100">�޴���</td>
	<td>�޼���</td>
	<td width="150">��û�Ͻ�</td>
	<td width="150">�����Ͻ�</td>
</tr>
<% if occssmscert.FResultCount > 0 then %>
<% for i = 0 to (occssmscert.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF" height="25">
        <td><%= occssmscert.FItemList(i).Fidx %></td>
		<td><%= occssmscert.FItemList(i).Fusercell %></td>
		<td><%= nl2br(occssmscert.FItemList(i).FMSG) %></td>
		<td><%= occssmscert.FItemList(i).FREQDATE %></td>
		<td><%= occssmscert.FItemList(i).FSENTDATE %></td>
    </tr>
<% next %>
<% else %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="30" colspan="5">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>

<br />

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* �˸��� ä�� : �ٹ����� ������. �ֱ� 100�Ǳ��� ���� �˴ϴ�.
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="50">IDX</td>
	<td width="100">�޴���</td>
	<td>�޼���</td>
	<td width="150">��û�Ͻ�</td>
	<td width="150">�����Ͻ�</td>
</tr>
<% if occssmscert_cs.FResultCount > 0 then %>
<% for i = 0 to (occssmscert_cs.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF" height="25">
        <td><%= occssmscert_cs.FItemList(i).Fidx %></td>
		<td><%= occssmscert_cs.FItemList(i).Fusercell %></td>
		<td><%= nl2br(occssmscert_cs.FItemList(i).FMSG) %></td>
		<td><%= occssmscert_cs.FItemList(i).FREQDATE %></td>
		<td><%= occssmscert_cs.FItemList(i).FSENTDATE %></td>
    </tr>
<% next %>
<% else %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="30" colspan="5">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>

<br />

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* �����ÿ� �� �뷮�߼�. �ֱ� 100�Ǳ��� ���� �˴ϴ�.
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="50">IDX</td>
	<td width="100">�޴���</td>
	<td>�޼���</td>
	<td width="150">��û�Ͻ�</td>
	<td width="150">�����Ͻ�</td>
</tr>
<% if occssmscert_mkt.FResultCount > 0 then %>
<% for i = 0 to (occssmscert_mkt.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF" height="25">
        <td><%= occssmscert_mkt.FItemList(i).Fidx %></td>
		<td><%= occssmscert_mkt.FItemList(i).Fusercell %></td>
		<td><%= nl2br(occssmscert_mkt.FItemList(i).FMSG) %></td>
		<td><%= occssmscert_mkt.FItemList(i).FREQDATE %></td>
		<td><%= occssmscert_mkt.FItemList(i).FSENTDATE %></td>
    </tr>
<% next %>
<% else %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="30" colspan="5">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>
<%
SET occssmscert = Nothing
SET occssmscert_cs = Nothing
SET occssmscert_mkt = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->