<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp" -->
<%
dim oldmisend, delaydate, notincludeupchecheck
delaydate = request("delaydate")
notincludeupchecheck = request("notincludeupchecheck")

if delaydate="" then delaydate=4

set oldmisend = New COldMiSend
oldmisend.FPageSize = 300
oldmisend.FRectDelayDate = delaydate
oldmisend.FRectNotInCludeUpcheCheck = notincludeupchecheck
'oldmisend.FRectNotIncludeItemList = "30633"
oldmisend.GetOldMisendListALL

dim i
%>
<table width="780"  class="a">
<form name="frm" method="get" action="">
<tr>
	<td><input type="checkbox" name="notincludeupchecheck" <% if notincludeupchecheck="on" then response.write "checked" %> >��üȮ������</td>
	<td align="center">******
	<select name=delaydate>
	<option value=2 <% if delaydate="2" then response.write "selected" %> >2
	<option value=3 <% if delaydate="3" then response.write "selected" %> >3
	<option value=4 <% if delaydate="4" then response.write "selected" %> >4
	<option value=5 <% if delaydate="5" then response.write "selected" %> >5
	<option value=6 <% if delaydate="6" then response.write "selected" %> >6
	<option value=7 <% if delaydate="7" then response.write "selected" %> >7
	</select>
	�� �̻� �̹�� ��� (�ִ� <%= oldmisend.FPageSize %>��) ******</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</form>
</table>
<table width="780" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0" class="a">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="70" align="center">�ֹ���ȣ</td>
    <td width="66" align="center">�Ա���</td>
    <td width="40" align="center">�ֹ���</td>
    <td width="80" align="center">��üID</td>
    <td width="50" align="center">��ǰ��ȣ</td>
    <td width="100" align="center">��ǰ��</td>
    <td width="60" align="center">�ɼ�</td>
    <td width="60" align="center">��۱���</td>
    <td width="60" align="center">����</td>
    <td width="100" align="center">���</td>
  </tr>
<% for i = 0 to oldmisend.FResultCount - 1 %>
  <tr height="20">
    <td><%= oldmisend.FItemList(i).ForderSerial %></td>
    <td><%= Left(oldmisend.FItemList(i).FIpkumDate,10) %></td>
    <td><%= oldmisend.FItemList(i).FBuyName %></td>
    <td><%= oldmisend.FItemList(i).FMakerID %></td>
    <td><%= oldmisend.FItemList(i).FItemID %></td>
    <td><%= oldmisend.FItemList(i).FItemName %></td>
    <td><%= oldmisend.FItemList(i).GetOptionName %></td>
    <td><font color="<%= oldmisend.FItemList(i).GetBeagonGubunColor %>"><%= oldmisend.FItemList(i).GetBeagonGubunName %></font></td>
    <td><font color="<%= oldmisend.FItemList(i).GetBeagonStateColor %>"><%= oldmisend.FItemList(i).GetBeagonStateName %></font></td>
    <td><%= oldmisend.FItemList(i).getMiSendCodeName %> <%= oldmisend.FItemList(i).getIpgoMayDay %></td>
  </tr>
<% next %>
</table>
<%
set oldmisend = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->