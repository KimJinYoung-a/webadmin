<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/test/tempdata/classes/eventcntcls.asp"-->

<%
dim oeventuserlist , i

	set oeventuserlist = new Ceventuserlist
	oeventuserlist.Feventuserlist3()
%>

<script language="javascript">
function excel()
{
var popup = window.open('/test/tempdata/40245_excel.asp','excel','width=1024,height=768,scrollbars=yes,resizable=yes');
popup.focus();
}
</script>

<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td colspan=13 align="center">īī����_���̾�Ʈ������(40245) �̺�Ʈ ����� �� ������</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">��¥</td>
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
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center">����Ǽ�</td>
		<% for i= 0 to oeventuserlist.FResultCount-1 %>
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
	<tr>
		<td>
			<input type="button" name="excelbox" value="�������Ϸ�����" class="button" onclick="excel();">
		</td>
	</tr>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->