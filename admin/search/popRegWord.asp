<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �Ǹ� ��� ����
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

%>
<script language="javascript">
function jsSubmit() {
	var frm = document.frm;

	if (frm.aliasWord.value == '') {
		alert('���Ǿ �Է��ϼ���.');
		return;
	}

	if (frm.mainWord.value == '') {
		alert('����Ű���带 �Է��ϼ���.');
		return;
	}

	if (confirm('�����Ͻðڽ��ϱ�?') == true) {
		frm.mode.value = 'ins';
		frm.submit();
	}
}
</script>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>��ǰ�� Ű���� ���</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="popRegWord_process.asp">
<input type="hidden" name="mode" value="">
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">���Ǿ�</td>
	<td align="left">
		<input type="text" class="text" name="aliasWord" value="" size="20">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">����Ű����</td>
	<td align="left">
		<input type="text" class="text" name="mainWord" value="" size="20">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">
	    <input type="button" class="button" value="���" onClick="jsSubmit();">
	    <input type="button" class="button" value="���" onClick="self.close();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
