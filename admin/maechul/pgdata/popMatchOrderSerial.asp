<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim idx
idx = request("idx")

%>

<script language="javascript">

function jsSubmitMatch(frm) {
	if (frm.OrderSerial.value == "") {
		alert("�ֹ���ȣ�� �Է��ϼ���.");
		return;
	}

	if (confirm("��Ī�Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsSubmitMatchForce(frm) {
	if (confirm("[������] ��Ī�Ͻðڽ��ϱ�?") == true) {
        frm.mode.value = 'forcematchorderserial';
		frm.submit();
	}
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="<%=stsAdmURL%>/admin/maechul/pgdata/pgdata_process.asp">
	<input type="hidden" name="mode" value="matchorderserial">
	<input type="hidden" name="logidx" value="<%= idx %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>�ֹ���ȣ</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td>
			<input type="text" class="text" name="OrderSerial" size="15">
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
    <input type="button" class="button" value="�Է��ϱ�" onClick="jsSubmitMatch(frm)">
    <% if C_ADMIN_AUTH then %>
    <input type="button" class="button" value="��������" onClick="jsSubmitMatchForce(frm)">
    <% end if %>
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
