<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim idx, gubun
idx = request("idx")
gubun = request("gubun")

%>

<script language="javascript">

function jsSubmitReg(frm) {
	if (frm.reasonGubun.value == "") {
		alert("������ �����ϼ���");
		return;
	}

	if (confirm("��� �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="<%=stsAdmURL%>/admin/maechul/pgdata/pgdata_process.asp">
	<input type="hidden" name="mode" value="regReasonGubun<%= CHKIIF(gubun="off", "Off", "")%>">
	<input type="hidden" name="logidx" value="<%= idx %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>�󼼻���</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td>
			<select class="select" name="reasonGubun">
				<option></option>
				<option value="001">������(����)</option>
				<option value="002">������(���޻� ����)</option>
                <option value="003">������(�̴Ϸ�Ż)</option>
				<option value="020">������(��ġ��)</option>
				<option value="025">������(��ġ��ȯ��)</option>
				<option value="030">������(����Ʈ)</option>
				<option value="035">������(����Ʈȯ��)</option>
                <option value="004">������(B2B ����)</option>
				<option value="">---------------</option>
				<option value="040">CS����</option>
				<option value="">---------------</option>
				<option value="950">�������Ȯ��</option>
				<option value="999">��Ҹ�Ī</option>
				<option value="901">�ΰŽ����ݸ���</option>
				<option value="800">���ڼ���</option>
				<option value="900">��Ÿ</option>
				<option value="">---------------</option>
				<option value="XXX">�Է�����</option>
			</select>
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
<input type="button" class="button" value="�Է��ϱ�" onClick="jsSubmitReg(frm)">
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
