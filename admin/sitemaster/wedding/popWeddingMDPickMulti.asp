<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : ���� ��ȹ�� ���������(�����)
' History : 2018.04.16 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<script language='javascript'>

	function SaveMainContents(frm){
	    if (frm.itemarr.value==""){
	        alert('������ ��ȣ�� �Է� �ϼ���.');
	        frm.itemarr.focus();
	        return;
	    }

	    if (confirm('���� �Ͻðڽ��ϱ�?')){
	        frm.submit();
	    }
	}

</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doWeddingMDPickUpdate.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="multi">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">�����۹�ȣ</td>
    <td>
		<textarea name="itemarr" rows="5" cols="50"></textarea>(�޸��� �������ּ���.)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
