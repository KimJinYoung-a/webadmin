<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���ϸ��� ���� 
' History : 2007.10.23 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mileage/mileage_class.asp"-->

<% 
dim menupos
	menupos = request("menupos")
%>

<script language="javascript">
	function reg(){
	if(document.frm.jukyocd.value == ""){
	alert('���ϸ��� �ڵ��ȣ�� �Է��ϼ���');
	document.frm.jukyocd.focus();
	}
	else if(document.frm.jukyoname.value == ""){
	alert('�ڵ���� �Է��ϼ���');
	document.frm.jukyoname.focus();
	}
	else if(document.frm.isusing.value == ""){
	alert('���¸� �Է��ϼ���');
	document.frm.isusing.focus();
	}
	else
	{
	document.frm.action = "mileage_add_process.asp";
	document.frm.submit();
	}
	}
</script>
	


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>���ϸ������� ���</strong></font>
		</td>
	</tr>

	<form name="frm" action="" method="get">
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">
			���ϸ����ڵ��ȣ
		</td>
		<td bgcolor=#FFFFFF>
			<input type="text" class="text" name="jukyocd" size="20" maxlength="20" style='IME-MODE: disabled'> �ѱ��Է� X
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">
			�ڵ��
		</td>
		<td bgcolor=#FFFFFF>
			<input type="text" class="text" name="jukyoname" size="20" maxlength="20">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">
			����
		</td>
		<td bgcolor=#FFFFFF>
			<select class="select" name="isusing">
				<option>����</option>
				<option value="Y">���</option>
				<option value="N">������</option>
			</select>
		</td>
	</tr>
	</form>
	
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<input type="button" class="button" value="����" onclick="javascript:reg();">
			&nbsp;
			<input type="button" class="button" value="�ݱ�" onclick="javascript:window.close();">
		</td>
	</tr>		
</table>	

		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
