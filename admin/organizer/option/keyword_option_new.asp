<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.23 �ѿ�� ����
'	Description : ���ų�����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<script language="javascript">

	function getsubmit(){

	if(!frm_edit.option_value.value){
		alert("�ɼǸ��� �Է����ּ���");
		frm_edit.key_idx.focus();
		return false;
	}

	if(!frm_edit.type.value){
		alert("Ÿ���� �������ּ���");
		frm_edit.type.focus();
		return false;
	}
		
		frm_edit.mode.value = 'new';	
		frm_edit.mode_type.value = 'keyword';
		frm_edit.submit();
	}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_edit" action="/admin/organizer/option/option_reg.asp" method="get">
	<input type="hidden" name="mode">
	<input type="hidden" name="mode_type">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">		
		<td align="center">�ɼǸ�</td>
		<td align="center">���ļ���</td>
		<td align="center">Ÿ��</td>
		<td align="center">��뿩��</td>
    </tr>
	<tr align="center" bgcolor="ffffff">		
				<td align="center">
					<input type="text" size=30 name="option_value" >
				</td>	
				<td align="center"><input type="text" size=10 name="option_order" ></td>
				<td align="center">
					<select name="type" >
						<option value="" >����</option>
						<option value="style" >����</option>
						<option value="format" >����</option>						
						<option value="size" >���λ�����</option>													
							
					</select>
				</td>
				<td align="center">
					<select name="isusing" >
						<option value="" >����</option>
						<option value="Y" >Y</option>
						<option value="N" >N</option>
					</select>
				</td>
    </tr>  
</form>
	<tr align="center" bgcolor="ffffff">		
		<td align="left" colspan=5><input type="button" class="button" value="����" onclick="getsubmit();"></td>
    </tr>	      
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
