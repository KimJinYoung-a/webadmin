<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� USB ����
' History : 2008.09.25 �ѿ�� ���� 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/admin_keyclass.asp" -->

<script language="javascript">

	function getsubmit(){

	if(!frm_edit.key_idx.value){
		alert("����kEY�� �Է����ּ���");
		frm_edit.key_idx.focus();
		return false;
	}

		frm_edit.mode.value = 'new';	
		frm_edit.submit();
	}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_edit" action="/admin/member/admin_keyprocess.asp" method="get">
	<input type="hidden" name="mode">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">		
		<td align="center">����KEY</td>
		<td align="center">Team</td>	
		<td align="center">�����</td>	
		<td align="center">�󼼻����</td>		
		<td align="center">��뿩��</td>
    </tr>
	<tr align="center" bgcolor="ffffff">		
		<td align="center"><input type="text" size=30 name="key_idx"></td>
		<td align="center">
			<select name="teamname">
				<option value="">����</option>
				<option value="CEO" >CEO</option>
				<option value="SYSTEM">SYSTEM</option>
				<option value="ONLINE">ONLINE</option>
				<option value="MARKETING">MARKETING</option>
				<option value="MD">MD</option>
				<option value="WD">WD</option>
				<option value="����">����</option>
				<option value="OFFLINE">OFFLINE</option>
				<option value="CS">CS</option>
				<option value="ITHINKSO">ITHINKSO</option>														
				<option value="�濵">�濵</option>
				<option value="FINGERS">FINGERS</option>
				<option value="�м�">�мǻ����</option>				
			</select>		
		</td>	
		<td align="center"><input type="text" name="username"></td>	
		<td align="center"><input type="text" name="username_detail"></td>		
		<td align="center">
			<select name="del_isusing">
				<option value="Y">���</option>
				<option value="N">����</option>
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
