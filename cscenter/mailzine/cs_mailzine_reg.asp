<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ������ ���
' History : 2007.12.20 �ѿ�� ����
'###########################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/mailzine/mailzinecls.asp"-->
<%
Dim omail,ix,idx
dim username,usermail,regdate,isusing
	idx = requestCheckVar(getNumeric(trim(request("idx"))),10)

set omail = new CMailzineList
	omail.frectidx = idx

	if idx <> "" then  	
		omail.Mailzine_oneitem
		
		username= ReplaceBracket(omail.FOneItem.fusername)
		usermail= ReplaceBracket(omail.FOneItem.fusermail)
		regdate= omail.FOneItem.fregdate	
		isusing= omail.FOneItem.fisusing	
	end if
%>
<script type='text/javascript'>

	function reg(idx){
		if (frm.username.value==''){
			alert('�̸��� �Է��ϼ���');
			frm.username.focus();
		}else if (frm.usermail.value==''){
			alert('�̸��� �ּҸ� �Է��ϼ���');
			frm.usermail.focus();
		}else if (frm.isusing.value==''){
			alert('��뿩�θ� �����ϼ���');		
		}else{			
			frm.action='/cscenter/mailzine/cs_mailzine_process.asp';
			frm.submit();
		}
	}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">			
	</td>
</tr>
</table>
<!-- �׼� �� -->

<form method="post" name="frm" style="margin:0px;">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
		<td  bgcolor="FFFFFF">
			<%= idx %><input type="hidden" name="idx" value="<%= idx %>">
		</td>	
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̸�</td>
		<td  bgcolor="FFFFFF">
			<input type="text" name="username" value="<%= username %>">
		</td>	
	</tr>	
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̸���</td>
		<td  bgcolor="FFFFFF">
			<input type="text" name="usermail" value="<%= usermail %>" size=40>
		</td>	
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td  bgcolor="FFFFFF">
			<%= regdate %>
		</td>
	</tr>	
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
		<td  bgcolor="FFFFFF">
			<select name="isusing">
			<option value="">����</option>
			<option value="Y" <% if isusing="Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing="N" then response.write " selected" %>>N</option>
			</select>
		</td>
	</tr>	
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="FFFFFF" colspan=2><input type="button" onclick="reg();" value="����" class="button"></td>
	</tr>		
</table>
</form>

<% set omail = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
