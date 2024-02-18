<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ȸ����
' Hieditor : 2009.12.28 �ѿ�� ����
'			 2022.07.06 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/specialshop/specialshop_cls.asp"-->

<%
dim id,openDate,status,regdate , i , statusstr , itemcount , isusing, title, endDate
dim mode
	id = requestCheckVar(getNumeric(request("id")),10)
	mode = requestCheckVar(request("mode"),32)

dim ospecialshop
set ospecialshop = new cspecialshop_list
	ospecialshop.frectid = id
	
	'//������� �ϰ�츸 ����
	if id <> "" then
	ospecialshop.fspecialshop_oneitem()

		if ospecialshop.ftotalcount > 0 then
			statusstr = ospecialshop.FOneItem.fstatusstr
			openDate = formatdate(ospecialshop.FOneItem.fopenDate,"0000-00-00")
			status = ospecialshop.FOneItem.fstatus
			regdate = ospecialshop.FOneItem.fregdate
			itemcount = ospecialshop.FOneItem.fitemcount
			isusing = ospecialshop.FOneItem.fisusing
			title = ReplaceBracket(ospecialshop.FOneItem.ftitle)
			endDate = ospecialshop.FOneItem.FendDate
		end if
	end if
%>

<script type='text/javascript'>

// ���&����
function reg(){
	if (frm.title.value==''){
		alert('�׸��� ����ϼ���');
		frm.title.focus();
		return;
	}
	
	if (frm.openDate.value==''){
		alert('�������� ����ϼ���');
		frm.openDate.focus();
		return;
	}
	
	<% If status = "" OR status = "1" Then %>
	if (frm.endDate.value==''){
		alert('�������� ����ϼ���');
		frm.endDate.focus();
		return;
	}
	<% End If %>

	if (frm.status.value==''){
		alert('���¸� ���� �ϼ���');
		frm.status.focus();
		return;
	}
	
	if (frm.isusing.value==''){
		alert('��뿩�θ� ���� �ϼ���');
		frm.isusing.focus();
		return;
	}	
	
	frm.mode.value='reg';	
	frm.action='/admin/shopmaster/specialshop/specialshop_process.asp';
	frm.submit();
}

</script>

<form name="frm" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25" width="70">ID</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
		<%= id %><input type="hidden" name="id" value="<%= id %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">�׸�</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
		<input type="text" name="title" size="70" value="<%= title %>">			
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">������</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
	<% If status = "" OR status = "0" OR status = "1" Then %>
		<input type="text" name="openDate" size=10 value="<%= openDate %>">	
		<a href="javascript:calendarOpen3(frm.openDate,'������',frm.openDate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
	<% Else %>
		<%= openDate %><input type="hidden" name="openDate" value="<%= openDate %>">
	<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">������</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
	<% If status = "" OR status = "0" OR status = "1" Then %>
		<input type="text" name="endDate" size=10 value="<%= endDate %>">			
		<a href="javascript:calendarOpen3(frm.endDate,'������',frm.endDate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
	<% Else %>
		<%= endDate %><input type="hidden" name="endDate" value="<%= endDate %>">
	<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="75">����</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
	<% If status = "" then status = "0" end if %>
		<% drawstatus "status" , status, id %>
		<br><br>&nbsp;* ���� �����ϸ� ������ 00�ÿ� ��뿩�ΰ� Y�ΰ͵��� �ڵ� ���� �˴ϴ�.
		<br>&nbsp;* �������� ������ �ڵ� ���� �˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">��뿩��</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>����</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2 height="50"><input type="button" onclick="reg();" value=" �� �� " class="button" style="width:80px;height:40px;"></td>
</tr>
</table>
</form>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->