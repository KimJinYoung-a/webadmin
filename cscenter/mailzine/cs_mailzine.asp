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
Dim omail,ix,page , isusing , username , usermail
	page = requestCheckVar(getNumeric(trim(request("page"))),10)
	isusing = requestCheckVar(trim(request("isusing")),1)
	username = requestCheckVar(trim(request("username")),32)
	usermail = requestCheckVar(trim(request("usermail")),128)

if page = "" then page = 1
if isusing = "" then isusing = "Y"
	
set omail = new CMailzineList
	omail.FPageSize = 50
	omail.FCurrPage = page
	omail.frectisusing = isusing
	omail.frectusername = username
	omail.frectusermail = usermail
	omail.MailzineList
%>
<script type='text/javascript'>

function addreg(idx){
	var addreg = window.open('/cscenter/mailzine/cs_mailzine_reg.asp?idx='+idx,'addreg','width=800,height=400,scrollbars=yes,resizable=yes');
	addreg.focus();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method=get action="" style="margin:0px;">
<input type="hidden" name="editor_no">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��뿩��: <select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>��뿩��</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			 			
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�̸�: <input type="text" name="username" value="<%=username%>">
		&nbsp;�̸���: <input type="text" name="usermail" value="<%=usermail%>">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* �̸��̳� �̸��� �ּҸ� ���� ���� �Է��ϼž� ������ ���ɴϴ�.
		<br>* ��뿩�� Y �ΰ�쿡�� ��ȸ�����Բ� �������� �߼� �˴ϴ�.
	</td>
	<td align="right">
		<input type="button" value="�űԵ��" onclick="addreg('','add');" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<form method=post name="monthly" style="margin:0px;">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		�˻���� : <b><%= omail.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= omail.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">idx</td>
	<td align="center">�̸�</td>
	<td align="center">�̸���</td>
	<td align="center">�����</td>		
	<td align="center">��뿩��</td>
	<td align="center">���</td>	
</tr>
<% if omail.FresultCount>0 then %>	
	<% for ix=0 to omail.FresultCount-1 %>
		<tr align="center" bgcolor="#FFFFFF">
			<td align="center"><% = omail.FItemList(ix).Fidx %></td>
			<td align="center"><% = ReplaceBracket(omail.FItemList(ix).fusername) %></td>
			<td align="center"><% = ReplaceBracket(omail.FItemList(ix).fusermail) %></td>		
			<td align="center"><% = FormatDate(omail.FItemList(ix).fregdate,"0000.00.00") %></td>
			<td align="center"><% = omail.FItemList(ix).fisusing %></td>
			<td align="center">
				<input type="button" value="����" class="button" onclick="addreg(<% = omail.FItemList(ix).Fidx %>);">
			</td>				
		</tr>
	<% next %>
	
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if omail.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= omail.StarScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for ix = 0 + omail.StarScrollPage to omail.StarScrollPage + omail.FScrollCount - 1 %>
				<% if (ix > omail.FTotalpage) then Exit for %>
				<% if CStr(ix) = CStr(omail.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= ix %></b></font></span>
				<% else %>
				<a href="?page=<%= ix %>" class="list_link"><font color="#000000"><%= ix %></font></a>
				<% end if %>
			<% next %>
			<% if omail.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= ix %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</form>

<% set omail = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
