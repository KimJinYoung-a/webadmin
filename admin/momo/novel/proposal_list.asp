<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������
' Hieditor : 2009.11.19 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i,page , isusing
	menupos = request("menupos")
	page = request("page")
			
	if page = "" then page = 1

'// ����Ʈ
set ocontents = new cnovel_list
	ocontents.FPageSize = 5
	ocontents.FCurrPage = page	
	ocontents.fproposal_list()
%>

<script language="javascript">
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>		
	<td align="left">
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

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
<br>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ocontents.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= ocontents.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= ocontents.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   		
		<td align="center">��ȣ</td>
		<td align="center">���ID</td>
		<td align="center">�����</td>
		<td align="center">����</td>
		<td align="center">����</td>
    </tr>
	<% for i=0 to ocontents.FresultCount-1 %>
		
    <% if ocontents.FItemList(i).fisusing = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
    <% else %>    
    <tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<% end if %>
		<td align="center">
			<%= ocontents.FItemList(i).fidx %>
		</td>
		<td align="center">
			<%= ocontents.FItemList(i).fuserid %>
		</td>			
		<td align="center">
			<%= FormatDate(ocontents.FItemList(i).fregdate,"0000.00.00") %>
		</td>
		<td align="center">
			<%= ocontents.FItemList(i).ftitle %>
		</td>
		<td align="center">
			<%= nl2br(ocontents.FItemList(i).fcontents) %>
		</td>			
    </tr>   
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
	       	<% if ocontents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ocontents.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocontents.StartScrollPage to ocontents.StartScrollPage + ocontents.FScrollCount - 1 %>
				<% if (i > ocontents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocontents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocontents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set ocontents = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->