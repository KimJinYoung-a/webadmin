<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� �������� �ڸ�Ʈ ����Ʈ
' Hieditor : 2009.10.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<% 
dim photoid , page
	photoid = request("photoid")
	page = request("page")
	if page = "" then page = 1

dim ocomment , i
set ocomment = new cphoto_list
	ocomment.FPageSize = 20
	ocomment.FCurrPage = page
	ocomment.frectphotoid = photoid
	ocomment.fphotocomment_list()
%>

<script language="javascript">

	function comment_delete(idx){
		var comment_delete = window.open('/admin/momo/photo/photo_comment_process.asp?idx='+idx,'comment_delete','width=800,height=600,scrollbars=yes,resizable=yes');
		comment_delete.focus();
	}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ocomment.fresultcount > 0 then %>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">
			idx
		</td>
		<td align="center">
			�̹���
		</td>			
		<td align="center">
			photoID
		</td>	
		<td align="center">
			��ID
		</td>	
		<td align="center">
			�ڸ�Ʈ
		</td>	
		<td align="center">
			�����
		</td>	
		<td align="center">
			��뿩��
		</td>			
		<td align="center">
			���
		</td>	
	</tr>
	<% for i = 0 to ocomment.fresultcount -1 %>
	<tr bgcolor="FFFFFF">
		<td align="center">
			<%= ocomment.fitemlist(i).fidx %>
		</td>
		<td align="center">
			<img src="<%=webImgUrl%>/momo/photo/user/<%= ocomment.FItemList(i).fmainimage %>" width=50 height=50>
		</td>			
		<td align="center">
			<%= ocomment.fitemlist(i).fphotoid %>
		</td>	
		<td align="center">
			<%= ocomment.fitemlist(i).fuserid %>
		</td>	
		<td align="center">
			<%= nl2br(ocomment.fitemlist(i).fcomment) %>
		</td>	
		<td align="center">
			<%= left(ocomment.fitemlist(i).fregdate,10) %>
		</td>	
		<td align="center">
			<%= ocomment.fitemlist(i).fisusing %>
		</td>			
		<td align="center">
			<input type="button" class="button" value="�������" onclick="comment_delete(<%= ocomment.fitemlist(i).fidx %>);">
		</td>	
	</tr>	
	<% next %>
<% else %>
<tr bgcolor="FFFFFF">
	<td align="center">�˻� ����� �����ϴ�.
	</td>	
</tr>
<% end if %>

    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if ocomment.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ocomment.StartScrollPage-1 %>&photoid=<%=photoid%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocomment.StartScrollPage to ocomment.StartScrollPage + ocomment.FScrollCount - 1 %>
				<% if (i > ocomment.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocomment.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&photoid=<%=photoid%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocomment.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&photoid=<%=photoid%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
set	ocomment = nothing
%>