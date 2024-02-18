<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 한줄소설
' Hieditor : 2009.11.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<% 
dim novelid , page
	novelid = request("novelid")
	page = request("page")
	if page = "" then page = 1

dim ocomment , i
set ocomment = new cnovel_list
	ocomment.FPageSize = 20
	ocomment.FCurrPage = page
	ocomment.frectnovelid = novelid
	ocomment.fnovelcomment_list()
%>

<script language="javascript">

	function comment_delete(idx){
		var comment_delete = window.open('/admin/momo/novel/novel_comment_process.asp?idx='+idx,'comment_delete','width=800,height=600,scrollbars=yes,resizable=yes');
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
			한줄소설id
		</td>	
		<td align="center">
			고객ID
		</td>	
		<td align="center">
			코맨트
		</td>	
		<td align="center">
			등록일
		</td>	
		<td align="center">
			사용여부
		</td>			
		<td align="center">
			비고
		</td>	
	</tr>
	<% for i = 0 to ocomment.fresultcount -1 %>
	<tr bgcolor="FFFFFF">
		<td align="center">
			<%= ocomment.fitemlist(i).fidx %>
		</td>	
		<td align="center">
			<%= ocomment.fitemlist(i).fnovelid %>
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
			<input type="button" class="button" value="노출안함" onclick="comment_delete(<%= ocomment.fitemlist(i).fidx %>);">
		</td>	
	</tr>	
	<% next %>
<% else %>
<tr bgcolor="FFFFFF">
	<td align="center">검색 결과가 없습니다.
	</td>	
</tr>
<% end if %>

    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if ocomment.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ocomment.StartScrollPage-1 %>&novelid=<%=novelid%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocomment.StartScrollPage to ocomment.StartScrollPage + ocomment.FScrollCount - 1 %>
				<% if (i > ocomment.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocomment.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&novelid=<%=novelid%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocomment.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&novelid=<%=novelid%>">[next]</a></span>
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