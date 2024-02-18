<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Event  
' History : 2009.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<% 
dim evt_code, userid, page
	evt_code = request("evt_code")
	userid = request("userid")
	page = request("page")
	if page = "" then page = 1

dim oip , i
set oip = new cevent_list
	oip.FPageSize = 20
	oip.FCurrPage = page
	oip.frectevt_code = evt_code
	oip.frectUserId = userid
	oip.fevent_comment_list()
%>

<script language="javascript">

	function comment_delete(idx){
		var comment_delete = window.open('/admin/culturestation/event_comment_process.asp?idx='+idx,'comment_delete','width=800,height=600,scrollbars=yes,resizable=yes');
		comment_delete.focus();
	}

	function goPage(pg) {
		document.frm.page.value=pg;
		document.frm.submit();
	}
</script>

<form name="frm" method="GET">
<input type="hidden" name="page" value="" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>" align="center">검색조건</td>
	<td bgcolor="white">
		이벤트번호 : <input type="text" name="evt_code" value="<%=evt_code%>" size="6" /> /
		고객ID : <input type="text" name="userid" value="<%=userid%>" size="12" />
	</td>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>" align="center"><input type="submit" value="검색" class="button" /></td>
</tr>
</table>
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oip.fresultcount > 0 then %>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">
			번호
		</td>	
		<td align="center">
			EventCode
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
	<% for i = 0 to oip.fresultcount -1 %>
	<tr bgcolor="FFFFFF">
		<td align="center">
			<%= oip.fitemlist(i).fidx %>
		</td>	
		<td align="center">
			<%= oip.fitemlist(i).fevt_code %>
		</td>	
		<td align="center">
			<%= oip.fitemlist(i).fuserid %>
		</td>	
		<td align="center">
			<%= nl2br(oip.fitemlist(i).fcomment) %>
		</td>	
		<td align="center">
			<%= left(oip.fitemlist(i).fregdate,10) %>
		</td>	
		<td align="center">
			<%= oip.fitemlist(i).fisusing %>
		</td>			
		<td align="center">
			<input type="button" class="button" value="삭제" onclick="comment_delete(<%= oip.fitemlist(i).fidx %>);">
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
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="javascript:goPage(<%= oip.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:goPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="javascript:goPage(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
set	oip = nothing
%>