<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<%
dim oip, i
set oip = new organizerCls
	oip.FOrgStoryList

%>

<!-- 리스트 시작 -->
<table border="0" cellpadding="0" cellspacing="0">
<tr>
	<td><input type="button" onClick="location.href='storyview.asp';" value="새글쓰기"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<% IF oip.ftotalcount>0 Then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.ftotalcount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap> 번호</td>
		<td nowrap> 요약글 </td>
		<td nowrap> 작성일 </td>
		<td nowrap> 수정 </td>
	</tr>

	<% For i =0 To  oip.ftotalcount-1 %>
	<tr bgcolor="#FFFFFF">
		<td nowrap align="center"><%= oip.FItemList(i).FOW_IDX %> </td>
		<td nowrap><%= Left(oip.FItemList(i).FOW_TITLE,50) %>... </td>
		<td nowrap align="center"><%=oip.FItemList(i).FOW_REGDATE %></td>
		<td nowrap align="center"><input type="button" onClick="location.href='storyview.asp?idx=<%=oip.FItemList(i).FOW_IDX %>';" value="수정"></td>
	</tr>
	<% Next %>
<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
<% End IF %>

</table>
<!-- 리스트 끝 -->

<% set oip = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->