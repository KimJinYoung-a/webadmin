<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_mailzinecls.asp"-->

<%

Dim omail,ix,page

page = RequestCheckvar(request("page"),10)
if page = "" then page = 1

set omail = new CMailzineList
omail.FPageSize = 10
omail.FCurrPage = page
omail.MailzineList

%>
<script language='javascript'>
function input_mailzine(){
	location.href='/academy/mailzine/mailzine_input.asp?menupos=<%= menupos %>';
}
</script>

<form method=post name="monthly">
<table width="100%" cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black">
<tr>
	<td colspan="5" align="right"><input type="button" value="메일진등록" onclick="input_mailzine()"></td>
</tr>
<tr class="page_link">
	<td align="center" height="30" width="100">번호</td>
	<td align="center" height="30" width="100">메일진등록일</td>
	<td align="center" height="30" width="300">메일제목</td>
	<td align="center" height="30" width="100">보여주기여부</td>
	<td align="center" height="30" width="100">코드보기</td>
</tr>
<% if omail.FresultCount<1 then %>
<tr>
	<td colspan="5" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% else %>
<% for ix=0 to omail.FresultCount-1 %>
<tr class="page_link">
	<td align="center" height="30"><% = omail.FItemList(ix).Fidx %></td>
	<td align="center" height="30"><a href="/academy/mailzine/mailzine_detail.asp?idx=<% = omail.FItemList(ix).Fidx %>"><% = omail.FItemList(ix).Fregdate %></a></td>
	<td height="30">&nbsp;<a href="/academy/mailzine/mailzine_detail.asp?idx=<% = omail.FItemList(ix).Fidx %>"><% = omail.FItemList(ix).Ftitle %></a></td>
	<td align="center" height="30"><% = omail.FItemList(ix).Fisusing %></td>
	<td align="center" height="30"><a href="/academy/mailzine/mailzine_code_view.asp?idx=<% = omail.FItemList(ix).Fidx %>">보기</a></td>
</tr>
<% next %>
<% end if %>
<tr>
	<td colspan="13" height="30" align="center" class="page_link">
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
</table>
</form>
<% set omail = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->