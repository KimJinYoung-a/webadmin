<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 오프라인 메일진
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_mailzinecls.asp"-->

<%
Dim omail,ix,page
page = requestCheckVar(request("page"),10)

if page = "" then page = 1

set omail = new CMailzineList
omail.FPageSize = 10
omail.FCurrPage = page
omail.MailzineList

%>
<script language="JavaScript">
<!--

function TnPreViewMailzine(idx){
	window.open('/admin/offshop/preview_mailzine.asp?idx=' + idx,'preview','scrollbars=1,width=620, height=540');
}
function TnSendMail(idx){
		document.mailform.idx.value=idx;
		document.mailform.submit();
}
//-->
</script>
<form method="post" action="/admin/offshop/lib/dooffshopmailzine.asp" name="mailform">
<input type="hidden" name="idx">
</form>
<br>
<a href="/admin/offshop/offshop_mailzine_register.asp"><font color="red">******** New ********</font></a>
<form method=post name="monthly">
<table cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black">
<tr class="page_link">
	<td align="center" height="30" width="100">번호</td>
	<td align="center" height="30" width="100">메일진등록일</td>
	<td align="center" height="30" width="300">메일제목</td>
	<td align="center" height="30" width="100">메일발송여부</td>
	<td align="center" height="30" width="100">미리보기</td>
	<td align="center" height="30" width="100">메일발송</td>
</tr>
<% if omail.FresultCount<1 then %>
<tr>
	<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% else %>
<% for ix=0 to omail.FresultCount-1 %>
<tr class="page_link">
	<td align="center" height="30"><% = omail.FItemList(ix).Fidx %></td>
	<td align="center" height="30"><% = omail.FItemList(ix).Fregdate %></td>
	<td height="30">&nbsp;<a href="/admin/offshop/offshop_mailzine_edit.asp?idx=<% = omail.FItemList(ix).Fidx %>"><% = omail.FItemList(ix).Ftitle %></a></td>
	<td align="center" height="30"><% if omail.FItemList(ix).Fmailyn = "Y" then %><font color="red"><% else %><font color="#000000"><% end if %></font><% = omail.FItemList(ix).Fmailyn %></font></td>
	<td align="center" height="30"><a href="javascript:TnPreViewMailzine(<% = omail.FItemList(ix).Fidx %>);">보기</a></td>
   <td align="center" height="30"><a href="javascript:TnSendMail(<% = omail.FItemList(ix).Fidx %>);">발송</a></td>
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