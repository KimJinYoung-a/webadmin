<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/board/offshop_noticecls.asp" -->

<%

dim idx
idx = request("idx")
dim nboard
set nboard = New CNoticeDetail
nboard.read(request("idx"))

%>
<script language="JavaScript">
<!--
function gotolist(){
location.href = "list.asp"
}
//-->
</script>
<input type="hidden" name="menupos" value="<%= menupos %>">


<table border="0" cellpadding="5" cellspacing="0" width="650" class="a">

<tr>
	<td bgcolor="white" style="padding:2" align="right" valign="bottom" colspan="2">
		<% = FormatDateTime(nboard.Fregdate,1) %>
	</td>
</tr>
<tr>
	<td bgcolor="#46699c" style="font-weight:bold;color:white">
		&nbsp; 제목 :  <%=nboard.Ftitle %> </td>
	<td bgcolor="#e0e0e0" align="right" width="150">
	 글쓴이 : <%=nboard.Fusername %>(<%=nboard.Fuserid %>)&nbsp;
	 </td>
</tr>
<tr>
	<td bgcolor="#EFEFEF" style="padding: 20 20 20 20;border-bottom:1 solid #99a9bc" colspan="2">
	<%= nl2br(nboard.FContents) %>
	<% if nboard.Ffile <> "" then %>
	<br><br>파일 링크 : <a href="<% = nboard.Ffilelink %>" target="_blank"><%=nboard.Ffile %></a>
	<% end if %>
	</td>
</tr>
</table>
 <table width="650" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="right"><input type="button" value="List" onclick="gotolist();"></td>
</tr>
</table>

<%
set nboard = Nothing
%>

<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->