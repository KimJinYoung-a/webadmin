<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim partner

response.write "partner:"&partner&"<br>"

partner = request("partner")
response.write "partner:"&partner&"<br>"

partner = request.QueryString("partner")
response.write "partner:"&partner&"<br>"

%>

<table width=800>
<tr>
	<td><a href="/admin/lib/popquickinfo.asp" target=_blank>판매/입출/재고/</a></td>
</tr>
<form name="frm" method="get">
<!-- input type="hidden" name="partner" value=""-->
<input type="button" value="Action" onclick="document.frm.submit();">
</form>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->