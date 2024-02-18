<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/cjmall/cjmallitemcls.asp"-->
<%
Dim ocjmall, i
Set ocjmall = new CCjmall
	ocjmall.getcjmallMdConfirmList
%>
갯수 : <%= ocjmall.FResultCount %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">텐바이텐 상품번호</td>
	<td width="60">CJMall 상품번호</td>
</tr>
<% For i=0 to ocjmall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= ocjmall.FItemList(i).FItemId %></td>
	<td><%= ocjmall.FItemList(i).FCjmallPrdNo %></td>
</tr>
<% Next %>
</table>
<% Set ocjmall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
