<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/test/jtest/listCls.asp" -->
<%
Dim i, oList
SET olist = new cList
	oList.getList()
%>
<h1>test3</h1>
<table>
<tr>
	<td>�۹�ȣ</td>
	<td>����</td>
	<td>����</td>
</tr>
<% For i = 0 to olist.FResultCount - 1 %>
<tr>
	<td><%= oList.FItemList(i).FIdx %></td>
	<td><%= oList.FItemList(i).FTitle %></td>
	<td><%= oList.FItemList(i).FContents %></td>
</tr>
<% Next %>
</table>
<% SET olist = nothing %>