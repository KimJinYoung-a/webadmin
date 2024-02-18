<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim sData, rst, i
sData = "[{""idx"": 1,""title"": ""123"",""contents"": ""23232"",""regdate"": ""2018-10-20T17:20:11.89""},{""idx"": 3,""title"": ""string"",""contents"": ""string"",""regdate"": ""2018-10-20T19:30:25.193""}]"
'response.write sData

Set rst = JSON.parse(sData)
%>
<h1>test</h1>
<table>
<tr>
	<td>글번호</td>
	<td>제목</td>
	<td>내용</td>
</tr>
<% For i = 0 to rst.length - 1 %>
<tr>
	<td><%= rst.get(i).idx %></td>
	<td><%= rst.get(i).title %></td>
	<td><%= rst.get(i).contents %></td>
</tr>
<% Next %>
</table>
<% SET rst = nothing %>