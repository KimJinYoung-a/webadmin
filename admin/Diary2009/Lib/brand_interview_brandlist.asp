<% option Explicit %>
<% Response.CharSet = "euc-kr" %>
<%
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cahce"
Response.AddHeader "cache-Control", "no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->

<%
	Dim vQuery, i
	i = 0
	vQuery = "select i.makerid, u.socname, u.socname_kor from [db_diary2010].[dbo].[tbl_DiaryMaster] AS d "
	vQuery = vQuery & "inner join [db_item].[dbo].[tbl_item] AS i on d.itemid = i.itemid "
	vQuery = vQuery & "inner join [db_user].[dbo].[tbl_user_c] AS u on i.makerid = u.userid "
	vQuery = vQuery & "where d.isusing = 'Y' "
	vQuery = vQuery & "group by i.makerid, u.socname, u.socname_kor "
	vQuery = vQuery & "order by i.makerid ASC "
	rsget.Open vQuery,dbget,1
%>
<table class="a">
<tr>
	<td>Á¤·Ä : id ¼ø</td>
</tr>
<%	Do Until i = rsget.RecordCount-1 %>
<tr>
	<td style="cursor:pointer;" onClick="selectMakerid('<%=rsget("makerid")%>');">[<b><%=rsget("makerid")%></b>], <%=db2html(rsget("socname"))%>, <%=db2html(rsget("socname_kor"))%></td>
</tr>
<%
	i = i + 1
	rsget.MoveNext
	Loop
	
	rsget.close()
%>
<tr>
	<td align="right"><input type="button" class="button" value="´Ý±â" onClick="document.getElementById('branddiv').style.display='none';"></td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->