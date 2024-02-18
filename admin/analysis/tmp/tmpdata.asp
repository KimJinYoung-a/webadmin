<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%    
    dim strSql, i, j     
    dim ret, param1, param2, param3, param4

    param1 = request("param1")
    param2 = request("param2")
    param3 = request("param3")
    param4 = request("param4")

    strSql = " EXECUTE [db_datamart].[dbo].[usp_cm_tmp_data] '"& param1 &"','"& param2 &"','"& param3 &"','"& param4 &"'"
	db3_dbget.CursorLocation = adUseClient
	db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly,adLockReadOnly

    if  not db3_rsget.EOF  then
        ret = db3_rsget.getRows()
    end if
    db3_rsget.Close
if isArray(ret) then
%>
<table>
<tr>
    <% for i = 0 to db3_rsget.Fields.Count - 1 %>
        <th><%=db3_rsget.Fields(i).Name%></th>			
    <% next %>
</tr>
<%
	for i=0 To UBound(ret,2)
%>
<tr>
<% for j=0 To UBound(ret,1) %>
    <td><%=ret(j,i)%></td>
<% next %>	
</tr>
<%  next %>
</table>
<%
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->