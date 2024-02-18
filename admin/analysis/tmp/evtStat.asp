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
    dim ret, evtcode, realtimeevt

    evtcode = request("evtcode")
    realtimeevt = request("realtimeevt")
    if evtcode = 111138 then
    strSql = " EXECUTE [db_datamart].[dbo].[usp_cm_evt_cmtstat_get2] '"& evtcode &"', '" & realtimeevt & "'"
    else
    strSql = " EXECUTE [db_datamart].[dbo].[usp_cm_evt_cmtstat_get] '"& evtcode &"', '" & realtimeevt & "'"
    end if
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
<% 
        for j=0 To UBound(ret,1) 
            if j = 0 then
%>
    <td><%=ret(j,i)%></td>
<%            
            else
%>
    <td><%=FormatNumber(ret(j,i), 0)%></td>
<%
            end if
%>
<% next %>	
</tr>
<%  next %>
</table>
<%
end if
%>
<style>
table th{height:36px; border:1px solid #72ac9c;}
table td{height:36px; border:1px solid #72ac9c;}
</style>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->