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
    dim ret, evtcode

    evtcode = request("evtcode")

    strSql = " EXECUTE [db_datamart].[dbo].[usp_cm_evt_commentstat_get] '"& evtcode &"'"
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
    <th>응모일자</th>
    <th>코멘트 수</th>
    <th>응모자 수</th>
    <th>신규가입자 수</th>
    <th>좋아요 수</th>
    <th>좋아요 참여자 수</th>
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

<%    
    strSql = " EXECUTE [db_datamart].[dbo].[usp_cm_evt_comment_totalstat_get] '"& evtcode &"'"
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
    <th>총 응모자</th>
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
<style>
table th{height:36px; border:1px solid #72ac9c;}
table td{height:36px; border:1px solid #72ac9c;}
</style>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->