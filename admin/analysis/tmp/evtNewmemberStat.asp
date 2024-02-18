<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트통계
' History : 2019.12.20 최종원 생성
'           2021.05.28 한용민 수정(오류수정)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%    
    dim strSql, i, j     
    dim ret, evtcode, realtimeevt, value

    evtcode = requestcheckvar(getNumeric(request("evtcode")),10)

    if evtcode="" or isnull(evtcode) then
        response.write "이벤트 코드가 없습니다."
        dbget.close() : db3_dbget.close() : response.end
    end if

    strSql = " EXECUTE [db_datamart].[dbo].[usp_cm_event_newmemberstat_get] '"& evtcode &"'"
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
        <th>이벤트코드</th>			
        <th>응모자수</th>			
        <th>신규 응모자 수</th>			
        <th>신규 구매자 수</th>			
        <th>신규 주문건 수</th>			
        <th>신규 매출</th>			
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
                If IsNumeric(ret(j,i)) Then
                    value = ret(j,i)
                Else
                    value = 0
                End If
%>
    <td><%=FormatNumber(value, 0)%></td>
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