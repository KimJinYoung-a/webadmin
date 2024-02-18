<OBJECT RUNAT=server PROGID=ADODB.Connection id=dblogicsget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rslogicsget></OBJECT>
<%
dblogicsget.Open Application("db_logics")
%>