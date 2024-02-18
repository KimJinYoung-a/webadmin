<OBJECT RUNAT=server PROGID=ADODB.Connection id=db3_dbget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=db3_rsget></OBJECT>
<%
db3_dbget.Open Application("db_datamart") ''db_logics
%>
