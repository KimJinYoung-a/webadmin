<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbagirl_dbget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=dbagirl_rsget></OBJECT>
<%
dbagirl_dbget.Open Application("db_agirl")
%>
