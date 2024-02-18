<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbEVTget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsEVTget></OBJECT>

<%
dbEVTget.Open Application("db_EVT") 
%>
