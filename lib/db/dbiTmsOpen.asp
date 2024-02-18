<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbiTms_dbget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=dbiTms_rsget></OBJECT>
<%
dbiTms_dbget.Open Application("db_iTms")
%>
