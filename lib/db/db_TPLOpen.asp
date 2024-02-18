<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbget_TPL></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsget_TPL></OBJECT>
<%
dbget_TPL.Open Application("db_threepl")
%>