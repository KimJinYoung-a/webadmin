<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbCompanyGet></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsCompanyGet></OBJECT>
<%
dbCompanyGet.Open Application("db_logics")
%>