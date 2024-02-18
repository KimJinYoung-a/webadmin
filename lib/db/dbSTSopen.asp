<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbSTSget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsSTSget></OBJECT>
<%
dbSTSget.Open Application("db_statistics") 
%>
