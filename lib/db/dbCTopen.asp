<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbCTget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsCTget></OBJECT>

<%
dbCTget.Open Application("db_appWish") 
%>
