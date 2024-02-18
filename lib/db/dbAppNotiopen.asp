<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbAppNotiget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsAppNotiget></OBJECT>

<%
dbAppNotiget.Open Application("db_appNoti") 
%>
