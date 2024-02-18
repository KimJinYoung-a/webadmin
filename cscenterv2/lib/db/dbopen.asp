<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsget></OBJECT>
<% dbget.Open Application(DATABASE_APPLICATION) %>

<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbget_CS></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsget_CS></OBJECT>
<% dbget_CS.Open Application(CS_DATABASE_APPLICATION) %>
