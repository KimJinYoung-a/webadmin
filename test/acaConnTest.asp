<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/function.asp"-->
<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbACADEMYget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsACADEMYget></OBJECT>

<%
dbACADEMYget.Open Application("db_academy")  '' "provider=sqloledb;Data Source=192.168.0.78;initial catalog=db_academy;user id=academyuser;password=wjddlswjddls" '' SQLNCLI
%>

<%

dbACADEMYget.Close

%>
