<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbACADEMYget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsACADEMYget></OBJECT>

<%
'/���� �ֱ��� ������Ʈ ���� ������ ó�� '2011.11.11 �ѿ�� ����
'/������� ������ �ֽð� ������ ���� �ּ���
Call serverupdate_underconstruction()

dbACADEMYget.Open Application("db_academy") 
%>
