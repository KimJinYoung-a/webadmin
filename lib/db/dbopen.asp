<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsget></OBJECT>

<%
'/���� �ֱ��� ������Ʈ ���� ������ ó�� '2011.11.11 �ѿ�� ����
'/������� ������ �ֽð� ������ ���� �ּ���
Call serverupdate_underconstruction()

dbget.Open Application("db_main")
%>
