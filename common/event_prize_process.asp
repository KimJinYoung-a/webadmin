<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� Event ��÷�ڵ�� ���� Y �� ��ȯ
' History : 2009.04.14 �ѿ�� ���� 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim eCode , egKindCode ,strSql
	eCode = requestCheckVar(request("eCode"),10)	'//�̺�Ʈ�ڵ�  ������ ��� ������ 4 �Դϴ�
	egKindCode = requestCheckVar(request("egKindCode"),10)	'//���Ľ����̼��� �̺�Ʈ�ڵ�

	'//�����̺�Ʈ ��÷�ڹ�ǥ �Ϸ� ó�� 
	if eCode = 4 then
		
		strSql = "update db_culture_station.dbo.tbl_culturestation_event set"+vbcrlf
		strSql = strSql & " prizeyn = 'Y'"+vbcrlf
		strSql = strSql & " where evt_code = "&egKindCode&""+vbcrlf
		
		'response.write strSql&"<br>"
		dbget.execute strSql		
		
	'//�Ϲ��̺�Ʈ ��÷�ڹ�ǥ �Ϸ� ó��
	else
		strSql = "update db_event.dbo.tbl_event set"+vbcrlf
		strSql = strSql & " prizeyn = 'Y'"+vbcrlf
		strSql = strSql & " where evt_code = "&eCode&""+vbcrlf
	
		'response.write strSql&"<br>"			
		dbget.execute strSql
	end if
%>
	<script type='text/javascript'>
		opener.location.reload();
		alert('ó���Ǿ����ϴ�');
		self.close();
	</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->

