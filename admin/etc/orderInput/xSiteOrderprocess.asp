<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �ֹ����� ����
' Hieditor : 2015.06.18 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->

<%
dim vmode, sql, i, voutmallorderseq, voutmallorderserial, vOrderName, vReceiveName
	voutmallorderseq = requestcheckvar(request("outmallorderseq"),10)
	voutmallorderserial = requestcheckvar(request("outmallorderserial"),32)
	vOrderName = requestcheckvar(request("OrderName"),32)
	vReceiveName = requestcheckvar(request("ReceiveName"),32)
	vmode = requestcheckvar(request("mode"),16)

'//�ű�����
if vmode = "orderedit" then
	if voutmallorderserial="" then
		response.write "<script type='text/javascript'>alert('�ش�Ǵ� �ֹ����� �����ϴ�.'); self.close();</script>"
		dbget.close()	:	response.end
	end if

	sql = "update db_temp.dbo.tbl_xSite_TMPOrder " + vbcrlf
	sql = sql & " set OrderName='"& html2db(vOrderName) &"'" + vbcrlf
	sql = sql & " ,ReceiveName='"& html2db(vReceiveName) &"' " + vbcrlf
	sql = sql & " where outmallorderserial = '"&voutmallorderserial&"'" + vbcrlf
	sql = sql & " and orderserial is NULL"+ vbcrlf
	
	'response.write sql
	dbget.execute sql

	response.write "<script type='text/javascript'>alert('ó�� �Ǿ����ϴ�.'); opener.location.reload(); self.close();</script>"
	dbget.close()	:	response.end
else
	response.write "<script type='text/javascript'>alert('�����ڰ� �����ϴ�.'); self.close();</script>"
	dbget.close()	:	response.end
end if	
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<script>
	opener.location.reload();
	self.close();
</script>
