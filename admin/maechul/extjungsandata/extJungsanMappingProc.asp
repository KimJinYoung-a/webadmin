<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim sqlStr, sellsite, yyyymm, ret, retval
dim mindate

sellsite = requestCheckVar(request("sellsite"), 32)
yyyymm = requestCheckVar(request("jsdate"), 10)

If sellsite = "" OR yyyymm = "" Then
	response.write	"<script language='javascript'>" &_
		"	alert('���޸� �Ǵ� ������� ���õ��� �ʾҽ��ϴ�. Ȯ�� �� ��õ� ���ּ���'); " &_
			"	location.href='about:blank;' " &_
		"</script>"
	response.end
Else
	rw "����!"
	response.flush
	response.clear
end If
dim objCmd, returnValue
Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping1]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 �߰�
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

if returnValue = 0 Then
	call Alert_return("��Ī1�ܰ� ����- �ٽ� �õ����ּ���")
	response.end
Else
	rw "1"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping1_D]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 �߰�
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

if returnValue = 0 Then
	call Alert_return("��Ī1-D�ܰ� ����- �ٽ� �õ����ּ���")
	response.end
Else
	rw "1-D"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping2]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 �߰�
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("��Ī2�ܰ� ����- �ٽ� �õ����ּ���")
 	response.end
Else
	rw "2"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping3]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 �߰�
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("��Ī3�ܰ� ����- �ٽ� �õ����ּ���")
 	response.end
Else
	rw "3"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping4]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 �߰�
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("��Ī4�ܰ� ����- �ٽ� �õ����ּ���")
 	response.end
Else
	rw "4"
	response.flush
	response.clear
end if


Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping5]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 �߰�
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("��Ī5�ܰ� ����- �ٽ� �õ����ּ���")
 	response.end
Else
	rw "5"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderMaster_mapping]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 �߰�
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("��ĪM�ܰ� ����- �ٽ� �õ����ּ���")
 	response.end
Else
	rw "M��..�Ϸ�!"
	response.flush
	response.clear
end if

%>
<script type="text/javascript">
alert("�ۼ��Ǿ����ϴ�.");
//location.href = "<%= manageUrl %>/common/popReloadOpener.asp";
opener.location.reload();
self.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->

