<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
'###############################################
' PageName : procMatching.asp
' Discription : ī�װ� ��Ī ���, ����
'###############################################
dim cdl, cdm, cds, dispCate, userid, mode, chkdisp
dim objCmd,returnValue, i
cdl = requestCheckvar(request("cd1"),3)
cdm = requestCheckvar(request("cd2"),3)
cds = requestCheckvar(request("cd3"),3)
mode = requestCheckvar(request("mode"),3)
dispCate = requestCheckvar(request("disp"),16) 
chkdisp = html2db(request("chkdisp"))
userid = session("ssBctId")

Select Case mode
Case "map"
	'����ó��
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_item].[dbo].sp_Ten_CategoryMatching_SetData("&dispCate&",'"&cdl&"','"&cdm&"','"&cds&"','"&userid&"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = Nothing
	
	IF returnValue = 1 THEN
			
		%>
			<script type="text/javascript">
				alert("��Ī��� �Ǿ����ϴ�.");
				opener.location.reload();
				self.close();
			</script>
	<%		dbget.Close
			response.end
		END IF 
		
		dbget.Close
		Alert_return("������ó���� ������ �߻��߽��ϴ�.")      
Case "del"
	'���� ����
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_item].[dbo].sp_Ten_CategoryMatching_DelData('" & chkdisp & "')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = Nothing
	
	IF returnValue = 1 THEN
			
	%>
		<script type="text/javascript">
			alert("��Ī�� �����Ǿ����ϴ�.");
			parent.location.reload();
			//opener.location.reload();
			//self.close();
		</script>
<%		dbget.Close
		response.end
	END IF 

	dbget.Close
	Alert_return("������ó���� ������ �߻��߽��ϴ�.")     
End Select
%>