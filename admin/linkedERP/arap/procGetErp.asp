<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����׸� ����Ʈ - ������
' History : 2011.11.15 ������  ����
'	jsSetARAP ��ũ��Ʈ �Լ� opener���� �����ؼ� ����ó��
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<%
Dim objCmd
		Set objCmd = Server.CreateObject("ADODB.COMMAND")  		
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdStoredProc	
			.CommandText = " db_partner.dbo.sp_TMS_get_SL_ACC_CD_sERP "			 
			.Execute, , adExecuteNoRecords
			End With	
						
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdStoredProc	
			.CommandText = " db_partner.[dbo].[sp_TMS_get_BA_ARAP_CD_sERP] "			 
			.Execute, , adExecuteNoRecords
			End With	 
		
	'	With objCmd
	'		.ActiveConnection = dbget
	'		.CommandType = adCmdStoredProc	
	'		.CommandText = " db_partner.dbo.sp_TMS_get_BA_PROD "			 
	'		.Execute, , adExecuteNoRecords
	'		End With		
			
	Set objCmd = nothing  
	
	response.redirect "/admin/approval/arap_edms/"
%>
 
<!-- #include virtual="/lib/db/dbClose.asp" -->