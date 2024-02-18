<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 수지항목 리스트 - 공통사용
' History : 2011.11.15 정윤정  생성
'	jsSetARAP 스크립트 함수 opener에서 생성해서 선택처리
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