<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<% 
Dim objCmd, returnValue, sMode
Dim icatedepth,scatename,scatecode,ipcateidx,icategoryidx,blnUsing 
Dim menupos

sMode		= requestCheckvar(Request("hidM"),1)
icategoryidx= requestCheckvar(Request("icidx"),10)
icatedepth	= requestCheckvar(Request("icd"),10)
ipcateidx	= requestCheckvar(Request("selCL"),10)
scatename	= requestCheckvar(Request("scn"),64)
scatecode	= requestCheckvar(Request("scc"),5)
blnUsing	= requestCheckvar(Request("blnU"),1)
menupos		= requestCheckvar(Request("menupos"),10)

if (checkNotValidHTML(scatename) = true) Then
	response.write "<script>alert('ī�װ����� HTML�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	dbget.Close
	response.End
End If

SELECT CASE sMode
Case "I"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_edms_category_insert]( "&icatedepth&" ,'"&scatename&"', '"&scatecode&"' ,"&ipcateidx&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue = "1" THEN 
		call Alert_closenmove("��ϵǾ����ϴ�.","categorylist.asp?selCL="&ipcateidx&"&menupos="&menupos)
	ELSEIF 	returnValue = "2" THEN 
			Call Alert_move ("�Է��Ͻ� ī�װ��ڵ尪�� ������ ������Դϴ�.�ٽ� �Է����ּ���","popcategorydata.asp?icidx="&icategoryidx&"&menupos="&menupos)	
	ELSE	
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1") 
	END IF
	response.end 
Case "U"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_edms_category_update]("&icategoryidx&", "&icatedepth&" ,'"&scatename&"', '"&scatecode&"' ,"&ipcateidx&","&blnUsing&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue = "1" THEN
		Call Alert_closenmove ("�����Ǿ����ϴ�.","categorylist.asp?selCL="&ipcateidx&"&menupos="&menupos) 
	ELSEIF 	returnValue = "2" THEN 
		Call Alert_move ("�Է��Ͻ� ī�װ��ڵ尪�� ������ ������Դϴ�.�ٽ� �Է����ּ���","popcategorydata.asp?icidx="&icategoryidx&"&menupos="&menupos)		
	ELSE	
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1") 
	END IF
	response.end 
CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.0")
END SELECT
%>
