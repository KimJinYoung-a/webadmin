<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<%
 dim iResult, sMode
 Dim iCodeValue, sCodeType, sCodeDesc, iCodeSort, blnUsing
 Dim strSql
 
 sMode		= requestCheckVar(Request("sM"),1)
 iCodeValue = requestCheckVar(Request("iCV"),10)
 sCodeType = requestCheckVar(Request("sCT"),20)
 sCodeDesc = requestCheckVar(Request("sCD"),20)
 iCodeSort = requestCheckVar(Request("iCS"),10)
 blnUsing = requestCheckVar(Request("rdoU"),1)


IF sMode = "I" THEN
	strSql = "SELECT  code_value FROM [db_partner].[dbo].[tbl_cooperate_comCode] Where code_type='"&sCodeType&"' and code_value="&iCodeValue	
	rsget.Open strSql,dbget
	IF not (rsget.eof or rsget.bof) then
		Response.Write "<script>alert('�̹� �����ϴ� �ڵ尪�Դϴ�.�ٽ� ������ּ���');history.back();</script>"
		dbget.close()
		response.End
	end if
	rsget.close	
	
	strSql = " INSERT INTO [db_partner].[dbo].[tbl_cooperate_comCode] (code_type, code_value, code_desc, code_useyn, code_sort)"&_
			" Values('"&sCodeType&"',"&iCodeValue&",'"&sCodeDesc&"','"&blnUsing&"','"&iCodeSort&"') "
	dbget.execute strSql
	
	'####### �α� ���� (INSERT:1, �ڵ� ����:91) #######
	Call LogInsert("0","1","91")
	'####### �α� ���� #######
	
ELSEIF sMode="U" THEN	
	strSql =" UPDATE [db_partner].[dbo].[tbl_cooperate_comCode] Set code_desc = '"&sCodeDesc&"', code_useyn ='"&blnUsing&"', code_sort = '"&iCodeSort&"'"&_
			" WHERE code_type ='"&sCodeType&"' and code_value="&iCodeValue
	dbget.execute strSql
	
	If blnUsing = "N" Then
		'####### �α� ���� (DELETE:3, �ڵ� ����:93) #######
		Call LogInsert("0","3","92")
		'####### �α� ���� #######
	Else
		'####### �α� ���� (UPDATE:2, �ڵ� ����:92) #######
		Call LogInsert("0","2","92")
		'####### �α� ���� #######
	End If
	
END IF	
	
	Response.Write "<script>alert('ó���Ǿ����ϴ�.');location.href='popManageCode.asp?selCT="&sCodeType&"';</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->