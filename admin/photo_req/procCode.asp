<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
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
	strSql = "SELECT  code_value FROM [db_partner].[dbo].[tbl_photo_code] Where code_type='"&sCodeType&"' and code_value="&iCodeValue	
	rsget.Open strSql,dbget
	IF not (rsget.eof or rsget.bof) then
		Response.Write "<script>alert('이미 존재하는 코드값입니다.다시 등록해주세요');history.back();</script>"
		dbget.close()
		response.End
	end if
	rsget.close	
	
	strSql = " INSERT INTO [db_partner].[dbo].[tbl_photo_code] (code_type, code_value, code_name, code_useyn, code_sort)"&_
			" Values('"&sCodeType&"',"&iCodeValue&",'"&sCodeDesc&"','"&blnUsing&"','"&iCodeSort&"') "
	dbget.execute strSql
	
ELSEIF sMode="U" THEN	
	strSql =" UPDATE [db_partner].[dbo].[tbl_photo_code] Set code_name = '"&sCodeDesc&"', code_useyn ='"&blnUsing&"', code_sort = '"&iCodeSort&"'"&_
			" WHERE code_type ='"&sCodeType&"' and code_value="&iCodeValue
	dbget.execute strSql
	
END IF	
	
	Response.Write "<script>alert('처리되었습니다.');location.href='/admin/photo_req/popManageCode.asp?sCT="&sCodeType&"&selCT="&sCodeType&"';</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->