<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/breakdown/breakdownCls.asp"-->
<%
 dim iResult, sMode
 Dim iCodeValue, sCodeType, sCodeDesc, iCodeSort, blnUsing, sCodeComp, sCodeProd, sCodeTypeOld
 Dim strSql
 
 sMode			= requestCheckVar(Request("sM"),1)
 iCodeValue 	= requestCheckVar(Request("iCV"),10)
 sCodeType 		= requestCheckVar(Request("sCT"),20)
 sCodeTypeOld 	= requestCheckVar(Request("sCTO"),20)
 sCodeComp 		= requestCheckVar(Request("sCC"),150)
 sCodeProd 		= requestCheckVar(Request("sCP"),150)
 sCodeDesc 		= requestCheckVar(Request("sCD"),350)
 iCodeSort 		= requestCheckVar(Request("iCS"),10)
 blnUsing 		= requestCheckVar(Request("rdoU"),1)


IF sMode = "I" THEN
	strSql = "SELECT  code_value FROM [db_temp].[dbo].[tbl_breakdown_comCode] Where code_type='"&sCodeType&"' and code_value="&iCodeValue	
	rsget.Open strSql,dbget
	IF not (rsget.eof or rsget.bof) then
		Response.Write "<script>alert('이미 존재하는 코드값입니다.다시 등록해주세요');history.back();</script>"
		dbget.close()
		response.End
	end if
	rsget.close	
	
	strSql = " INSERT INTO [db_temp].[dbo].[tbl_breakdown_comCode] (code_type, code_value, code_comp, code_prod, code_desc, code_useyn, code_sort)"&_
			" Values('"&sCodeType&"',"&iCodeValue&",'"&sCodeComp&"','"&sCodeProd&"','"&sCodeDesc&"','"&blnUsing&"','"&iCodeSort&"') "
	dbget.execute strSql

ELSEIF sMode="U" THEN	
	strSql =" UPDATE [db_temp].[dbo].[tbl_breakdown_comCode] Set code_type ='"&sCodeType&"', code_desc = '"&sCodeDesc&"', code_useyn ='"&blnUsing&"', code_sort = '"&iCodeSort&"'"&_
			" , code_comp = '"&sCodeComp&"', code_prod = '"&sCodeProd&"'"&_
			" WHERE code_type ='"&sCodeTypeOld&"' and code_value="&iCodeValue

	dbget.execute strSql

END IF	
	
	Response.Write "<script>alert('처리되었습니다.');location.href='/admin/breakdown/popManageCode.asp?selCT="&sCodeType&"';</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->