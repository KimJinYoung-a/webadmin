<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  이벤트 공통코드 DB 처리
' History : 2008.03.14 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
 dim iResult, sMode
 Dim iCodeValue, sCodeType, sCodeDesc, iCodeSort, blnUsing, sCodeDispYN
 Dim strSql
 
 sMode		= requestCheckVar(Request("sM"),1)
 iCodeValue = requestCheckVar(Request("iCV"),10)
 sCodeType = requestCheckVar(Request("sCT"),20)
 sCodeDesc = requestCheckVar(Request("sCD"),20)
 iCodeSort = requestCheckVar(Request("iCS"),10)
 blnUsing = requestCheckVar(Request("rdoU"),1)
sCodeDispYN=requestCheckVar(Request("rdoD"),1)

IF sMode = "I" THEN
	strSql = "SELECT  code_value FROM [db_event].[dbo].[tbl_event_commoncode] Where code_type='"&sCodeType&"' and code_value="&iCodeValue	
	rsget.Open strSql,dbget
	IF not (rsget.eof or rsget.bof) then
		Call sbAlertMsg ("이미 존재하는 코드값입니다.다시 등록해주세요", "back", "") 
		dbget.close()	:	response.End
	end if
	rsget.close	
	
	strSql = " INSERT INTO [db_event].[dbo].[tbl_event_commoncode] (code_type,code_value, code_desc, code_sort, code_using,code_dispYN)"&_
			" Values('"&sCodeType&"',"&iCodeValue&",'"&sCodeDesc&"',"&iCodeSort&",'"&blnUsing&"','"&sCodeDispYN&"') "
	dbget.execute strSql			
ELSEIF sMode="U" THEN	
	strSql =" UPDATE [db_event].[dbo].[tbl_event_commoncode] Set code_desc = '"&sCodeDesc&"', code_sort = "&iCodeSort&" , code_using ='"&blnUsing&"', code_dispYN ='"&sCodeDispYN&"' "&_
			" WHERE code_type ='"&sCodeType&"' and code_value="&iCodeValue
		 
	dbget.execute strSql				
END IF	
	
	Call sbAlertMsg ("처리되었습니다.", "/admin/member/tenbyten/popManageCode.asp?selCT="&sCodeType&"&sCT="&sCodeType, "self") 
	
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->