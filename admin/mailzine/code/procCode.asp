<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
 dim iResult, sMode
 Dim iCodeValue, sCodeType, sCodeDesc, iCodeSort, blnUsing, sCodeDispYN
 Dim strSql, codevalue, sSortNo, i
 
 sMode		= requestCheckVar(Request("sM"),1)
 iCodeValue = requestCheckVar(Request("iCV"),10)
 sCodeType = requestCheckVar(Request("sCT"),20)
 sCodeDesc = requestCheckVar(Request("sCD"),20)
 iCodeSort = requestCheckVar(Request("iCS"),10)
 blnUsing = requestCheckVar(Request("rdoU"),1)
sCodeDispYN=requestCheckVar(Request("rdoD"),1)

if sCodeType="contentsKind" and Not(C_ADMIN_AUTH) then
	Call sbAlertMsg ("관리자 권한 이외에는 컨텐츠 종류 코드 생성이 불가능합니다.", "back", "") 
	dbget.close()	:	response.End
end if

IF sMode = "I" THEN
	strSql = "SELECT  code_value FROM [db_sitemaster].[dbo].[tbl_mailzine_code] Where code_type='"&sCodeType&"' and code_value="&iCodeValue	
	rsget.Open strSql,dbget
	IF not (rsget.eof or rsget.bof) then
		Call sbAlertMsg ("이미 존재하는 코드값입니다.다시 등록해주세요", "back", "") 
		dbget.close()	:	response.End
	end if
	rsget.close	
	
	strSql = " INSERT INTO [db_sitemaster].[dbo].[tbl_mailzine_code] (code_type,code_value, code_desc, code_sort, code_using,code_dispYN)"&_
			" Values('"&sCodeType&"',"&iCodeValue&",'"&sCodeDesc&"',"&iCodeSort&",'"&blnUsing&"','"&sCodeDispYN&"') "
	dbget.execute strSql			
ELSEIF sMode="U" THEN	
	strSql =" UPDATE [db_sitemaster].[dbo].[tbl_mailzine_code] Set code_desc = '"&sCodeDesc&"', code_sort = "&iCodeSort&" , code_using ='"&blnUsing&"', code_dispYN ='"&sCodeDispYN&"' "&_
			" WHERE code_type ='"&sCodeType&"' and code_value="&iCodeValue
		 
	dbget.execute strSql
ELSEIF sMode="D" THEN	
	strSql =" UPDATE [db_sitemaster].[dbo].[tbl_mailzine_code] Set code_using ='N', code_dispYN ='N' "&_
			" WHERE code_type ='"&sCodeType&"' and code_value="&iCodeValue
		 
	dbget.execute strSql
elseif sMode = "S" THEN
	'//리스트에서수정
	for i=1 to request.form("code_value").count
		codevalue = request.form("code_value")(i)
		sSortNo = request.form("viewidx")(i)
		if sSortNo="" then sSortNo="0"

		strSql = strSql & "Update [db_sitemaster].[dbo].[tbl_mailzine_code] Set " & vbCrLf
		strSql = strSql & " code_sort=" & sSortNo & "" & vbCrLf
		strSql = strSql & " Where code_value='" & codevalue & "';"
	Next
	If strSql <> "" then
		dbget.Execute strSql
	end if
END IF
	
	Call sbAlertMsg ("처리되었습니다.", "/admin/mailzine/code/popManageCode.asp?selCT="&sCodeType, "self") 
	
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->