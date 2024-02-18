<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################################
' PageName : iteminfolink_process.asp
' Discription : I형(통합형) 이벤트 마케팅 상품 연동 등록
' History : 2022.06.16 정태훈
'###############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->

<%

dim eventCode, eventType, itemArray, existsCheck, sqlStr, eventFileURL

eventCode = requestCheckVar(Request.Form("eventCode"),10)
eventType = requestCheckVar(Request.Form("eventType"),10)
itemArray = requestCheckVar(Request.Form("itemArray"),256)

if itemArray <> "" then
	if checkNotValidHTML(itemArray) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if eventCode="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
	response.write "</script>"
	response.End
end if

	sqlStr = " SELECT top 1 eventCode" + vbcrlf
	sqlStr = sqlStr + " FROM [db_event].[dbo].[tbl_event_item_infoView]" + vbcrlf
	sqlStr = sqlStr + " WHERE eventCode=" & eventCode
	rsget.Open sqlStr, dbget
	IF not rsget.EOF THEN
		existsCheck = rsget("eventCode")
	END IF	
	rsget.close	

	if existsCheck > 0 then
		sqlStr = " Update [db_event].[dbo].[tbl_event_item_infoView]" & vbCrLf
		sqlStr = sqlStr & " Set eventType='" & eventType & "'" & vbCrLf
		sqlStr = sqlStr & " ,itemArray='" & itemArray & "'" & vbCrLf
		sqlStr = sqlStr & " ,updateUserID='" & session("ssBctId") & "'" & vbCrLf
		sqlStr = sqlStr & " ,lastUpdateDate=getdate()" & vbCrLf
		sqlStr = sqlStr & " Where eventCode='" & eventCode & "'"
		dbget.Execute sqlStr
	else
		sqlStr =" insert into [db_event].[dbo].[tbl_event_item_infoView]" & VbCrlf
		sqlStr = sqlStr & " (eventCode, eventType, itemArray, registeredUserID)" & VbCrlf
		sqlStr = sqlStr & " values(" & CStr(eventCode) & "," & eventType & ",'" & itemArray & "','" & session("ssBctId") & "')" & VbCrlf
		dbget.execute sqlStr
	end if

	if eventType="1" then
		eventFileURL = "/apps/appcom/wish/web2014/event/etc/secretshop/secret_shop.asp"
	end if
	if eventFileURL <> "" then
		sqlStr = " Update [db_event].[dbo].[tbl_event_display]" & vbCrLf
		sqlStr = sqlStr & " Set evt_execFile_mo='" & eventFileURL & "'" & vbCrlf
		sqlStr = sqlStr & " ,evt_isExec_mo=1" & vbCrlf
		sqlStr = sqlStr + " where evt_code=" & CStr(eventCode)
		dbget.Execute sqlStr
	end if

	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eventCode) + "&togglediv=1&viewset='+opener.document.frmEvt.viewset.value);"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->