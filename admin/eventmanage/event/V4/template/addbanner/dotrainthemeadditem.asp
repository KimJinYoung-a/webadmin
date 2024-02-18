<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/eventmanage/event_process.asp
' Description :  이벤트 개요 데이터처리 - 등록, 수정, 삭제
' History : 2007.02.12 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
'--------------------------------------------------------
' 변수선언 & 파라미터 값 받기
'--------------------------------------------------------
Dim k, sqlStr, i
Dim vCnt : vCnt = Request.Form("cksel").count
Dim eC : eC = requestCheckVar(Request.Form("eC"),9)
Dim mode : mode = requestCheckVar(Request.Form("mode"),3)
Dim stype : stype = requestCheckVar(Request.Form("stype"),1)
Dim upback : upback = requestCheckVar(Request.Form("upback"),1)
Dim reUrl : reUrl = Request.ServerVariables("HTTP_REFERER")
Dim GroupItemCheck : GroupItemCheck = requestCheckVar(Request.Form("GroupItemCheck"),1)

if Request.Form("cksel") <> "" then
	if checkNotValidHTML(Request.Form("cksel")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if Request.Form("sitemname") <> "" then
	if checkNotValidHTML(Request.Form("sitemname")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If


'배열로 처리
redim arritemcode(vCnt)
redim arritemname(vCnt)
for i=1 to vCnt
	arritemcode(i) = Request.Form("cksel")(i)
	arritemname(i) = Request.Form("sitemname")(i)
next
if mode="MR" then
	sqlStr = " Update [db_event].[dbo].[tbl_event_md_theme]"
	sqlStr = sqlStr & " Set GroupItemType='T'"
	sqlStr = sqlStr & " ,GroupItemCheck='" & GroupItemCheck & "'"
	sqlStr = sqlStr & " Where evt_code=" & eC
	dbget.Execute sqlStr
	Response.write "<script>parent.MainWindowReloadClose();</script>"
	response.End 
ElseIf mode="del" Then
	dbget.beginTrans
			sqlStr = " delete FROM [db_event].[dbo].[tbl_event_manual_group] WHERE evt_code=" & eC & " and itemid in (" & Request.Form("cksel") & ")"
			'Response.write sqlStr
			'Response.end
			dbget.execute sqlStr
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans 
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
			response.End 
		END IF
	dbget.CommitTrans
	sqlStr = " Update [db_event].[dbo].[tbl_event_md_theme]"
	sqlStr = sqlStr & " Set GroupItemType='T'"
	sqlStr = sqlStr & " ,GroupItemCheck='" & GroupItemCheck & "'"
	sqlStr = sqlStr & " Where evt_code=" & eC
	dbget.Execute sqlStr
	Response.write "<script>alert('삭제 되었습니다.');parent.TnDelThemeItemBanner();</script>"
	response.End 
Else
	If vCnt >= 1 Then
	dbget.beginTrans
			sqlStr = " delete FROM [db_event].[dbo].[tbl_event_manual_group] WHERE grouptype='T' and evt_code=" & eC
			dbget.execute sqlStr
		For k=1 To vCnt
			sqlStr = " IF Not Exists(SELECT IDX FROM [db_event].[dbo].[tbl_event_manual_group] WHERE grouptype='T' and itemid='" & arritemcode(k) & "' and evt_code="&eC& ")"			
			sqlStr = sqlStr + "	BEGIN "
			sqlStr = sqlStr+ " 			INSERT INTO [db_event].[dbo].[tbl_event_manual_group] (evt_code, itemid, itemname, viewidx, grouptype)"
			sqlStr = sqlStr + "     	VALUES (" & eC & ", " & arritemcode(k) &",'" & arritemname(k) & "'," & k & ", 'T')"
			sqlStr = sqlStr + " 	END "
			sqlStr = sqlStr + " ELSE "
			sqlStr = sqlStr + " 	BEGIN "			
			sqlStr = sqlStr + "			UPDATE [db_event].[dbo].[tbl_event_manual_group]"
			sqlStr = sqlStr + " 		SET viewidx ='" & k & "'"
			sqlStr = sqlStr + " 		WHERE grouptype='T' and evt_code = '" & eC & "' "
			sqlStr = sqlStr + " 		and itemid ="&arritemcode(k)&""
			sqlStr = sqlStr + " 	END "
			dbget.execute sqlStr
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans 
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
			response.End 
		END IF
		Next
		dbget.CommitTrans

		sqlStr = " Update [db_event].[dbo].[tbl_event_md_theme]"
		sqlStr = sqlStr & " Set GroupItemType='T'"
		sqlStr = sqlStr & " ,GroupItemCheck='" & GroupItemCheck & "'"
		sqlStr = sqlStr & " Where evt_code=" & eC
		dbget.Execute sqlStr
	End If
End If

If upback = "Y" Then
	Response.write "<script>alert('등록 되었습니다.');parent.TnSaveThemeItemBanner();</script>"
End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->