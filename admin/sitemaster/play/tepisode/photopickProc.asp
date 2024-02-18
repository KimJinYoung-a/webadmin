<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : T-Episode
' Hieditor : 이종화 생성
'			 2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim idx, viewtitle, subtitle, isusing, PPimg, style_html_m
Dim mode, sqlStr
	idx = requestCheckVar(getNumeric(request("idx")),10)
viewtitle		= request("viewtitle")
subtitle		= request("subtitle")
	isusing = requestCheckVar(request("isusing"),1)
PPimg			= request("photopickimg")
style_html_m = request("style_html_m")

If idx = "" then
	idx = 0
End If

If idx = 0 Then
	mode = "add"
Else
	mode = "edit"
End If

If (mode = "add") Then
	if viewtitle <> "" and not(isnull(viewtitle)) then
		viewtitle = ReplaceBracket(viewtitle)
	end If
	if subtitle <> "" and not(isnull(subtitle)) then
		subtitle = ReplaceBracket(subtitle)
	end If

	sqlStr = "INSERT into db_sitemaster.dbo.tbl_play_photo_pick "
	sqlStr = sqlStr & " (viewtitle, subtitle, isusing, PPimg)"
	sqlStr = sqlStr & " VALUES("
	sqlStr = sqlStr & " '" & html2db(viewtitle) & "'"
	sqlStr = sqlStr & " ,'" & html2db(subtitle) & "'"
	sqlStr = sqlStr & " ,'" & isusing & "'"
	sqlStr = sqlStr & " ,'" & html2db(PPimg) & "'"
	sqlStr = sqlStr & " )"
	dbget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_play_photo_pick') as idx"

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not Rsget.Eof then
		idx = rsget("idx")
	end if
	rsget.close
ElseIf mode = "edit" Then
	if viewtitle <> "" and not(isnull(viewtitle)) then
		viewtitle = ReplaceBracket(viewtitle)
	end If
	if subtitle <> "" and not(isnull(subtitle)) then
		subtitle = ReplaceBracket(subtitle)
	end If
	if style_html_m <> "" and not(isnull(style_html_m)) then
		style_html_m = ReplaceBracket(style_html_m)
	end If

	sqlStr = "UPDATE db_sitemaster.dbo.tbl_play_photo_pick SET "
	sqlStr = sqlStr & " viewtitle='" & html2db(viewtitle) & "'"
	sqlStr = sqlStr & " ,subtitle='" & html2db(subtitle) & "'"
	sqlStr = sqlStr & " ,isusing='" & isusing & "'"
	sqlStr = sqlStr & " ,PPimg='" & html2db(PPimg) & "'"
	sqlStr = sqlStr & " ,style_html_m='" & html2db(style_html_m) & "'"
	sqlStr = sqlStr & " WHERE idx=" & CStr(idx)

	'response.write sqlStr & "<Br>"
	'response.end
	dbget.Execute sqlStr
End If
response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
response.write "<script type='text/javascript'>location.replace('" & manageUrl & "/admin/sitemaster/play/tepisode/photopickEdit.asp?idx=" & Cstr(idx)& "&reload=on')</script>"
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
