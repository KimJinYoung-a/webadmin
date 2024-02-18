<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : goodnote_process.asp
' Discription : 굿노트 스티커 등록 프로세스
' History : 2023.04.03 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
dim idx, title, font_color, title_image, bg_color, contents_image, contents_title, eFolder
dim contents, brand_button_color, brand_url, file_url, start_date, end_date, isusing, mode
dim sqlStr

mode = requestCheckVar(Request.Form("mode"),10)
idx = requestCheckVar(Request.Form("idx"),10)
title = requestCheckVar(Request.Form("title"),128)
font_color = requestCheckVar(Request.Form("font_color"),8)
title_image = requestCheckVar(Request.Form("title_image"),128)
bg_color = requestCheckVar(Request.Form("bg_color"),8)
contents_image = requestCheckVar(Request.Form("contents_image"),128)
contents_title = requestCheckVar(Request.Form("contents_title"),128)
contents = html2db(Request.Form("contents"))
brand_button_color	= requestCheckVar(Request.form("brand_button_color"),8)
brand_url	= requestCheckVar(Request.form("brand_url"),128)
file_url	= requestCheckVar(Request.form("file_url"),128)
start_date	= requestCheckVar(Request.form("start_date"),10)
end_date	= requestCheckVar(Request.form("end_date"),10)
isusing	= requestCheckVar(Request.form("isusing"),1)

end_date = end_date & " 23:59:59"

if contents <> "" then
	if checkNotValidHTML(contents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if title <> "" then
	if checkNotValidHTML(title) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if title_image <> "" then
	if checkNotValidHTML(title_image) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if contents_image <> "" then
	if checkNotValidHTML(contents_image) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if brand_url <> "" then
	if checkNotValidHTML(brand_url) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if file_url <> "" then
	if checkNotValidHTML(file_url) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

select case mode
case "add"
	dbget.beginTrans
    '===========================================================
    '-- 추가
		sqlStr = "Insert Into db_event.dbo.tbl_goodnote " & vbcrlf
		sqlStr = sqlStr + " (title,font_color,title_image,bg_color,contents_image,contents_title,contents,brand_button_color,brand_url,file_url,start_date,end_date,isusing)"  & vbcrlf
		sqlStr = sqlStr + " values('" & title  & "'"  & vbcrlf
		sqlStr = sqlStr + " ,'" & font_color &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & title_image &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & bg_color &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & contents_image &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & contents_title &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & contents &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & brand_button_color &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & brand_url &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & file_url &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & start_date &"'"  & vbcrlf
        sqlStr = sqlStr + " ,'" & end_date &"'"  & vbcrlf
		sqlStr = sqlStr + " ,'" & isusing &"')"
		dbget.Execute(sqlStr)

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
            response.End
        end if
    '===========================================================
	dbget.CommitTrans

	response.write "<script type='text/javascript'>"
	response.write "	opener.location.reload();"
	response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case "edit"
	dbget.beginTrans
    '===========================================================
    '-- 수정
        sqlStr = "UPDATE db_event.dbo.tbl_goodnote" & vbCrlf
        sqlStr = sqlStr + " SET title='" & title & "'" & vbCrlf
        sqlStr = sqlStr + " , font_color='" & font_color & "'" & vbCrlf
        sqlStr = sqlStr + " , title_image='" & title_image & "'" & vbCrlf
        sqlStr = sqlStr + " , bg_color='" & bg_color & "'" & vbCrlf
        sqlStr = sqlStr + " , contents_image='" & contents_image & "'" & vbCrlf
        sqlStr = sqlStr + " , contents_title='" & contents_title & "'" & vbCrlf
        sqlStr = sqlStr + " , contents='" & contents & "'" & vbCrlf
        sqlStr = sqlStr + " , brand_button_color='" & brand_button_color & "'" & vbCrlf
        sqlStr = sqlStr + " , brand_url='" & brand_url & "'" & vbCrlf
        sqlStr = sqlStr + " , file_url='" & file_url & "'" & vbCrlf
        sqlStr = sqlStr + " , start_date='" & start_date & "'" & vbCrlf
        sqlStr = sqlStr + " , end_date='" & end_date & "'" & vbCrlf
        sqlStr = sqlStr + " , isusing='" & isusing & "'" & vbCrlf
        sqlStr = sqlStr + " where idx=" & idx
        dbget.execute sqlStr

        if Err.Number <> 0 then
            dbget.RollBackTrans
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans

	response.write "<script type='text/javascript'>"
	response.write "	opener.location.reload();"
	response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case else
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->