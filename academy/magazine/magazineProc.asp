<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description :  핑거스 아카데미 매거진 등록,수정 처리 페이지
' History : 2016-03-04 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx, listimg, viewimg1, viewimg2, viewimg3, viewtitle , viewtext1, viewtext2, viewtext3, startdate , state , mode, viewno, videourl, classcode, catecode, isusing

idx			= RequestCheckVar(request("idx"),10)
mode		= RequestCheckVar(request("mode"),10)
state		= RequestCheckVar(request("state"),10)
viewno		= RequestCheckVar(request("viewno"),10)
isusing	= requestCheckvar(request("isusing"),2)
listimg	= RequestCheckVar(request("listimg"),100)
catecode	= requestCheckvar(request("catecode"),10)
viewimg1	= RequestCheckVar(request("viewimg1"),100)
viewimg2	= RequestCheckVar(request("viewimg2"),100)
viewimg3	= RequestCheckVar(request("viewimg3"),100)
videourl	= RequestCheckVar(request("videourl"),150)
viewtitle	= RequestCheckVar(request("viewtitle"),50)
startdate	= RequestCheckVar(request("startdate"),10)
viewtext1	= RequestCheckVar(request("viewtext1"),800)
viewtext2	= RequestCheckVar(request("viewtext2"),800)
viewtext3	= RequestCheckVar(request("viewtext3"),800)
classcode	= requestCheckvar(request("classcode"),255)
if listimg <> "" then
	if checkNotValidHTML(listimg) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if viewimg1 <> "" then
	if checkNotValidHTML(viewimg1) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if viewimg2 <> "" then
	if checkNotValidHTML(viewimg2) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if viewimg3 <> "" then
	if checkNotValidHTML(viewimg3) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if videourl <> "" then
	if checkNotValidHTML(videourl) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if viewtitle <> "" then
	if checkNotValidHTML(viewtitle) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if viewtext1 <> "" then
	if checkNotValidHTML(viewtext1) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if viewtext2 <> "" then
	if checkNotValidHTML(viewtext2) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if classcode <> "" then
	if checkNotValidHTML(classcode) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if viewtext3 <> "" then
	if checkNotValidHTML(viewtext3) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if


if idx = "" then
	idx = 0
end If

if idx = 0 then
	''신규 등록
	mode = "add"
else
	''수정
	mode = "edit"
end if

dim sqlStr

if (mode = "add") then
''신규 등록
    sqlStr = " insert into [db_academy].[dbo].tbl_academy_magazine" + VbCrlf
    sqlStr = sqlStr + " (listimg, viewtitle, viewtext1, viewtext2, viewtext3, viewimg1, viewimg2, viewimg3, startdate, isusing, state, viewno, videourl, classcode, catecode)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + listimg + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewtitle + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewtext1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewtext2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewtext3 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewimg1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewimg2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewimg3 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + startdate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + state + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewno + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + videourl + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + classcode + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + catecode + "'" + VbCrlf
    sqlStr = sqlStr + " )"
'	response.write sqlStr
'	response.end
    dbACADEMYget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('[db_academy].[dbo].tbl_academy_magazine') as idx"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	If Not rsACADEMYget.Eof then
		idx = rsACADEMYget("idx")
	end if
	rsACADEMYget.close

	sqlStr = " update  [db_academy].[dbo].[tbl_academy_magazine_keyword] set vidx='" & idx & "' where vidx=0"
	dbACADEMYget.Execute sqlStr

elseif mode = "edit" Then
''수정
   sqlStr = " update  [db_academy].[dbo].tbl_academy_magazine " + VbCrlf
   sqlStr = sqlStr + " set " + VbCrlf
   sqlStr = sqlStr + " state='" + state + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewno='" + viewno + "'" + VbCrlf
   sqlStr = sqlStr + " ,listimg='" + listimg + "'" + VbCrlf
   sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewimg1='" + viewimg1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewimg2='" + viewimg2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewimg3='" + viewimg3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,videourl='" + videourl + "'" + VbCrlf
   sqlStr = sqlStr + " ,catecode='" + catecode + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewtext1='" + viewtext1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewtext2='" + viewtext2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewtext3='" + viewtext3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewtitle='" + viewtitle + "'" + VbCrlf
   sqlStr = sqlStr + " ,startdate='" + startdate + "'" + VbCrlf
   sqlStr = sqlStr + " ,classcode='" + classcode + "'" + VbCrlf
   sqlStr = sqlStr + " where vidx=" + CStr(idx)
   dbACADEMYget.Execute sqlStr
end if

'dim referer
'referer = request.ServerVariables("HTTP_REFERER")
'response.write "<script>alert('저장되었습니다.');</script>"
'response.write "<script>location.href='/academy/magazine/popmagazineEdit.asp?idx=" + Cstr(idx) + "&reload=on'</script>"
%>
<script language = "javascript">
	alert("저장되었습니다.");
	opener.location.reload();
	self.close();
</script>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->