<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description : 매거진 태그처리 페이지
' Hieditor : 2016-03-04 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx
dim sqlStr
Dim tagname, mode
Dim arrtagname, arrtagnamecnt , i 

idx		= RequestCheckVar(request("idx"),10)

tagname	= RequestCheckVar(request("tagname"),500)

mode = RequestCheckVar(request("mode"),10)
if tagname <> "" then
	if checkNotValidHTML(tagname) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
arrtagname = Split(tagname,",")
arrtagnamecnt = UBound(arrtagname)

'Response.end

if (mode = "tag") then

	sqlStr = " delete from db_academy.dbo.tbl_academy_magazine_keyword where vidx = '"& idx &"'" & vbCrLf
'	response.write sqlStr
	If arrtagnamecnt > 0 Then
		For i = 0 To arrtagnamecnt
			If Trim(arrtagname(i)) <> "" then
			sqlStr = sqlStr & " insert into db_academy.dbo.tbl_academy_magazine_keyword (vidx, searchkw) values ( '"& idx &"','"&Trim(arrtagname(i))&"') " & vbCrLf
			End If 
		Next 
	Else
			sqlStr = sqlStr & " insert into db_academy.dbo.tbl_academy_magazine_keyword (vidx, searchkw) values ( '"& idx &"','"&tagname&"') " & vbCrLf
	End If 
'	Response.write sqlStr
'	Response.end
    dbACADEMYget.Execute sqlStr

End If 

dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.href='" & manageUrl & "/academy/magazine/lib/pop_tagReg.asp?idx=" + Cstr(idx) +"&reload=on'</script>"

%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->