<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description : 매거진 카테고리 관리 처리 페이지
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
dim Cidx
dim sqlStr
Dim catecodename, mode
Dim arrcatecodename, arrcatecodenamecnt , i 

Cidx			= RequestCheckVar(request("cidx"),10)
catecodename	= RequestCheckVar(request("catecodename"),500)

mode = RequestCheckVar(request("mode"),11)
  	if catecodename <> "" then
		if checkNotValidHTML(catecodename) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
'arrcatecodename = Split(catecodename,",")
'arrcatecodenamecnt = UBound(arrcatecodename)

'Response.end
if (mode = "catecode") then
	sqlStr = sqlStr & " insert into db_academy.dbo.tbl_academy_magazine_catecode (catename) values ('"&catecodename&"') " & vbCrLf
	
	
'	sqlStr = " update db_academy.dbo.tbl_academy_magazine_catecode set isusing='N'" & vbCrLf
'	response.write sqlStr
'	If arrcatecodenamecnt > 0 Then
'		For i = 0 To arrcatecodenamecnt
'			If Trim(arrcatecodename(i)) <> "" then
'			sqlStr = sqlStr & " insert into db_academy.dbo.tbl_academy_magazine_catecode (catecode) values ('"&Trim(arrcatecodename(i))&"') " & vbCrLf
'			End If 
'		Next 
'	Else
'			sqlStr = sqlStr & " insert into db_academy.dbo.tbl_academy_magazine_catecode (catecode) values ('"&catecodename&"') " & vbCrLf
'	End If 
'	Response.write sqlStr
'	Response.end
    dbACADEMYget.Execute sqlStr
elseif (mode = "catecodedel") then
	sqlStr = " update [db_academy].[dbo].[tbl_academy_magazine_catecode] set isusing='N' where idx='" & Cidx & "' "
'	Response.write sqlStr
'	Response.end
    dbACADEMYget.Execute sqlStr
End If 

dim referer
referer = request.ServerVariables("HTTP_REFERER")
if mode = "catecode" then
	response.write "<script>alert('저장되었습니다.');</script>"
else
	response.write "<script>alert('삭제되었습니다.');</script>"
end if
response.write "<script>location.href='/academy/magazine/lib/pop_catecodeReg.asp'</script>"

%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->