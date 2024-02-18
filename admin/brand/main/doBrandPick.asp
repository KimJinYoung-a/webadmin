<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : doBrandPick.asp
' Discription : ºê·£µå ÇÈ ÄÁÅÙÃ÷ µî·Ï
' History : 2019.11.08 Á¤ÅÂÈÆ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
Dim mode, idx, imagepath, linkpath, image_order, sqlStr, makerid, isusing, refer

refer = request.ServerVariables("HTTP_REFERER")

idx			= requestCheckVar(request.Form("idx"),10)
makerid 	= requestCheckVar(request.Form("makerid"),32)
imagepath 	= requestCheckVar(request.Form("imagepath"),128)
linkpath 	= requestCheckVar(request.Form("linkpath"),128)
isusing 	= requestCheckVar(request.Form("isusing"),1)
image_order	= requestCheckVar(request.Form("image_order"),10)

If idx="" Then 
	idx=0
	mode = "add"
Else
	mode = "edit"
End If


If (Mode = "add") Then
	sqlStr = ""
	sqlStr = sqlStr & " INSERT INTO db_brand.dbo.tbl_2013brand_image "
	sqlStr = sqlStr & " (makerid, imagepath, linkpath, isusing, regdate, gubun, image_order) VALUES "
	sqlStr = sqlStr & " ('"&makerid&"', '" & imagepath & "', '"&linkpath&"', 'Y', getdate(), '2', '"&image_order&"') "
	dbget.Execute sqlStr

	sqlStr = "SELECT IDENT_CURRENT('db_brand.dbo.tbl_2013brand_image') as idx"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
		idx = rsget("idx")
	End If
	rsget.close
ElseIf mode = "edit" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_brand.dbo.tbl_2013brand_image SET "
	sqlStr = sqlStr & " linkpath='" + linkpath + "'"
	sqlStr = sqlStr & " ,makerid='" + makerid + "'"
    sqlStr = sqlStr & " ,imagepath='" + imagepath + "'"
	sqlStr = sqlStr & " ,isusing='" + isusing + "'"
	sqlStr = sqlStr & " ,image_order=" + image_order + ""
	sqlStr = sqlStr & " WHERE idx=" + CStr(idx)
	dbget.Execute sqlStr
End If

response.write "<script>"
response.write "	document.domain ='10x10.co.kr';"
response.write "	alert('OK');"
response.write "	opener.location.reload();"
response.write "	self.close();"
response.write "</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->