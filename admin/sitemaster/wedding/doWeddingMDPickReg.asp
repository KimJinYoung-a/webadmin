<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim idx, itemid, upload_img, DispOrder, mode
	
	idx = requestCheckVar(request("idx"),3)
	itemid = requestCheckVar(request("itemid"),10)
	upload_img = requestCheckVar(request("upload_img"),128)
	DispOrder = requestCheckVar(request("DispOrder"),3)

	if idx="" then idx=0
	If idx=0 Then
		mode = "add"
	Else
		mode = "edit"
	End If

dim sqlStr

if (mode = "add") then
    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_wedding_md_pick]" + VbCrlf
    sqlStr = sqlStr + " (itemid, DispOrder, LastUser, upload_img)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + itemid + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + DispOrder + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" +  session("ssBctCname") + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + upload_img + "'" + VbCrlf
    sqlStr = sqlStr + " )"
    dbget.Execute sqlStr
elseif mode = "edit" then
   sqlStr = " update [db_sitemaster].[dbo].[tbl_wedding_md_pick]" + VbCrlf
   sqlStr = sqlStr + " set itemid='" + itemid + "'" + VbCrlf
   sqlStr = sqlStr + " ,DispOrder='" + DispOrder + "'" + VbCrlf
   sqlStr = sqlStr + " ,LastUser='" + session("ssBctCname") + "'" + VbCrlf
   sqlStr = sqlStr + " ,upload_img='" + upload_img + "'" + VbCrlf
   sqlStr = sqlStr + " where idx=" + cstr(idx)
   dbget.Execute sqlStr
end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>opener.location.reload();self.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->