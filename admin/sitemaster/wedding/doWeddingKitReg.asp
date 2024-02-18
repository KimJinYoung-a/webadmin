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
Dim idx, itemid, upload_img1, upload_img2, mode, copy1, copy2, copy3, copy4, DispOrder
	
	idx = requestCheckVar(request("idx"),3)
	itemid = requestCheckVar(request("itemid"),10)
	copy1 = requestCheckVar(request("copy1"),128)
	copy2 = requestCheckVar(request("copy2"),128)
	copy3 = requestCheckVar(request("copy3"),128)
	copy4 = requestCheckVar(request("copy4"),128)
	upload_img1 = requestCheckVar(request("upload_img1"),128)
	upload_img2 = requestCheckVar(request("upload_img2"),128)
	DispOrder = requestCheckVar(request("DispOrder"),3)

	if idx="" then idx=0
	If idx=0 Then
		mode = "add"
	Else
		mode = "edit"
	End If

dim sqlStr

if (mode = "add") then
    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_wedding_kit]" + VbCrlf
    sqlStr = sqlStr + " (ItemID, Copy1, Copy2, Copy3, Copy4, upload_img1, upload_img2, LastUser, DispOrder)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + itemid + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + copy1 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + copy2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + copy3 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + copy4 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + upload_img1 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + upload_img2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" +  session("ssBctCname") + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + DispOrder + "'" + VbCrlf
	sqlStr = sqlStr + " )"
    dbget.Execute sqlStr
elseif mode = "edit" then
   sqlStr = " update [db_sitemaster].[dbo].[tbl_wedding_kit]" + VbCrlf
   sqlStr = sqlStr + " set itemid='" + itemid + "'" + VbCrlf
   sqlStr = sqlStr + " ,Copy1='" + copy1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Copy2='" + copy2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Copy3='" + copy3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Copy4='" + copy4 + "'" + VbCrlf
   sqlStr = sqlStr + " ,upload_img1='" + upload_img1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,upload_img2='" + upload_img2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,DispOrder='" + DispOrder + "'" + VbCrlf
   sqlStr = sqlStr + " ,LastUser='" + session("ssBctCname") + "'" + VbCrlf
   sqlStr = sqlStr + " where idx=" + cstr(idx)
   
   dbget.Execute sqlStr
end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>opener.location.reload();self.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->