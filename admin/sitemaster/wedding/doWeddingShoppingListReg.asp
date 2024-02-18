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
Dim WeddingStepID, itemid1, upload_img1, itemid2, upload_img2, itemid3, upload_img3
Dim mode, itemid4, upload_img4, itemid5, upload_img5, itemid6, upload_img6
	
	
	WeddingStepID = requestCheckVar(request("WeddingStepID"),3)
	itemid1 = requestCheckVar(request("itemid1"),10)
	upload_img1 = requestCheckVar(request("upload_img1"),128)
	itemid2 = requestCheckVar(request("itemid2"),10)
	upload_img2 = requestCheckVar(request("upload_img2"),128)
	itemid3 = requestCheckVar(request("itemid3"),10)
	upload_img3 = requestCheckVar(request("upload_img3"),128)
	itemid4 = requestCheckVar(request("itemid4"),10)
	upload_img4 = requestCheckVar(request("upload_img4"),128)
	itemid5 = requestCheckVar(request("itemid5"),10)
	upload_img5 = requestCheckVar(request("upload_img5"),128)
	itemid6 = requestCheckVar(request("itemid6"),10)
	upload_img6 = requestCheckVar(request("upload_img6"),128)

	if WeddingStepID="" then WeddingStepID=0
	If WeddingStepID=0 Then
		mode = "add"
	Else
		mode = "edit"
	End If

dim sqlStr

if (mode = "add") then

    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_wedding_shopping_list]" + VbCrlf
    sqlStr = sqlStr + " (itemid1, upload_img1, itemid2, upload_img2, itemid3, upload_img3, itemid4, upload_img4, itemid5, upload_img5, itemid6, upload_img6, LastUser)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + itemid1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + upload_img1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + itemid2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + upload_img2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + itemid3 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + upload_img3 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + itemid4 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + upload_img4 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + itemid5 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + upload_img5 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + itemid6 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + upload_img6 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" +  session("ssBctCname") + "'" + VbCrlf
    sqlStr = sqlStr + " )"
    dbget.Execute sqlStr

elseif mode = "edit" then
   sqlStr = " update [db_sitemaster].[dbo].[tbl_wedding_shopping_list]" + VbCrlf
   sqlStr = sqlStr + " set itemid1='" + itemid1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,upload_img1='" + upload_img1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,itemid2='" + itemid2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,upload_img2='" + upload_img2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,itemid3='" + itemid3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,upload_img3='" + upload_img3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,itemid4='" + itemid4 + "'" + VbCrlf
   sqlStr = sqlStr + " ,upload_img4='" + upload_img4 + "'" + VbCrlf
   sqlStr = sqlStr + " ,itemid5='" + itemid5 + "'" + VbCrlf
   sqlStr = sqlStr + " ,upload_img5='" + upload_img5 + "'" + VbCrlf
   sqlStr = sqlStr + " ,itemid6='" + itemid6 + "'" + VbCrlf
   sqlStr = sqlStr + " ,upload_img6='" + upload_img6 + "'" + VbCrlf
   sqlStr = sqlStr + " ,LastUser='" + session("ssBctCname") + "'" + VbCrlf
   sqlStr = sqlStr + " where WeddingStepID=" + cstr(WeddingStepID)
   
   dbget.Execute sqlStr
end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>opener.location.reload();self.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->