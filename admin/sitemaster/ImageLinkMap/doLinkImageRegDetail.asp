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
Dim idx, title, Link_Image, Isusing
Dim mode, masterIdx, posX, posY, IconType, ItemID
	
	masterIdx = requestCheckVar(request("masteridx"),10)
	idx = requestCheckVar(request("idx"),10)
    posX = RequestCheckVar(request("posX"),256)
    posY = requestCheckVar(request("posY"),256)
    Isusing = requestCheckVar(request("Isusing"),1)
    IconType = requestCheckVar(request("IconType"),1)
    ItemID = requestCheckVar(request("ItemID"),256)

	if idx="" then idx=0

	If idx=0 Then
	    mode = "add"
	Else
	    mode = "edit"
	End If

dim sqlStr


if (mode = "add") then

    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_ImageLink_Detail]" + VbCrlf
    sqlStr = sqlStr + " (MasterIdx, XValue, YValue, ItemId, IconType, IsUsing, RegUser, ModifyUser, RegDate, LastUpDate)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + masterIdx + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + posX + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + posY + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + ItemID + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + IconType + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + IsUsing + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Session("ssBctId") + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Session("ssBctId") + "'" + VbCrlf
    sqlStr = sqlStr + " ,getdate()" + VbCrlf
    sqlStr = sqlStr + " ,getdate()" + VbCrlf
    sqlStr = sqlStr + " )"
    dbget.Execute sqlStr

elseif mode = "edit" then

   sqlStr = " update [db_sitemaster].[dbo].[tbl_ImageLink_Detail]" + VbCrlf
   sqlStr = sqlStr + " set ItemID='" + ItemID + "'" + VbCrlf
   sqlStr = sqlStr + " ,IconType='" + IconType + "'" + VbCrlf
   sqlStr = sqlStr + " ,IsUsing='" + IsUsing + "'" + VbCrlf
   sqlStr = sqlStr + " ,ModifyUser='" + Session("ssBctId") + "'" + VbCrlf
   sqlStr = sqlStr + " ,LastUpDate=getdate()" + VbCrlf
   sqlStr = sqlStr + " where idx=" + cstr(idx)   
   dbget.Execute sqlStr

end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>opener.location.reload();self.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->