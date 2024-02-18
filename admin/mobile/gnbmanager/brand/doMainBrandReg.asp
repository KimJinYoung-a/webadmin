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
Dim idx, BrandID, BrandName, MainCopy, Main_Image
Dim StartDate, EndDate, DispOrder, Isusing, mode, itemID
	
	
	idx = requestCheckVar(request("idx"),10)
	BrandID = requestCheckVar(request("BrandID"),32)
	BrandName = requestCheckVar(request("BrandName"),32)
	MainCopy = requestCheckVar(request("MainCopy"),128)
	Main_Image = requestCheckVar(request("Main_Image"),128)
	
	itemID = requestCheckVar(request("itemID"),10)
	StartDate = requestCheckVar(request("StartDate"),10)
	EndDate = requestCheckVar(request("EndDate"),10)
	DispOrder = requestCheckVar(request("DispOrder"),3)
	Isusing = requestCheckVar(request("Isusing"),1)

	if idx="" then idx=0
	If idx=0 Then
	mode = "add"
	Else
	mode = "edit"
	End If

dim sqlStr

If itemID > "0" Then
	sqlStr = "select top 1 basicimage "
	sqlStr = sqlStr + " from [db_item].[dbo].[tbl_item]"
	sqlStr = sqlStr + " where itemid='" + CStr(ItemID) + "'"
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof then
		Main_Image = webImgUrl & "/image/basic/" + GetImageSubFolderByItemid(ItemID) + "/"  + rsget("basicimage")
	end if
	rsget.Close
End If

if (mode = "add") then

    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_mobile_gnb_brand]" + VbCrlf
    sqlStr = sqlStr + " (makerid, BrandName, SubCopy, brandIMG, itemID, StartDate, EndDate, DispOrder, Isusing, RegUser)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + BrandID + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + BrandName + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + MainCopy + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Main_Image + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + itemID + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + StartDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + EndDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + DispOrder + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" +  session("ssBctCname") + "'" + VbCrlf
    sqlStr = sqlStr + " )"
	'Response.write sqlStr
	'Response.end
    dbget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('[db_sitemaster].[dbo].[tbl_mobile_gnb_brand]') as idx"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
		idx = rsget("idx")
	end if
	rsget.close

elseif mode = "edit" then
   sqlStr = " update [db_sitemaster].[dbo].[tbl_mobile_gnb_brand]" + VbCrlf
   sqlStr = sqlStr + " set makerid='" + BrandID + "'" + VbCrlf
   sqlStr = sqlStr + " ,BrandName='" + BrandName + "'" + VbCrlf
   sqlStr = sqlStr + " ,brandIMG='" + Main_Image + "'" + VbCrlf
   sqlStr = sqlStr + " ,SubCopy='" + MainCopy + "'" + VbCrlf
   sqlStr = sqlStr + " ,itemID='" + itemID + "'" + VbCrlf
   sqlStr = sqlStr + " ,StartDate='" + StartDate + "'" + VbCrlf
   sqlStr = sqlStr + " ,EndDate='" + EndDate + "'" + VbCrlf
   sqlStr = sqlStr + " ,DispOrder='" + DispOrder + "'" + VbCrlf
   sqlStr = sqlStr + " ,Isusing='" + Isusing + "'" + VbCrlf
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