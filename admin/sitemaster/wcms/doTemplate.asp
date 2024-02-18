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
Dim siteDiv, pageDiv, sqlStr
Dim tplIdx, tplType, tplName, isTimeUse, isIconUse, isSubNumUse, isTopImgUse, isTopLinkUse
Dim isImageUse, isTextUse, isLinkUse, isItemUse, isVideoUse, isBGColorUse, isImgDescUse, isExtDataUse, tplinfoDesc, tplSortNo

siteDiv			= request("site")
pageDiv			= request("pageDiv")
tplIdx			= request("tplIdx")
tplType			= request("tplType")
tplName			= request("tplName")
isTimeUse		= request("isTimeUse")
isIconUse		= request("isIconUse")
isSubNumUse		= request("isSubNumUse")
isTopImgUse		= request("isTopImgUse")
isTopLinkUse	= request("isTopLinkUse")
isImageUse		= request("isImageUse")
isTextUse		= request("isTextUse")
isLinkUse		= request("isLinkUse")
isItemUse		= request("isItemUse")
isVideoUse		= request("isVideoUse")
isBGColorUse	= request("isBGColorUse")
isImgDescUse	= request("isImgDescUse")
isExtDataUse	= request("isExtDataUse")
tplinfoDesc		= request("tplinfoDesc")
tplSortNo		= request("tplSortNo")

if (tplIdx<>"") then
    sqlStr = " update [db_sitemaster].[dbo].tbl_cms_template" + VbCrlf
    sqlStr = sqlStr + " set pageDiv='" + pageDiv + "'" + VbCrlf
    sqlStr = sqlStr + " ,tplType='" + tplType + "'" + VbCrlf
    sqlStr = sqlStr + " ,tplName='" + html2db(tplName) + "'" + VbCrlf
    sqlStr = sqlStr + " ,isTimeUse='" + isTimeUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isIconUse='" + isIconUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isSubNumUse='" + isSubNumUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isTopImgUse='" + isTopImgUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isTopLinkUse='" + isTopLinkUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isImageUse='" + isImageUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isTextUse='" + isTextUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isLinkUse='" + isLinkUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isItemUse='" + isItemUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isVideoUse='" + isVideoUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isBGColorUse='" + isBGColorUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isExtDataUse='" + isExtDataUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,isImgDescUse='" + isImgDescUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,tplinfoDesc='" + html2db(tplinfoDesc) + "'" + VbCrlf
    sqlStr = sqlStr + " ,tplSortNo='" + tplSortNo + "'" + VbCrlf
    sqlStr = sqlStr + " where tplIdx=" + CStr(tplIdx) + VbCrlf
    
    dbget.Execute sqlStr
else
    sqlStr = " insert into [db_sitemaster].[dbo].tbl_cms_template" + VbCrlf
    sqlStr = sqlStr + " (siteDiv, pageDiv, tplType, tplName, isTimeUse, isIconUse, isSubNumUse, isTopImgUse, isTopLinkUse"+ VbCrlf
	sqlStr = sqlStr + " ,isImageUse, isTextUse, isLinkUse, isItemUse, isVideoUse, isBGColorUse, isImgDescUse, isExtDataUse, tplinfoDesc, tplSortNo)"+ VbCrlf
    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " '" + siteDiv + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + pageDiv + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + tplType + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + html2db(tplName) + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isTimeUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isIconUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isSubNumUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isTopImgUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isTopLinkUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isImageUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isTextUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isLinkUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isItemUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isVideoUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isBGColorUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isImgDescUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isExtDataUse + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + html2db(tplinfoDesc) + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + tplSortNo + "'" + VbCrlf
    sqlStr = sqlStr + " )" + VbCrlf
    
    dbget.Execute sqlStr
end if


dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->