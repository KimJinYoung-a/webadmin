<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/SitemasterClass/ImgCommentCls.asp"-->
<%

dim mode,reviewid,itemid
dim viewDate
dim rd_nousing
dim iconName,imageName,imageCName,imageDName

mode= requestCheckVar(request("mode"),4)
reviewid = requestCheckVar(request("reviewid"),5)
itemid = requestCheckVar(request("itemid"),10)

viewDate = requestCheckVar(request("vDate"),10)
rd_nousing = request("rd_nousing")

iconName = requestCheckVar(request("iconName"),32)
imageName = requestCheckVar(request("imageName"),32)
imageCName = requestCheckVar(request("imageCName"),32)
imageDName = requestCheckVar(request("imageDName"),32)


dim strSQL


IF mode="add" then
	strSQL =" INSERT INTO db_sitemaster.dbo.tbl_ImageComment (itemid,image,imageConfirm,imageDown,icon,viewdate,isUsing) " &_
			" VALUES ('" & itemid & "','" & imageName & "','" & imageCName & "','"& imageDName &"','" & iconName & "','" & CStr(viewdate) & "','" & rd_nousing & "')"
Else
	strSQL =" UPDATE db_sitemaster.dbo.tbl_ImageComment " &_
			" SET itemid ='" & itemid & "'" &_
			" , image = '" & imageName & "' " &_
			" , imageconfirm = '" & imageCName & "' " &_
			" , imageDown = '" & imageDName & "' " &_
			" , icon = '" & iconName & "'" &_
			" , viewdate = '" & viewdate & "'" &_
			" , isUsing='" & rd_nousing & "'" &_
			" WHERE idx='" & reviewid & "'"
End If

'response.write strSQL
'dbget.close()	:	response.End

dbget.beginTrans
dbget.execute(strSQL)

IF Err.Number >0 then
	dbget.rollbackTrans
	response.write "<script>alert('오류발생.');history.go(-1);</script>"
Else
	dbget.commitTrans
	response.write "<script>alert('저장되었습니다.');history.go(-1);</script>"

end if

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
