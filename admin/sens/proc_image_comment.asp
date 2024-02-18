<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.04.15 한용민 생성
'	Description : 감성엽서
'#######################################################
%>
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
<!-- #include virtual="/lib/classes/sens/image_commentcls.asp"-->
<%

dim mode,reviewid,itemid
dim yyyy1,mm1,dd1,viewdate
dim rd_nousing
dim iconName,imageName,imageCName

mode= request("mode")
reviewid = request("reviewid")
itemid = request("itemid")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

viewdate = yyyy1&"-"&mm1&"-"&dd1'dateserial(yyyy1,mm1,dd1)

rd_nousing = request("rd_nousing")

iconName = request("iconName")
imageName = request("imageName")
imageCName = request("imageCName")


dim strSQL


IF mode="add" then
	strSQL =" INSERT INTO db_contents.dbo.tbl_sens_postcard (itemid,image,imageConfirm,icon,viewdate,isUsing) " &_
			" VALUES ('" & itemid & "','" & imageName & "','" & imageCName & "','" & iconName & "','" & CStr(viewdate) & "','" & rd_nousing & "')"
Else
	strSQL =" UPDATE db_contents.dbo.tbl_sens_postcard " &_
			" SET itemid ='" & itemid & "'" &_
			" , image = '" & imageName & "' " &_
			" , imageconfirm = '" & imageCName & "' " &_
			" , icon = '" & iconName & "'" &_
			" , viewdate = '" & viewdate & "'" &_
			" , isUsing='" & rd_nousing & "'" &_
			" WHERE idx='" & reviewid & "'" 
End If

response.write strSQL
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