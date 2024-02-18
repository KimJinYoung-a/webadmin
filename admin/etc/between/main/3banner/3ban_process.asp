<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<%
Dim mode, sqlStr
Dim imgurl, imglink, sortno, startdate, enddate, regdate, adminid, isusing, idx, gender
mode		= request("mode")
idx			= request("idx")
imgurl		= request("ban")
imglink		= request("imglink")
sortno		= request("sortno")
startdate	= request("startdate")
enddate		= request("enddate")
isusing		= request("isusing")
gender		= request("gender")

startdate	= startdate & " 00:00:00"
enddate		= enddate & " 23:59:59"

Select Case mode
	Case "I"
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_between_main_3banner (gender, imgurl, imglink, sortno, startdate, enddate, regdate, adminid, isusing) VALUES "
		sqlStr = sqlStr & " ('"&gender&"', '"&imgurl&"', '"&imglink&"', '"&sortno&"', '"&startdate&"', '"&enddate&"', getdate(), '"&session("ssBctId")&"', 'Y') "
		dbCTget.execute sqlStr
	Case "U"
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_main_3banner SET "
		sqlStr = sqlStr & " gender = '"&gender&"', "
		sqlStr = sqlStr & " imgurl = '"&imgurl&"', "
		sqlStr = sqlStr & " imglink = '"&imglink&"', "
		sqlStr = sqlStr & " sortno = '"&sortno&"', "
		sqlStr = sqlStr & " startdate = '"&startdate&"', "
		sqlStr = sqlStr & " enddate = '"&enddate&"', "
		sqlStr = sqlStr & " lastupdate = getdate(), "
		sqlStr = sqlStr & " lastadminid = '"&session("ssBctId")&"', "
		sqlStr = sqlStr & " isusing = '"&isusing&"' "
		sqlStr = sqlStr & " WHERE idx = '"&idx&"' "
		dbCTget.execute sqlStr
End Select
Response.Write "<script language='javascript'>alert('저장 되었습니다.');location.href='/admin/etc/between/main/3banner/index.asp?menupos="&menupos&"';</script>"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->