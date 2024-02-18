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
Dim mode, sqlStr, idx
Dim gubun, subject, contents, isusing
menupos		= request("menupos")
mode		= request("mode")
gubun		= request("gubun")
subject		= html2db(Trim(request("subject")))
contents	= html2db(request("contents"))
isusing		= request("isusing")
idx			= request("idx")

Select Case mode
	Case "I"
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_between_noticefaq (gubun, subject, contents) VALUES "
		sqlStr = sqlStr & " ('"&gubun&"', '"&subject&"', '"&contents&"') "
		dbCTget.execute sqlStr
	Case "U"
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_noticefaq SET "
		sqlStr = sqlStr & " subject = '"&subject&"' "
		sqlStr = sqlStr & " ,contents = '"&contents&"' "
		sqlStr = sqlStr & " ,isusing = '"&isusing&"' "
		sqlStr = sqlStr & " WHERE idx = '"&idx&"' "
		dbCTget.execute sqlStr
End Select

Response.Write "<script>alert('저장되었습니다.');location.href='/admin/etc/between/noticefaq/notice_list.asp?menupos="&menupos&"';</script>"
dbCTget.close()
Response.End
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->