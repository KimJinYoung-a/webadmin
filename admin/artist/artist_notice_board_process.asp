<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'History	:  2012.03.22 김진영
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/boardnoticecls.asp" -->
<%
Dim sql, isusing
Dim idx, mode, title, contents, noticetype, fixyn
Dim SearchKey, SearchString, param, page, listtype, menupos

idx 		= requestCheckVar(request("idx"),12)
mode 		= requestCheckVar(request("mode"),128)
title 		= html2db(request("title"))
contents 	= html2db(request("contents"))
fixyn 		= requestCheckVar(request("fixyn"),8)
menupos 	= requestCheckVar(request("menupos"),128)
isusing 	= requestCheckVar(request("isusing"),8)

If (mode = "write") Then
	sql = " insert into db_contents.dbo.tbl_artist_notice_board (title, contents, regdate, fixyn) "
	sql = sql + " values('"&title&"', '"&contents&"', getdate(), '"&fixyn&"') "
	dbget.execute sql

	Response.Write "<script>alert('저장되었습니다.');location.href='/admin/artist/artist_notice_board_list.asp?menupos="&menupos&"';</script>"
	dbget.close()
	Response.End	
End If

If (mode = "modify") Then
	sql = "update db_contents.dbo.tbl_artist_notice_board set " + VbCrlf
	sql = sql + " title = '"&title&"'," + VbCrlf
	sql = sql + " contents = '"&contents&"'," + VbCrlf
	sql = sql + " fixyn = '"&fixyn&"', " + VbCrlf
	sql = sql + " isusing = '"&isusing&"'" + VbCrlf
	sql = sql + " where (idx = '"&idx&"') "
	dbget.execute sql

	Response.Write "<script>alert('수정되었습니다.');location.href='/admin/artist/artist_notice_board_list.asp?menupos="&menupos&"';</script>"
	dbget.close()
	Response.End
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->