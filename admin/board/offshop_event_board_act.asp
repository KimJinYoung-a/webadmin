<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/offshop_event_boardcls.asp" -->
<%

dim id, mode, gubun,userid,title,contents,enddate,isusing

id = request("id")
mode = request("mode")
gubun = request("gubun")
userid = request("userid")
title = request("title")
contents = request("contents")
enddate = request("enddate")
isusing = request("isusing")

dim sql

if (mode = "write") then

		sql = " insert into [db_board].[10x10].tbl_offshop_event_board(gubun, userid, title, contents, enddate) "
		sql = sql + " values('" + gubun + "','" + userid + "', '" + title + "', '" + contents + "','" + enddate + "') "
		rsget.Open sql, dbget, 1


        response.write "<script>alert('저장되었습니다.')</script>"
        response.write "<script>location.replace('offshop_event_board_list.asp')</script>"

elseif  (mode = "edit") then

		sql = "update [db_board].[10x10].tbl_offshop_event_board " + VbCRlf
		sql = sql + " set gubun = '" + gubun + "'," + VbCRlf
		sql = sql + " title = '" + title + "'," + VbCRlf
		sql = sql + " contents = '" + contents + "', " + VbCRlf
		sql = sql + " enddate = '" + enddate + "', " + VbCRlf
		sql = sql + " isusing = '" + isusing + "' " + VbCRlf
		sql = sql + " where (id = " + id + ") " + VbCRlf
		rsget.Open sql, dbget, 1
        response.write "<script>alert('수정되었습니다.')</script>"
        response.write "<script>location.replace('offshop_event_board_list.asp')</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->