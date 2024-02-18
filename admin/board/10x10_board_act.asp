<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/10x10_boardcls.asp"-->
<%

dim boardnotice
dim boarditem
dim idx, mode, gubun, title, contents, userid, username

idx = request("idx")
mode = request("mode")
gubun = request("gubun")
title = html2db(request("title"))
contents = html2db(request("contents"))
userid = request("userid")
username = request("username")



if (mode = "write") then

		set boardnotice = New CHopeBoardDetail
        boardnotice.write gubun,userid,username,title,contents

        response.write "<script>alert('저장되었습니다.')</script>"
        response.write "<script>location.replace('10x10_board_list.asp')</script>"
end if

if (mode = "modify") then
        set boardnotice = New CHopeBoardDetail
        boardnotice.modify idx,gubun,userid,username,title,contents

        response.write "<script>alert('수정되었습니다.')</script>"
        response.write "<script>location.replace('10x10_board_list.asp')</script>"
end if

if (mode = "delete") then
        set boardnotice = New CHopeBoardDetail
        boardnotice.delete idx

        response.write "<script>alert('삭제되었습니다.')</script>"
        response.write "<script>location.replace('10x10_board_list.asp')</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->