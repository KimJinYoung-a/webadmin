<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/db2_manianewscls.asp" -->
<%

dim boardnotice
dim boarditem
dim id, mode, title, contents, isusing

id = request("id")
mode = request("mode")
title = request("title")
contents = request("contents")
isusing = request("isusing")


if (mode = "write") then
        set boardnotice = New CBoardNotice
        set boarditem = new CBoardNoticeItem

        boarditem.title = html2db(title)
        boarditem.contents = html2db(contents)
        boarditem.isusing = isusing
        boardnotice.write(boarditem)

        response.write "<script>location.replace('mania_notice_board_list.asp')</script>"
end if

if (mode = "modify") then
        set boardnotice = New CBoardNotice
        set boarditem = new CBoardNoticeItem

        boarditem.id = id
        boarditem.title = html2db(title)
        boarditem.contents = html2db(contents)
        boarditem.isusing = isusing
        boardnotice.modify(boarditem)

        response.write "<script>location.replace('mania_notice_board_list.asp')</script>"
end if

%>
<!-- #include virtual="/lib/db/db2close.asp" -->