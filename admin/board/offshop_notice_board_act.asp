<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/offshop_noticecls.asp" -->
<%

dim boardnotice
dim boarditem
dim id, mode, title, contents, yuhyostart, yuhyoend, malltype, noticetype

id = request("id")
malltype = request("malltype")
noticetype = request("noticetype")
mode = request("mode")
title = request("title")
contents = request("contents")
yuhyostart = request("yuhyostart")
yuhyoend = request("yuhyoend")



if (mode = "write") then
        set boardnotice = New CBoardNotice
        set boarditem = new CBoardNoticeItem

        boarditem.malltype = malltype
        boarditem.noticetype = noticetype
        boarditem.title = html2db(title)
        boarditem.contents = html2db(contents)
        boarditem.yuhyostart = yuhyostart
        boarditem.yuhyoend = yuhyoend

        boardnotice.write(boarditem)

        response.write "<script>location.replace('offshop_notice_board_list.asp')</script>"
end if

if (mode = "modify") then
        set boardnotice = New CBoardNotice
        set boarditem = new CBoardNoticeItem

        boarditem.id = id
        boarditem.malltype = malltype
        boarditem.noticetype = noticetype
        boarditem.title = html2db(title)
        boarditem.contents = html2db(contents)
        boarditem.yuhyostart = yuhyostart
        boarditem.yuhyoend = yuhyoend

        boardnotice.modify(boarditem)

        response.write "<script>location.replace('offshop_notice_board_list.asp')</script>"
end if

if (mode = "delete") then
        set boardnotice = New CBoardNotice

        boardnotice.delete(id)

        response.write "<script>location.replace('offshop_notice_board_list.asp')</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->