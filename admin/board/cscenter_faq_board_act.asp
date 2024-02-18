<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/boardfaqcls.asp" -->
<%

dim boardfaq
dim boarditem
dim id, mode, divcd, subcd, title, contents

id = request("id")
mode = request("mode")
divcd = request("divcd")
subcd = request("subcd")
title = request("title")
contents = request("contents")


if (mode = "write") then
        set boardfaq = New CBoardFAQ
        set boarditem = new CBoardFAQItem

        boarditem.divcd = divcd
        boarditem.subcd = subcd
        boarditem.title = title
        boarditem.contents = html2db(contents)

        boardfaq.write(boarditem)

        response.write "<script>location.replace('cscenter_faq_board_list.asp')</script>"
end if

if (mode = "modify") then
        set boardfaq = New CBoardFAQ
        set boarditem = new CBoardFAQItem

        boarditem.id = id
        boarditem.divcd = divcd
        boarditem.subcd = subcd
        boarditem.title = title
        boarditem.contents = html2db(contents)


        boardfaq.modify(boarditem)

        response.write "<script>location.replace('cscenter_faq_board_list.asp')</script>"
end if

if (mode = "delete") then
        set boardfaq = New CBoardFAQ

        boardfaq.delete(id)

        response.write "<script>location.replace('cscenter_faq_board_list.asp')</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->