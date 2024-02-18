<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/upche_qnacls.asp" -->
<%

dim boardqna
dim idx, mode, title, contents, userid, username, gubun, masterid

idx = request("idx")
mode = request("mode")
gubun = request("gubun")
masterid = request("masterid")
title = request("title")
contents = request("contents")
userid = session("ssBctId")
username = session("ssBctCname")

set boardqna = New CUpcheQnADetail

if (mode = "write") then
        boardqna.write masterid, gubun, title, contents, userid, username
        response.write "<script>location.replace('upche_qna_board_list.asp')</script>"
elseif (mode = "edit") then
        boardqna.modify idx, gubun, title, contents
        response.write "<script>location.replace('upche_qna_board_reply.asp?idx=" + idx + "')</script>"
elseif (mode = "del") then
        boardqna.del idx
        response.write "<script>location.replace('upche_qna_board_list.asp')</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->