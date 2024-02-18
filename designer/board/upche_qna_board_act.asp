<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<%

dim boardqna
dim idx, mode, title, contents, userid, username, gubun, masterid, workerid

idx = requestCheckVar(request("idx"),10)
mode = requestCheckVar(request("mode"),30)
gubun = requestCheckVar(request("gubun"),50)
masterid = requestCheckVar(request("masterid"),50)
title = html2db(request("title"))
contents = html2db(request("contents"))
userid = session("ssBctId")
username = session("ssBctCname")
workerid = requestCheckVar(request("workerid"),50)

if (checkNotValidHTML(title) = True) Then
	response.write "<script>alert('제목에는 HTML을 사용하실 수 없습니다.');</script>"
	dbget.close()	:	response.End
End If

if (checkNotValidHTML(contents) = True) Then
	response.write "<script>alert('내용에는 HTML을 사용하실 수 없습니다.');</script>"
	dbget.close()	:	response.End
End If


dim sql, tmp
IF idx <> "" Then
	sql = "select count(*) from [db_board].[dbo].tbl_upche_qna where idx = '" & idx & "' and userid = '" & userid & "'"
	rsget.open sql,dbget,1
	tmp = rsget(0)
	rsget.close
	If tmp < 1 Then
		Response.Write "<script>alert('잘못된 접근입니다.');top.location.href='/';</script>"
		dbget.close()
		Response.End
	End IF
End IF

set boardqna = New CUpcheQnADetail

if (mode = "write") then
        boardqna.write masterid, gubun, title, contents, userid, username, workerid
        response.write "<script>location.replace('upche_qna_board_list.asp')</script>"
elseif (mode = "edit") then
        boardqna.modify idx, gubun, title, contents, workerid
        response.write "<script>location.replace('upche_qna_board_reply.asp?idx=" + idx + "')</script>"
elseif (mode = "del") then
        boardqna.del idx
        response.write "<script>location.replace('upche_qna_board_list.asp')</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->