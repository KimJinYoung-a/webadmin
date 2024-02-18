<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [업체]게시판 
' Hieditor : 2015.05.27 이상구 생성
'			 2020.03.12 한용민 수정(이메일발송. 메일러로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<%
dim mailcontent, boardqna, boarditem, idx, mode, replytitle, replycontents, replyuser, isCsMailSend
dim page, SearchKey, SearchString, gubun, replyYn, param, workerid, selDate, sDate, eDate
	page = requestcheckvar(getNumeric(Request("page")),10)
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	gubun = requestcheckvar(getNumeric(Request("gubun")),2)
	replyYn = requestcheckvar(Request("replyYn"),1)
	workerid = requestcheckvar(Request("workerid"),32)
	selDate		= requestCheckVar(Request("selDate"),1)
	sDate 		= requestCheckVar(Request("sDate"),10)
	eDate 		= requestCheckVar(Request("eDate"),10)
	isCsMailSend = (request.Form("csmailsend")="on")

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&gubun=" & gubun & "&replyYn=" & replyYn & "&selDate=" & selDate & "&sDate=" & sDate & "&eDate=" & eDate

	idx = requestcheckvar(getNumeric(Request("idx")),10)
	mode 		= requestCheckVar(Request("mode"),32)
	replytitle = html2db(request("replytitle"))
	replycontents = html2db(request("replycontents"))
	replyuser = session("ssBctId")

if (checkNotValidHTML(replytitle) = True) then
	Alert_return("제목에는 HTML을 사용하실 수 없습니다.")
	dbget.Close : response.end
end If

if (checkNotValidHTML(replycontents) = True) then
	Alert_return("내용에는 HTML을 사용하실 수 없습니다.")
	dbget.Close : response.end
end If

if (mode = "reply") then
	set boardqna = New CUpcheQnADetail
		boardqna.reply idx, replytitle, replycontents, replyuser

		' 답변시 이메일 발송	' 2020-03-12 한용민 생성
		If (isCsMailSend) then
			Call SendUpheBoardMail(idx)
		end if
	set boardqna = Nothing

	response.write "<script type='text/javascript'>"
	response.write "	alert('저장되었습니다.');"
	response.write "	location.replace('upche_qna_board_reply.asp?idx=" + idx + "&page=" + page + Param + "');"
	response.write "	</script>"

elseif (mode = "edit") then
	set boardqna = New CUpcheQnADetail
	boardqna.changeworker idx, workerid
	set boardqna = Nothing
	
	Response.Write "<script>alert('저장되었습니다.');location.href='upche_qna_board_list.asp?page=" + page + Param + "';</script>"
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->