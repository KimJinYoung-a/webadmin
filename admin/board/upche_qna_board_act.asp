<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [��ü]�Խ��� 
' Hieditor : 2015.05.27 �̻� ����
'			 2020.03.12 �ѿ�� ����(�̸��Ϲ߼�. ���Ϸ��� ����)
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
	Alert_return("���񿡴� HTML�� ����Ͻ� �� �����ϴ�.")
	dbget.Close : response.end
end If

if (checkNotValidHTML(replycontents) = True) then
	Alert_return("���뿡�� HTML�� ����Ͻ� �� �����ϴ�.")
	dbget.Close : response.end
end If

if (mode = "reply") then
	set boardqna = New CUpcheQnADetail
		boardqna.reply idx, replytitle, replycontents, replyuser

		' �亯�� �̸��� �߼�	' 2020-03-12 �ѿ�� ����
		If (isCsMailSend) then
			Call SendUpheBoardMail(idx)
		end if
	set boardqna = Nothing

	response.write "<script type='text/javascript'>"
	response.write "	alert('����Ǿ����ϴ�.');"
	response.write "	location.replace('upche_qna_board_reply.asp?idx=" + idx + "&page=" + page + Param + "');"
	response.write "	</script>"

elseif (mode = "edit") then
	set boardqna = New CUpcheQnADetail
	boardqna.changeworker idx, workerid
	set boardqna = Nothing
	
	Response.Write "<script>alert('����Ǿ����ϴ�.');location.href='upche_qna_board_list.asp?page=" + page + Param + "';</script>"
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->