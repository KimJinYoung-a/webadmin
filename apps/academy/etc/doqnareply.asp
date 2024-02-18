<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/upcheitemqna/LecDiyqnaCls.asp"-->
<%
Dim idx, ansContents, ridx, MakerID, mode, pagegubun, lastRegdate
Dim lec_idx, diyitemid, masterQTitle, newQnAmailContents, qnacount
Dim lastSMSok, lastSmsNum, MakerName, usermail
mode			= requestCheckVar(request("mode"),10)
idx				= requestCheckVar(request("idx"),10)
ansContents		= html2db(request("ansContents"))
ridx			= requestCheckVar(request("ridx"),10)
MakerID			= requestCheckVar(request.cookies("partner")("userid"),32)
MakerName		= requestCheckVar(request.cookies("partner")("ssBctCname"),32)
pagegubun 		= requestCheckVar(request("pagegubun"),10)
lec_idx 		= requestCheckVar(request("lec_idx"),10)
diyitemid		= requestCheckVar(request("diyitemid"),10)
masterQTitle	= requestCheckVar(request("masterQTitle"),200)
lastSMSok		= requestCheckVar(request("lastSMSok"),1)
lastSmsNum		= requestCheckVar(request("lastSmsNum"),15)
usermail		= requestCheckVar(request("usermail"),255)
lastRegdate		= requestCheckVar(request("lastRegdate"), 30)

if (checkNotValidHTML(ansContents) = true) Then
	response.write "<script>alert('Script 또는 Action이나 HTML을 사용하실 수 없습니다.');history.back();</script>"
	response.End
End If

if (checkNotValidHTML(masterQTitle) = true) Then
	response.write "<script>alert('Script 또는 Action이나 HTML을 사용하실 수 없습니다.');history.back();</script>"
	response.End
End If

if (checkNotValidHTML(usermail) = true) Then
	response.write "<script>alert('Script 또는 Action이나 HTML을 사용하실 수 없습니다.');history.back();</script>"
	response.End
End If

Dim sqlStr, msg, mailcontent
Dim maxreplynum, minidx
qnacount=0
IF ridx="" Or mode="" Or MakerID="" THEN
	msg="잘못된 접속 입니다."
Else
	Select Case mode
		Case "C"				'문의분야 변경
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE db_academy.[dbo].[tbl_academy_qna_NEW] SET "
			sqlStr = sqlStr & " lecture_gubun = '"&gubunVal&"' "
			sqlStr = sqlStr & " WHERE idx = '"&idx&"' "
			dbACADEMYget.execute(sqlStr)
			'returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&ridx="&ridx
			msg = "수정하였습니다."
		Case "D"				'문의글 삭제
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE db_academy.[dbo].[tbl_academy_qna_NEW] SET "
			sqlStr = sqlStr & " isusing = 'N' "
			sqlStr = sqlStr & " Where reply_group_idx = '"&ridx&"' " 
			dbACADEMYget.execute(sqlStr)
			'returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaList.asp?menupos="&menupos
			msg = "삭제하였습니다."
		Case "addreply"			'문의글 답변
			sqlStr =	" select max(reply_num) as maxreplynum, min(idx) as rowidx from [db_academy].[dbo].[tbl_academy_qna_new] where reply_group_idx="&ridx
			rsACADEMYget.Open sqlStr,dbACADEMYget,1
				maxreplynum = rsACADEMYget(0)
				minidx = rsACADEMYget(1)
			rsACADEMYget.Close
			maxreplynum = maxreplynum+1

			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO [db_academy].[dbo].[tbl_academy_qna_new] " & VBCRLF
			sqlStr = sqlStr & " (reply_group_idx, reply_depth, reply_num, makerid, itemid, lec_idx, userid, username, userlevel, replyuserid, comment, qna, pagegubun, device) values " & VBCRLF
			sqlStr = sqlStr & " (" & ridx & " " & VBCRLF
			sqlStr = sqlStr & " , 1 " & VBCRLF
			sqlStr = sqlStr & " , " & maxreplynum & " " & VBCRLF
			sqlStr = sqlStr & " , '" & MakerID & "' " & VBCRLF
			
			if pagegubun = "L" then
				sqlStr = sqlStr & " , null " & VBCRLF
				sqlStr = sqlStr & " , " & lec_idx & " " & VBCRLF
			elseif pagegubun = "D" then
				sqlStr = sqlStr & " , " & diyitemid & " " & VBCRLF
				sqlStr = sqlStr & " , null "  & VBCRLF
			end if

			sqlStr = sqlStr & " ,'" & MakerID & "'" & VBCRLF
			sqlStr = sqlStr & " ,'" & MakerName & "'" & VBCRLF
			sqlStr = sqlStr & " ,'7'" & VBCRLF											'SCM이라 레벨 7로..
			sqlStr = sqlStr & " ,'" & MakerID & "'" & VBCRLF
			sqlStr = sqlStr & " ,'" & html2db(ansContents) & "'" & VBCRLF
			sqlStr = sqlStr & " ,'A'" & VBCRLF
			sqlStr = sqlStr & " , '" & pagegubun & "' " & VBCRLF
			sqlStr = sqlStr & " ,'M');" & VBCRLF										'어드민이라 WEB인 W로
			sqlStr = sqlStr & " Update [db_academy].[dbo].[tbl_academy_qna_new] Set " & VBCRLF
			sqlStr = sqlStr & " 	  answerYN = 'Y'" & VBCRLF
			sqlStr = sqlStr & " Where reply_group_idx=" & ridx & " "
	'response.write FormatDate(lastRegdate,"0000.00.00")
	'response.end
			dbACADEMYget.execute(sqlStr)
			'returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&ridx="&ridx
			msg = "저장하였습니다."

			'답변 메일 발송
			If Cstr(usermail) <> "" Then
				'메일 템플릿 접수
	'			mailcontent = ReadLocalFile("tpl_fingers_qna.html", "/academy/lib/mail_templete")
	'			mailcontent = Replace(mailcontent,"#ansContents#",nl2br(ansContents))
	'			mailcontent = Replace(mailcontent,"#qstContents#",nl2br(qstContents))
	'			mailcontent = Replace(mailcontent,"#regdate#",FormatDate(lastRegdate,"0000-00-00"))
	'			mailcontent = Replace(mailcontent,"#qstUserName#",masterQRegName)
	'			mailcontent = Replace(mailcontent,"#qstTitle#",masterQTitle)       
				
				mailcontent = ReadLocalFile("mail_counsel_reply.html", "/academy/lib/mail_templete")
				mailcontent = Replace(mailcontent,"#regdate#",FormatDate(lastRegdate,"0000.00.00"))
				mailcontent = Replace(mailcontent,"#qstTitle#",masterQTitle)
				newQnAmailContents = getQnaContents(ridx)
				mailcontent = Replace(mailcontent,"#qnaTBLS#",newQnAmailContents)
				
				Call sendmail("customer@thefingers.co.kr", usermail, "[더 핑거스] 문의하신 내용에 대한 답변이 등록되었습니다.", mailcontent)
			End If

			'SMS체크했다면 SMS전송
			If (Len(lastsmsok) > 0 AND lastsmsok = "Y") AND (Len(lastSmsNum) > 0) Then
				sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"&lastSmsNum&"','1644-1557','[더 핑거스] 문의하신 내용에 대한 답변이 등록되었습니다.'"
				dbget.Execute sqlStr
			End If
			'뱃지 카운트 확인
			sqlStr = "exec [db_academy].[dbo].[sp_Academy_App_IconBadgeCountQnASet] '" + Cstr(MakerID) + "'"
			rsACADEMYget.CursorLocation = adUseClient
			rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
			if not rsACADEMYget.EOF Then
				qnacount=rsACADEMYget("qnacnt")
			end if
			rsACADEMYget.Close

		Case "edit"				'답변글 수정
			sqlStr = ""
			sqlStr = sqlStr & " Update [db_academy].[dbo].[tbl_academy_qna_new] Set " & VBCRLF
			sqlStr = sqlStr & " comment = '" & html2db(ansContents) & "'" & VBCRLF
			sqlStr = sqlStr & " Where idx=" & idx  & VBCRLF
			dbACADEMYget.execute(sqlStr)
			'returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&ridx="&ridx
			msg = "수정하였습니다."

			'답변 메일 발송
			If Cstr(usermail) <> "" Then
				'메일 템플릿 접수
	'			mailcontent = ReadLocalFile("tpl_fingers_qna.html", "/academy/lib/mail_templete")
	'			mailcontent = Replace(mailcontent,"#ansContents#",nl2br(ansContentsEdit))
	'			mailcontent = Replace(mailcontent,"#qstContents#",nl2br(qstContents))
	'			mailcontent = Replace(mailcontent,"#regdate#",FormatDate(lastRegdate,"0000-00-00"))
	'			mailcontent = Replace(mailcontent,"#qstUserName#",masterQRegName)
	'			mailcontent = Replace(mailcontent,"#qstTitle#",masterQTitle)       
				
				mailcontent = ReadLocalFile("mail_counsel_reply.html", "/academy/lib/mail_templete")
				mailcontent = Replace(mailcontent,"#regdate#",FormatDate(lastRegdate,"0000.00.00"))
				mailcontent = Replace(mailcontent,"#qstTitle#",masterQTitle)
				newQnAmailContents = getQnaContents(ridx)
				mailcontent = Replace(mailcontent,"#qnaTBLS#",newQnAmailContents)
					
				Call sendmail("customer@thefingers.co.kr", usermail, "[더 핑거스] 문의하신 내용에 대한 답변이 등록되었습니다.", mailcontent)
			End If

			'SMS체크했다면 SMS전송 // 답글 수정시에는 SMS 보내지 않음.
			''If (Len(lastsmsok) > 0 AND lastsmsok = "Y") AND (Len(lastSmsNum) > 0) Then
			''	sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"&lastSmsNum&"','1644-1557','[더 핑거스] 문의하신 내용에 대한 답변이 등록되었습니다.'"
			''	dbget.Execute sqlStr
			''End If

		Case "adel"				'답변글 삭제
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE [db_academy].[dbo].[tbl_academy_qna_new] SET isusing = 'N' WHERE idx = '" & idx & "'; "
			sqlStr = sqlStr & " UPDATE [db_academy].[dbo].[tbl_academy_qna_new] SET answerYN = 'Y' WHERE reply_group_idx = '" & ridx & "' "
			dbACADEMYget.execute(sqlStr)
			'returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&ridx="&ridx
			msg = "삭제하였습니다."
	End Select
END IF

%>
<script>
<!--
parent.fnQnARelyEnd("<%=msg%>","<%=mode%>",<%=qnacount%>);
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->