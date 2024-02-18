<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/upcheitemqna/LecDiyqnaCls.asp"-->
<%
'####################################################
' Description :  강좌&상품 Q&A 관리 처리페이지
' History : 2016.08.05 유태욱 생성
'####################################################
%>
<%
Dim idx, gridx, gubunVal, mode, menupos, returnurl, reidx
Dim sqlStr, msg, maxreplynum, minidx, ansContents, ansContentsEdit
Dim mailcontent, usermail, qstContents, lastRegdate, masterQRegName, masterQTitle
Dim pagegubun, diyitemid, lec_idx, makerid
Dim lastSMSok, lastSmsNum, newQnAmailContents
Dim loops

idx				= getNumeric(requestCheckVar(request("idx"),9))
gridx			= getNumeric(requestCheckVar(request("gridx"),9))
reidx			= getNumeric(requestCheckVar(request("reidx"),9))
gubunVal		= getNumeric(requestCheckVar(request("gubunVal"),9))
menupos			= requestCheckVar(request("menupos"),10)
usermail		= requestCheckVar(request("usermail"),255)
ansContents		= request("ansContents")
ansContentsEdit	= request("ansContentsEdit")
mode 			= requestCheckVar(request("mode"),10)
if ansContents <> "" then
	if checkNotValidHTML(ansContents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if ansContentsEdit <> "" then
	if checkNotValidHTML(ansContentsEdit) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

pagegubun 			= requestCheckVar(request("pagegubun"),10)
diyitemid 			= requestCheckVar(request("diyitemid"),10)
lec_idx 			= requestCheckVar(request("lec_idx"),10)
makerid 			= requestCheckVar(request("makerid"),10)

qstContents		= request("qstContents")
lastRegdate		= requestCheckVar(request("lastRegdate"), 30)
masterQRegName	= requestCheckVar(request("masterQRegName"),24)
masterQTitle	= requestCheckVar(request("masterQTitle"),200)

lastSMSok		= requestCheckVar(request("lastSMSok"),1)
lastSmsNum		= requestCheckVar(request("lastSmsNum"),15)
if qstContents <> "" then
	if checkNotValidHTML(qstContents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if masterQTitle <> "" then
	if checkNotValidHTML(masterQTitle) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

If mode <> "C" and mode <> "D" and mode <> "addreply" and mode <> "adel" and mode <> "edit" Then
	Call Alert_Return("잘못된 접속 입니다.")
	response.end
End If

If (checkNotValidHTML(ansContents) = true) Then
	response.write "<script>alert('답글에는 Script 또는 Action을 사용하실 수 없습니다.');history.back();</script>"
	dbACADEMYget.Close
	session.codePage = 949  
	response.End
End If

If (checkNotValidHTML(ansContentsEdit) = true) Then
	response.write "<script>alert('답글에는 Script 또는 Action을 사용하실 수 없습니다.');history.back();</script>"
	dbACADEMYget.Close
	session.codePage = 949  
	response.End
End If

Select Case mode
	Case "C"				'문의분야 변경
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_academy.[dbo].[tbl_academy_qna_NEW] SET "
		sqlStr = sqlStr & " lecture_gubun = '"&gubunVal&"' "
		sqlStr = sqlStr & " WHERE idx = '"&idx&"' "
		dbACADEMYget.execute(sqlStr)
		returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&gridx="&gridx
		msg = "수정하였습니다."
	Case "D"				'문의글 삭제
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_academy.[dbo].[tbl_academy_qna_NEW] SET "
		sqlStr = sqlStr & " isusing = 'N' "
		sqlStr = sqlStr & " Where reply_group_idx = '"&gridx&"' " 
		dbACADEMYget.execute(sqlStr)
		returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaList.asp?menupos="&menupos
		msg = "삭제하였습니다."
	Case "addreply"			'문의글 답변
		sqlStr =	" select max(reply_num) as maxreplynum, min(idx) as rowidx from [db_academy].[dbo].[tbl_academy_qna_new] where reply_group_idx="&gridx
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			maxreplynum = rsACADEMYget(0)
			minidx = rsACADEMYget(1)
		rsACADEMYget.Close
		maxreplynum = maxreplynum+1

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO [db_academy].[dbo].[tbl_academy_qna_new] " & VBCRLF
		sqlStr = sqlStr & " (reply_group_idx, reply_depth, reply_num, makerid, itemid, lec_idx, userid, username, userlevel, replyuserid, comment, qna, pagegubun, device) values " & VBCRLF
		sqlStr = sqlStr & " (" & gridx & " " & VBCRLF
		sqlStr = sqlStr & " , 1 " & VBCRLF
		sqlStr = sqlStr & " , " & maxreplynum & " " & VBCRLF
		sqlStr = sqlStr & " , '" & makerid & "' " & VBCRLF
		
		if pagegubun = "L" then
			sqlStr = sqlStr & " , null " & VBCRLF
			sqlStr = sqlStr & " , " & lec_idx & " " & VBCRLF
		elseif pagegubun = "D" then
			sqlStr = sqlStr & " , " & diyitemid & " " & VBCRLF
			sqlStr = sqlStr & " , null "  & VBCRLF
		end if

		sqlStr = sqlStr & " ,'" & session("ssBctID") & "'" & VBCRLF
		sqlStr = sqlStr & " ,'" & session("ssBctCname") & "'" & VBCRLF
		sqlStr = sqlStr & " ,'7'" & VBCRLF											'SCM이라 레벨 7로..
		sqlStr = sqlStr & " ,'" & session("ssBctID") & "'" & VBCRLF
		sqlStr = sqlStr & " ,'" & html2db(ansContents) & "'" & VBCRLF
		sqlStr = sqlStr & " ,'A'" & VBCRLF
		sqlStr = sqlStr & " , '" & pagegubun & "' " & VBCRLF
		sqlStr = sqlStr & " ,'M');" & VBCRLF										'어드민이라 WEB인 W로
		sqlStr = sqlStr & " Update [db_academy].[dbo].[tbl_academy_qna_new] Set " & VBCRLF
		sqlStr = sqlStr & " 	  answerYN = 'Y'" & VBCRLF
		sqlStr = sqlStr & " Where reply_group_idx=" & gridx & " "
'response.write sqlStr
'response.end
		dbACADEMYget.execute(sqlStr)
		returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&gridx="&gridx
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
			newQnAmailContents = getQnaContents(gridx)
			mailcontent = Replace(mailcontent,"#qnaTBLS#",newQnAmailContents)    
			
			Call sendmail("customer@thefingers.co.kr", usermail, "[더 핑거스] 문의하신 내용에 대한 답변이 등록되었습니다.", mailcontent)
		End If

		'SMS체크했다면 SMS전송
		If (Len(lastsmsok) > 0 AND lastsmsok = "Y") AND (Len(lastSmsNum) > 0) Then
			sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"&lastSmsNum&"','1644-1557','[더 핑거스] 문의하신 내용에 대한 답변이 등록되었습니다.'"
			dbget.Execute sqlStr
		End If

	Case "edit"				'답변글 수정
		sqlStr = ""
		sqlStr = sqlStr & " Update [db_academy].[dbo].[tbl_academy_qna_new] Set " & VBCRLF
		sqlStr = sqlStr & " comment = '" & html2db(ansContentsEdit) & "'" & VBCRLF
		sqlStr = sqlStr & " Where idx=" & reidx  & VBCRLF
		dbACADEMYget.execute(sqlStr)
		returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&gridx="&gridx
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
			newQnAmailContents = getQnaContents(gridx)
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
		sqlStr = sqlStr & " UPDATE [db_academy].[dbo].[tbl_academy_qna_new] SET isusing = 'N' WHERE idx = '" & reidx & "'; "
		sqlStr = sqlStr & " UPDATE [db_academy].[dbo].[tbl_academy_qna_new] SET answerYN = 'Y' WHERE reply_group_idx = '" & gridx & "' "
		dbACADEMYget.execute(sqlStr)
		returnurl = "/cscenterv2/upcheitemqna/Qna/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&gridx="&gridx
		msg = "삭제하였습니다."
End Select
response.write "<script>alert('"&msg&"');location.replace('"&returnurl&"');</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->