<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/myqnaCls.asp"-->
<%
'####################################################
' Description :  1:1 ������ ���μ���
' History : 2016.07.28 ������ ����
'####################################################
%>
<%
Dim idx, gridx, gubunVal, mode, menupos, returnurl, reidx
Dim sqlStr, msg, maxreplynum, minidx, ansContents, ansContentsEdit
Dim mailcontent, usermail, qstContents, lastRegdate, masterQRegName, masterQTitle
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

qstContents		= request("qstContents")
lastRegdate		= requestCheckVar(request("lastRegdate"), 30)
masterQRegName	= requestCheckVar(request("masterQRegName"),24)
masterQTitle	= requestCheckVar(request("masterQTitle"),200)

lastSMSok		= requestCheckVar(request("lastSMSok"),1)
lastSmsNum		= requestCheckVar(request("lastSmsNum"),15)

If mode <> "C" and mode <> "D" and mode <> "addreply" and mode <> "adel" and mode <> "edit" Then
	Call Alert_Return("�߸��� ���� �Դϴ�.")
	response.end
End If

If (checkNotValidHTML(ansContents) = true) Then
	response.write "<script>alert('��ۿ��� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	dbACADEMYget.Close
	session.codePage = 949  
	response.End
End If

If (checkNotValidHTML(ansContentsEdit) = true) Then
	response.write "<script>alert('��ۿ��� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	dbACADEMYget.Close
	session.codePage = 949  
	response.End
End If

Select Case mode
	Case "C"				'���Ǻо� ����
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_academy.[dbo].[tbl_academy_qna_NEW] SET "
		sqlStr = sqlStr & " lecture_gubun = '"&gubunVal&"' "
		sqlStr = sqlStr & " WHERE idx = '"&idx&"' "
		dbACADEMYget.execute(sqlStr)
		returnurl = "/academy/board/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&gridx="&gridx
		msg = "�����Ͽ����ϴ�."
	Case "D"				'���Ǳ� ����
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_academy.[dbo].[tbl_academy_qna_NEW] SET "
		sqlStr = sqlStr & " isusing = 'N' "
		sqlStr = sqlStr & " Where reply_group_idx = '"&gridx&"' " 
		dbACADEMYget.execute(sqlStr)
		returnurl = "/academy/board/myqnaList.asp?menupos="&menupos
		msg = "�����Ͽ����ϴ�."
	Case "addreply"			'���Ǳ� �亯
		sqlStr =	" select max(reply_num) as maxreplynum, min(idx) as rowidx from [db_academy].[dbo].[tbl_academy_qna_new] where reply_group_idx="&gridx
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			maxreplynum = rsACADEMYget(0)
			minidx = rsACADEMYget(1)
		rsACADEMYget.Close
		maxreplynum = maxreplynum+1

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO [db_academy].[dbo].[tbl_academy_qna_new] " & VBCRLF
		sqlStr = sqlStr & " (reply_group_idx, reply_depth, reply_num, userid, username, userlevel, replyuserid, comment, qna, pagegubun, device) values " & VBCRLF
		sqlStr = sqlStr & " (" & gridx & " " & VBCRLF
		sqlStr = sqlStr & " , 1 " & VBCRLF
		sqlStr = sqlStr & " , " & maxreplynum & " " & VBCRLF
		sqlStr = sqlStr & " ,'" & session("ssBctID") & "'" & VBCRLF
		sqlStr = sqlStr & " ,'" & session("ssBctCname") & "'" & VBCRLF
		sqlStr = sqlStr & " ,'7'" & VBCRLF											'SCM�̶� ���� 7��..
		sqlStr = sqlStr & " ,'" & session("ssBctID") & "'" & VBCRLF
		sqlStr = sqlStr & " ,'" & html2db(ansContents) & "'" & VBCRLF
		sqlStr = sqlStr & " ,'A'" & VBCRLF
		sqlStr = sqlStr & " ,'F'" & VBCRLF
		sqlStr = sqlStr & " ,'W');" & VBCRLF										'�����̶� WEB�� W��
		sqlStr = sqlStr & " Update [db_academy].[dbo].[tbl_academy_qna_new] Set " & VBCRLF
		sqlStr = sqlStr & " 	  answerYN = 'Y'" & VBCRLF
		sqlStr = sqlStr & " Where reply_group_idx=" & gridx & " "
		dbACADEMYget.execute(sqlStr)
		returnurl = "/academy/board/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&gridx="&gridx
		msg = "�����Ͽ����ϴ�."

		'�亯 ���� �߼�
		If Cstr(usermail) <> "" Then
			'���� ���ø� ����
			mailcontent = ReadLocalFile("mail_counsel_reply.html", "/academy/lib/mail_templete")
			mailcontent = Replace(mailcontent,"#regdate#",FormatDate(lastRegdate,"0000.00.00"))
			mailcontent = Replace(mailcontent,"#qstTitle#",masterQTitle)
			newQnAmailContents = getQnaContents(gridx)
			mailcontent = Replace(mailcontent,"#qnaTBLS#",newQnAmailContents)
			Call sendmail("customer@thefingers.co.kr", usermail, "[���ΰŽ�] 1:1 ���ǰԽ��ǿ� �亯�� ��ϵǾ����ϴ�.", mailcontent)
		End If

		'SMSüũ�ߴٸ� SMS����
		If (Len(lastsmsok) > 0 AND lastsmsok = "Y") AND (Len(lastSmsNum) > 0) Then
			sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"&lastSmsNum&"','1644-1557','[���ΰŽ�] 1:1 ���ǰԽ��ǿ� �亯�� ��ϵǾ����ϴ�...'"
			dbget.Execute sqlStr
		End If

	Case "edit"				'�亯�� ����
		sqlStr = ""
		sqlStr = sqlStr & " Update [db_academy].[dbo].[tbl_academy_qna_new] Set " & VBCRLF
		sqlStr = sqlStr & " comment = '" & html2db(ansContentsEdit) & "'" & VBCRLF
		sqlStr = sqlStr & " Where idx=" & reidx  & VBCRLF
		dbACADEMYget.execute(sqlStr)
		returnurl = "/academy/board/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&gridx="&gridx
		msg = "�����Ͽ����ϴ�."

		'�亯 ���� �߼�
		If Cstr(usermail) <> "" Then
			'���� ���ø� ����
			mailcontent = ReadLocalFile("mail_counsel_reply.html", "/academy/lib/mail_templete")
			mailcontent = Replace(mailcontent,"#regdate#",FormatDate(lastRegdate,"0000.00.00"))
			mailcontent = Replace(mailcontent,"#qstTitle#",masterQTitle)
			newQnAmailContents = getQnaContents(gridx)
			mailcontent = Replace(mailcontent,"#qnaTBLS#",newQnAmailContents)
			Call sendmail("customer@thefingers.co.kr", usermail, "[���ΰŽ�] 1:1 ���ǰԽ��ǿ� �亯�� ��ϵǾ����ϴ�.", mailcontent)
		End If

		'SMSüũ�ߴٸ� SMS���� // ��� �����ÿ��� SMS ������ ����.
		'If (Len(lastsmsok) > 0 AND lastsmsok = "Y") AND (Len(lastSmsNum) > 0) Then
		'	sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"&lastSmsNum&"','1644-1557','[���ΰŽ�] 1:1 ���ǰԽ��ǿ� �亯�� ��ϵǾ����ϴ�.'"
		'	dbget.Execute sqlStr
		'End If

	Case "adel"				'�亯�� ����
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE [db_academy].[dbo].[tbl_academy_qna_new] SET isusing = 'N' WHERE idx = '" & reidx & "'; "
		sqlStr = sqlStr & " UPDATE [db_academy].[dbo].[tbl_academy_qna_new] SET answerYN = 'N' WHERE reply_group_idx = '" & gridx & "' "
		dbACADEMYget.execute(sqlStr)
		returnurl = "/academy/board/myqnaView.asp?menupos="&menupos&"&idx="&idx&"&gridx="&gridx
		msg = "�����Ͽ����ϴ�."
End Select
response.write "<script>alert('"&msg&"');location.replace('"&returnurl&"');</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->