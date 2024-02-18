<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% 
	response.Charset="UTF-8"
	Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 응원톡 쓰기"
'####################################################
' Description : 응원톡 수정/답글
' History : 2017-01-11 이종화 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/apps/academy/etc/talk_cls.asp" -->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<%
Dim mode, gubun, ridx, idx, paramid, vUserid, commContents, returnurl, SuccessUrl
Dim sqlStr, maxidx, maxreplynum, depth
mode 			= request.Form("mode")
commContents	= request.Form("commContents")
gubun			= request.Form("gubun")
paramid			= request.Form("paramid")
ridx 			= request.Form("ridx")
idx 			= request.Form("idx")
commContents	= request.Form("commContents")
depth			= request.Form("depth")
vUserid			= requestCheckVar(request.cookies("partner")("userid"),32)

returnurl = "/apps/academy/etc/talk_write.asp"
SuccessUrl = "/apps/academy/etc/talk.asp"

If mode <> "del" and mode <> "add" and mode <> "reply" and mode <> "addreply" and mode <> "edit" then
	Call Alert_Return("잘못된 접속 입니다.1")
	response.end
	response.redirect returnurl
End If

'내용 검사
If checkNotValidHTML(commContents) Then
	Call Alert_Return("HTML은 사용하실 수 없습니다.")
	response.end
	response.redirect returnurl
End If
commContents = ReplaceBracket(commContents)

If mode <> "del" then
	If paramid = "" Then
		Call Alert_Return("잘못된 접속 입니다.1")
		response.end
		response.redirect returnurl
	End If
	
	If mode = "reply" and ridx = "" Then
		Call Alert_Return("잘못된 접속 입니다.2")
		response.end
		response.redirect returnurl
	End If
End If

sqlStr = ""
sqlStr = sqlStr & " SELECT MAX(idx) as maxidx FROM [db_academy].[dbo].[tbl_academy_cheertalk_comments] "
rsACADEMYget.Open sqlStr,dbACADEMYget,1
	maxidx = rsACADEMYget(0)
rsACADEMYget.Close

If isnull(maxidx) Then
	maxidx = 1
Else
	maxidx = maxidx + 1
End If 

'// Mode에 따른 실행 쿼리 선택 //
Select Case mode
	Case "add"
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO [db_academy].[dbo].[tbl_academy_cheertalk_comments] " & vbcrlf
		sqlStr = sqlStr & "	(gubun, paramid, reply_group_idx, reply_depth, reply_num, userid, comment, device, isusing, regdate) VALUES " & vbcrlf
		sqlStr = sqlStr & "	('" & gubun & "'" & vbcrlf
		sqlStr = sqlStr & "	,'" & paramid & "'" & vbcrlf
		sqlStr = sqlStr & "	,'" & maxidx & "'" & vbcrlf
		sqlStr = sqlStr & "	, 0 " & vbcrlf
		sqlStr = sqlStr & "	, 0 " & vbcrlf
		sqlStr = sqlStr & "	,'" & vUserid & "'" & vbcrlf
		sqlStr = sqlStr & "	,'" & html2db(commContents) & "'" & vbcrlf
		sqlStr = sqlStr & "	,'M' " & vbcrlf
		sqlStr = sqlStr & "	,'Y' " & vbcrlf
		sqlStr = sqlStr & "	,getdate()) "
	Case "reply"
		sqlStr = ""
		sqlStr = sqlStr & " SELECT MAX(reply_num) as maxreplynum FROM [db_academy].[dbo].[tbl_academy_cheertalk_comments] "
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			maxreplynum = rsACADEMYget(0)
		rsACADEMYget.Close
		maxreplynum = maxreplynum + 1

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO [db_academy].[dbo].[tbl_academy_cheertalk_comments] " & vbcrlf
		sqlStr = sqlStr & "	(gubun, paramid, reply_group_idx, reply_depth, reply_num, userid, comment, device, isusing, regdate) VALUES " & vbcrlf
		sqlStr = sqlStr & "	('" & gubun & "'" & vbcrlf
		sqlStr = sqlStr & "	,'" & paramid & "'" & vbcrlf
		sqlStr = sqlStr & "	," & ridx & vbcrlf
		sqlStr = sqlStr & "	, 1 " & vbcrlf
		sqlStr = sqlStr & "	, " & maxreplynum & vbcrlf
		sqlStr = sqlStr & "	,'" & vUserid & "'" & vbcrlf
		sqlStr = sqlStr & "	,'" & html2db(commContents) & "'" & vbcrlf
		sqlStr = sqlStr & "	,'M' " & vbcrlf
		sqlStr = sqlStr & "	,'Y' " & vbcrlf
		sqlStr = sqlStr & "	,getdate()) "
	Case "edit"
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE [db_academy].[dbo].[tbl_academy_cheertalk_comments] SET " & vbcrlf
		sqlStr = sqlStr & "	comment = '" & html2db(commContents) & "'" & vbcrlf
		sqlStr = sqlStr & " WHERE idx='" & idx & "'"  & vbcrlf
		sqlStr = sqlStr & " and userid = '"&vUserid&"' "& vbcrlf
		sqlStr = sqlStr & " and paramid = '"&paramid&"' "
	Case "del"
		Dim sqlsearch

		If depth = "0" Then
			sqlsearch = sqlsearch & " and reply_group_idx= '"&ridx&"'"
		Else
			sqlsearch = sqlsearch & " and idx= '"&idx&"'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " UPDATE [db_academy].[dbo].[tbl_academy_cheertalk_comments] SET " & vbcrlf
		sqlStr = sqlStr & " isusing = 'N' "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and userid = '"&vUserid&"' "& vbcrlf
		sqlStr = sqlStr & " and paramid = '"&paramid&"' "
		sqlStr = sqlStr & sqlsearch
End Select

'// DB실행 및 페이지 이동 //

	'트랜젝션 시작
	dbACADEMYget.beginTrans

	'실행
	dbACADEMYget.execute(sqlStr)

	'오류검사 및 반영
    If Err.Number = 0 Then   
    	dbACADEMYget.CommitTrans				'커밋(정상)
    	'response.redirect SuccessUrl
		If mode = "del" Then
			Response.write "<script>parent.fnTalkListRelold();</script>"
		Else
			Response.write "<script>fnAPPopenerJsCallClose('fnTalkListRelold(\'\')');</script>"
		End if
    Else
        dbACADEMYget.RollBackTrans				'롤백(에러발생시)
        Call Alert_Return("처리중 에러가 발생했습니다.")
    End If
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->