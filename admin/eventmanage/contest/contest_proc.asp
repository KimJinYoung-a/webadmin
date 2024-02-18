<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 공모전리스트
' History : 이상구 생성
'			한용민 수정(isms취약점조치)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->

<%
	Dim vQuery, vSubject, vContest, vEntrySDate, vEntryEDate, vVoteSDate, vVoteEDate, vResultDate, vUseYN, vRegdate
	
	vContest 		= requestCheckVar(Request("contest"),6)
	vSubject		= html2db(requestCheckVar(trim(Request("subject")),100))
	vEntrySDate		= Request("entry_sdate")
	vEntryEDate		= Request("entry_edate")
	vVoteSDate		= Request("vote_sdate")
	vVoteEDate		= Request("vote_edate")
	vResultDate		= Request("result_date")
	vUseYN			= Request("useyn")

	if vSubject <> "" and not(isnull(vSubject)) then
		vSubject = ReplaceBracket(vSubject)

		'if checkNotValidHTML(vSubject) then
		'	response.write "<script type='text/javascript'>"
		'	response.write "	alert('주제에는 HTML을 사용하실 수 없습니다.');history.back();"
		'	response.write "</script>"
		'	response.End
		'end if
	end If

	If vContest <> "" Then
		vQuery = "UPDATE [db_event].[dbo].[tbl_contest_master] " & _
				 "		SET " & _
				 "			subject = '" & vSubject & "', " & _
				 "			entry_sdate = '" & vEntrySDate & "', " & _
				 "			entry_edate = '" & vEntryEDate & "', " & _
				 "			vote_sdate = '" & vVoteSDate & "', " & _
				 "			vote_edate = '" & vVoteEDate & "', " & _
				 "			result_date = '" & vResultDate & "', " & _
				 "			useyn = '" & vUseYN & "' " & _
				 "	WHERE contest = '" & vContest & "' "
		dbget.execute vQuery
	Else
		vQuery = "SELECT TOP 1 contest FROM [db_event].[dbo].[tbl_contest_master] ORDER BY Cast(replace(contest,'con','') as int) DESC"
		rsget.Open vQuery,dbget,1
		If Not rsget.Eof Then
			vContest = CInt(Replace(rsget(0),"con","")) + 1
			vContest = "con" & Num2Str(vContest,3,"0","R")
		Else
			vContest = "con01"
		End If
		rsget.close()
			
		vQuery = "INSERT INTO [db_event].[dbo].[tbl_contest_master](contest, subject, entry_sdate, entry_edate, vote_sdate, vote_edate, result_date, useyn) " & _
				 "	VALUES('" & vContest & "','" & vSubject & "','" & vEntrySDate & "','" & vEntryEDate & "', '" & vVoteSDate & "','" & vVoteEDate & "','" & vResultDate & "','" & vUseYN & "')"
		dbget.execute vQuery
	End If
	
	Response.Write "<script>alert('저장되었습니다.');opener.location.reload();window.close();</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->