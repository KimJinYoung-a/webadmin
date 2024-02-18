<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->

<%
	Dim vContest, vUserIdx, vUserID, vSubject, vContents, vGubun, i, vQuery
	vGubun		= Request("gubun")
	vContest	= Request("divnum")
	vUserIdx	= Request("idx")
	vUserID		= Request("userid")
	vSubject	= html2db(requestCheckVar(Request("subject"),50))
	vContents	= html2db(Request("contents"))

	If vGubun = "insert" Then
		If vUserIdx <> "" Then
			vQuery = "UPDATE [db_event].[dbo].[tbl_contest_poll] SET userid = '" & vUserID & "', subject = '" & vSubject & "', contents = '" & vContents & "' WHERE contest = '" & vContest & "' AND idx = '" & vUserIdx & "'"
			dbget.execute vQuery
		ElseIf vUserIdx = "" Then
			vQuery = "SELECT COUNT(userid) FROM [db_user].[dbo].[tbl_user_n] WHERE userid = '" & vUserID & "'"
			rsget.Open vQuery,dbget,1
			If rsget(0) < 1 Then
				Response.Write "<script>alert('" & vUserID & " -> 없는 아이디입니다.');history.back();</script>"
				dbget.close()
				Response.End
			End If
			rsget.close()
			
			vQuery = "SELECT COUNT(idx) FROM [db_event].[dbo].[tbl_contest_poll] WHERE contest = '" & vContest & "' AND userid = '" & vUserID & "'"
			rsget.Open vQuery,dbget,1
			If rsget(0) > 0 Then
				Response.Write "<script>alert('" & vUserID & " -> 이미 저장된 아이디입니다.');history.back();</script>"
				dbget.close()
				Response.End
			End If
			rsget.close()
			
			vQuery = "INSERT INTO [db_event].[dbo].[tbl_contest_poll](contest,userid,subject,contents) VALUES('" & vContest & "','" & vUserID & "','" & vSubject & "','" & vContents & "')"
			dbget.execute vQuery
		End If
	ElseIf vGubun = "del" Then
		vQuery = "DELETE [db_event].[dbo].[tbl_contest_poll] WHERE contest = '" & vContest & "' AND idx = '" & vUserIdx & "'"
		dbget.execute vQuery
	ElseIf vGubun = "pollplus" Then
		vQuery = "UPDATE [db_event].[dbo].[tbl_contest_poll] SET poll_count = poll_count + 1 WHERE contest = '" & vContest & "' AND idx = '" & vUserIdx & "'"
		dbget.execute vQuery
	End If
	
	Response.Write "<script>alert('저장되었습니다.');location.href='popup_finallist.asp?divnum=" & vContest & "';</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->