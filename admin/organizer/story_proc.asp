<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%

'// 오거나이저 저장폼

dim idx, title, contents, isusing
idx = request.Form("idx")
title = request.Form("title")
contents= request.Form("contents")
isusing = request.Form("isusing")

dim strSQL,msg
	 
IF idx="" Then
	strSQL =" INSERT INTO db_diary2010.[dbo].tbl_organizer_story "&_
			" (TITLE, CONTENTS, ISUSING) "&_
			" VALUES(" &_
			"'" & html2db(title) & "' " &_
			",'" & html2db(contents) & "' " &_
			",'" & isusing & "') " 

	'response.write strSQL&"<br>"
	dbget.execute(strSQL)

ELSE
	strSQL =" UPDATE db_diary2010.[dbo].tbl_organizer_story "&_
			" SET TITLE = '" & html2db(title) & "', " &_
			"	CONTENTS = '" & html2db(contents) & "', " &_
			"	ISUSING = '" & isusing & "' "&_
			" WHERE IDX = '" & idx & "' "

	'response.write strSQL&"<br>"
	dbget.execute(strSQL)

End IF


%>

<script>alert('저장되었습니다.');location.href='story.asp?menupos=1136';</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->