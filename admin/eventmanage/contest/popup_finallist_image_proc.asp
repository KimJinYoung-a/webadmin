<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->

<%
	Dim vContest, vPoll_Idx, vUserID, vSubject, vContents, vGubun, i, vQuery, vImgName, vImgName2, vImgCode, vIdx, vSortNo
	vGubun		= Request("gubun")
	vIdx		= Request("idx")
	vContest	= Request("contest")
	vPoll_Idx	= Request("poll_idx")
	vImgName	= Request("image_url")
	vImgName2	= Request("image_url2")
	vImgCode	= Request("code")
	vSortNo		= Request("sortno")
	
	If vGubun = "" Then
		If vIdx = "" Then
			vQuery = "INSERT INTO [db_event].[dbo].[tbl_contest_poll_image](poll_idx, contest, img_code, img_name, img_name2, sortno)" & _
					 "	VALUES('" & vPoll_Idx & "', '" & vContest & "', '" & vImgCode & "', '" & vImgName & "', '" & vImgName2 & "', '" & vSortNo & "') "
			dbget.execute vQuery
		Else
			vQuery = "UPDATE [db_event].[dbo].[tbl_contest_poll_image] SET " & _
					 "		img_code = '" & vImgCode & "', " & _
					 "		img_name = '" & vImgName & "', " & _
					 "		img_name2 = '" & vImgName2 & "', " & _
					 "		sortno = '" & vSortNo & "' " & _
					 "	WHERE idx = '" & vIdx & "' AND poll_idx = '" & vPoll_Idx & "' AND contest = '" & vContest & "' "
			dbget.execute vQuery
		End If
	ElseIf vGubun = "del" Then
		vQuery = "DELETE [db_event].[dbo].[tbl_contest_poll_image] WHERE idx = '" & vIdx & "' AND poll_idx = '" & vPoll_Idx & "' AND contest = '" & vContest & "'"
		dbget.execute vQuery
		
		vIdx = ""
	End If
	
	Response.Write "<script language='javascript'>alert('처리되었습니다.');location.href='popup_finallist_image.asp?contest="&vContest&"&usernum="&vPoll_Idx&"&idx="&vIdx&"';</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->