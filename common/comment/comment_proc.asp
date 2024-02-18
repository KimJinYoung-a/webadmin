<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/common/comment/commentCls.asp"-->

<%
	Dim strSql, vFileTemp, i, iCurrentpage
	Dim vCols, vRows, vBtnWidth, vBtnHeight, vParentIdx, vCommentIdx, vComment, vSMS, vRegistId, vBoardType, vBoardGubun, vEtc1, vEtc2
	
	vCols			= NullFillWith(requestCheckVar(Request("cols"),3),97)
	vRows			= NullFillWith(requestCheckVar(Request("rows"),3),3)
	vBtnWidth		= NullFillWith(requestCheckVar(Request("btnwidth"),3),80)
	vBtnHeight		= NullFillWith(requestCheckVar(Request("btnheight"),3),50)
	vParentIdx		= requestCheckVar(Request("pidx"),10)
	vCommentIdx		= requestCheckVar(Request("cidx"),10)
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	vComment		= html2db(Request("comment"))
	vSMS			= NullFillWith(Request("sms_send"),"")
	vRegistId		= NullFillWith(Request("registid"),"")
	vBoardType		= requestCheckVar(Request("boardtype"),2)
	vBoardGubun		= requestCheckVar(Request("boardgubun"),50)
	vEtc1			= requestCheckVar(Request("etc1"),100)
	vEtc2			= requestCheckVar(Request("etc2"),100)
	
	
	If vCommentIdx = "" Then
	
		'####### 답변 저장 #######
		strSql = " INSERT INTO [db_board].[dbo].[tbl_scm_comment] " & _
				 "		(boardGubun, parentIdx, comment, etc1, etc2, regUserid, deleteyn) " & _
				 "	VALUES " & _
				 "		('" & vBoardGubun & "', '" & vParentIdx & "', '" & vComment & "', " & CHKIIF(vEtc1="","null","'"&vEtc1&"'") & ", " & CHKIIF(vEtc2="","null","'"&vEtc2&"'") & ", '" & session("ssBctId") & "', 'n') "
		dbget.execute strSql
		
		
		'####### SMS 전송 ######
		If vSMS = "o" Then
			Dim vBoardName
			vBoardName = fnBoardName(vBoardType, vBoardGubun)
			Call SendNormalSMS(fnGetMemberHp(vRegistId),"",""&session("ssBctCname")&"님께서 " & vBoardName & " 답변을 남기셨습니다.(No." & vParentIdx & ")")
		End If
		
	Else
	
		If requestCheckVar(Request("del"),1) = "o" Then
			
			'####### 답변 삭제 #######
			strSql = " UPDATE [db_board].[dbo].[tbl_scm_comment] SET " & _
					 "		deleteyn = 'y' " & _
					 "	WHERE " & _
					 "		cIdx = '" & vCommentIdx & "' "
			dbget.execute strSql

		Else
		
			'####### 답변 저장 #######
			strSql = " UPDATE [db_board].[dbo].[tbl_scm_comment] SET " & _
					 "		comment = '" & vComment & "' "
			
			If vEtc1 <> "" Then
				strSql = strSql & " , etc1 = '" & vEtc1 & "' "
			End IF
			
			If vEtc2 <> "" Then
				strSql = strSql & " , etc2 = '" & vEtc2 & "' "
			End IF
			
			strSql = strSql & "	WHERE cIdx = '" & vCommentIdx & "' "
			dbget.execute strSql

		End IF

	End If
	
	Response.Write "<script>alert('처리되었습니다.');parent.document.location.reload();</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->