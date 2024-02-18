<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  공지사항 프로세스
' History : 2012.03.09 김진영 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/photo_req/boardCls.asp"-->
<%
Dim sMode, sbrd_Id, sBrd_sn, sPosit_sn, sBrd_subject, sBrd_content, sBrd_team, sBrd_fixed
Dim sDoc_File, sDoc_RealFile, sBrd_isusing
Dim	vFileTemp, vRFileTemp
Dim strSql, i
	sbrd_Id				= session("ssBctId")
	sBrd_sn				= Request("brd_sn")

	sBrd_subject 		= Request("brd_subject")
	sBrd_content 		= Request("brd_content")
	sBrd_fixed 			= Request("brd_fixed")
	sMode	 			= Request("mode")
	sDoc_File			= NullFillWith(Request("doc_file"),"")
	sDoc_RealFile		= NullFillWith(Request("doc_realfile"),"")
	sBrd_isusing		= Request("brd_isusing")


If sMode = "add" Then

	strSql = ""
	strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_bbs " & vbcrlf
	strSql = strSql & " (bbs_id, bbs_title, bbs_content, bbs_regdate, Brd_isusing, brd_fixed) " & vbcrlf
	strSql = strSql & "	VALUES " & vbcrlf
	strSql = strSql & "	('" & sbrd_Id & "', '" & html2db(sBrd_subject) & "', '" & html2db(sBrd_content) & "', getdate(), 'N', '" & sBrd_fixed & "')"
	dbget.execute strSql


	'####### 첨부파일 저장 #######
	If sDoc_File <> "" Then
		strSql = ""
		If sBrd_sn <> "" Then
			strSql = " DELETE [db_partner].[dbo].tbl_photo_file WHERE bbs_no = '" & sBrd_sn & "' "
		End If
		vFileTemp 	= Split(sDoc_File, ",")
		vRFileTemp	= Split(sDoc_RealFile, ",")

		For i = 0 To UBOUND(vFileTemp)
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_file " & _
							  "		(file_name, real_name, bbs_no, file_regdate) " & _
							  "	VALUES " & _
							  "		('" & Trim(vFileTemp(i)) & "','" & Trim(vRFileTemp(i)) & "','" & sBrd_sn & "', getdate()) " & vbCrLf

		Next
		dbget.execute strSql
	Else
		If requestCheckVar(Request("isfile"),1) = "o" Then
			dbget.execute " DELETE [db_partner].[dbo].tbl_photo_file WHERE bbs_no = '" & sBrd_sn & "' "
		End If
	End If


	Response.Write "<script>alert('저장되었습니다.');location.href='/admin/photo_req/board_list.asp?';</script>"
	dbget.close()
	Response.End

ElseIf sMode = "modify" Then
	Dim fixed, isusing
	fixed 	= Request("fixed")
	isusing	= Request("isusing")
	
	strSql = ""
	strSql = strSql & " UPDATE [db_partner].[dbo].tbl_photo_bbs SET " & vbcrlf
	strSql = strSql & " bbs_title = '" & html2db(sBrd_subject) & "' ,bbs_content = '" & html2db(sBrd_content) & "' ,brd_isusing = '" & sBrd_isusing & "', brd_fixed = '" & fixed & "'"
	strSql = strSql & " where bbs_no = '"& sBrd_sn &"' "
	'response.write strSql
	dbget.execute strSql

	'####### 첨부파일 저장 #######
	If sDoc_File <> "" Then
		strSql = ""
		If sBrd_sn <> "" Then
			strSql = " DELETE [db_partner].[dbo].tbl_photo_file WHERE bbs_no = '" & sBrd_sn & "' "
		End If
		vFileTemp 	= Split(sDoc_File, ",")
		vRFileTemp	= Split(sDoc_RealFile, ",")
		'response.write UBOUND(vFileTemp)
		For i = 0 To UBOUND(vFileTemp)
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_file " & _
							  "		(file_name, real_name, bbs_no, file_regdate) " & _
							  "	VALUES " & _
							  "		('" & Trim(vFileTemp(i)) & "', '" & Trim(vRFileTemp(i)) & "', '" & sBrd_sn & "', getdate()) " & vbCrLf
		Next
		dbget.execute strSql
	Else
		If requestCheckVar(Request("isfile"),1) = "o" Then
			dbget.execute " DELETE [db_partner].[dbo].tbl_photo_file WHERE bbs_no = '" & sBrd_sn & "' "
		End If
	End If

	
	Response.Write "<script>alert('수정되었습니다.');location.href='/admin/photo_req/board_list.asp';</script>"
	dbget.close()
	Response.End
	
ElseIf sMode = "count" Then
	strSql = " UPDATE [db_partner].[dbo].tbl_photo_bbs SET bbs_hit = bbs_hit + 1 where bbs_no = "& sBrd_sn
	dbget.execute strSql
	response.redirect "board_view.asp?brd_sn="& sBrd_sn
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
