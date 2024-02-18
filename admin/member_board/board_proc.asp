<%@ language="VBScript" %>
<% option explicit %>

<%
'####################################################
' Description :  공지사항 프로세스
' History : 2011.02.23 김진영 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<%
Dim g_MenuPos
IF application("Svr_Info")="Dev" THEN
	g_MenuPos   = "1288"		'### 메뉴번호 지정.
Else
	g_MenuPos   = "1304"		'### 메뉴번호 지정.
End If

Dim sMode, sbrd_Id, sBrd_sn, sOpen_team, sPart_sn, sJob_sn, sPosit_sn, sBrd_subject, sBrd_content, sBrd_fixed, sBrd_team, sBrd_type
Dim isNotify
Dim sDoc_File, sDoc_RealFile, sBrd_isusing
Dim	vFileTemp, vRFileTemp
Dim strSql, i, AssignedRow
dim department_id
dim startDate, endDate
dim menupos

sbrd_Id				= session("ssBctId")
sBrd_sn				= requestCheckvar(Request("brd_sn"),8)
sOpen_team 			= requestCheckvar(Request("open_team"),1)
startDate 			= requestCheckvar(Request("startDate"),10)
endDate 			= requestCheckvar(Request("endDate"),10)
isNotify			= requestCheckvar(Request("isNotify"),1)
menupos				= requestCheckvar(request("menupos"),8)

'If Isempty(Request("part_sn")) = "True" Then
'	sPart_sn = "1"
'	sPart_sn			= Split(sPart_sn,",")
'Else
'	sPart_sn			= Split(Request("part_sn"),",")
'End If

If sOpen_team = "Y" Then
	department_id       = 0 '전체공지
	department_id       =split(department_id,",")
	sBrd_team = "부서전체"
Else
	department_id       = split(request("arrdid"),",") '팀공지
	For i = 0 to Ubound(department_id)
		strSql = " select departmentNameFull from db_partner.dbo.vw_user_department where cid = '" & Trim(department_id(i)) & "' "
		rsget.Open strSql,dbget,1
		if not rsget.eof then
			sBrd_team = sBrd_team & rsget("departmentNameFull")&","
		end if
		rsget.close
	Next
End If

sJob_sn 			= requestCheckvar(Request("job_sn"),8)
sPosit_sn 			= requestCheckvar(Request("posit_sn"),8)
sBrd_subject 		= html2db(Request("brd_subject"))
sBrd_content 		= html2db(Request("brd_content"))
sBrd_fixed 			= requestCheckvar(Request("brd_fixed"),2)
sBrd_isusing		= requestCheckvar(Request("brd_isusing"),2)
sBrd_type			= requestCheckvar(Request("brd_type"),3)
sMode	 			= requestCheckvar(Request("mode"),6)
sDoc_File			= NullFillWith(Request("sFile"),"")
sDoc_RealFile		= NullFillWith(Request("sRFile"),"")
isNotify			= requestCheckvar(Request("isNotify"),1)

If (checkNotValidHTML(sBrd_subject) = true) Then
	response.write "<script>alert('공지사항 제목에는 HTML을 사용하실 수 없습니다.');history.back();</script>"
	dbget.Close
	response.End
End If

'' imgsrc / ahref 도 체크하는 이유?	checkNotValidHTML = > checkNotValidHTMLcritical
''if (checkNotValidHTMLcritical(sBrd_content) = true) Then			'// img 태그 허용으로 수정 > 검사항목 단일화
If (checkNotValidHTML(sBrd_content) = true) Then
	response.write "<script>alert('공지사항 내용에는 Script 또는 Action을 사용하실 수 없습니다.');history.back();</script>"
	dbget.Close
	response.End
End If

If sMode = "add" Then
	strSql = ""
	strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_cooperate_board " & vbcrlf
	strSql = strSql & " (id, brd_subject, brd_content, brd_fixed, brd_team, brd_type, startDate, endDate, isNotify) " & vbcrlf
	strSql = strSql & "	VALUES " & vbcrlf
	strSql = strSql & "	('" & sbrd_Id & "', '" & html2db(sBrd_subject) & "', '" & html2db(sBrd_content) & "', '" & sBrd_fixed & "', '" & sBrd_team & "', '" & sBrd_type & "', '" & startDate & "', '" & endDate & "', '" & isNotify & "')"
  	dbget.execute strSql

	For i = 0 to Ubound(department_id)
		strSql = ""
		strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_cooperate_board_part " & vbcrlf
		strSql = strSql & " (part_sn, posit_sn, job_sn, brd_sn,department_id) " & vbcrlf
		strSql = strSql & "	VALUES " & vbcrlf
		strSql = strSql & "	('0', '" & sPosit_sn & "', '" & sJob_sn & "', '" & sBrd_sn & "','"&trim(department_id(i))&"')"
		dbget.execute strSql
	Next

	'####### 첨부파일 저장 #######
	If sDoc_File <> "" Then
		strSql = ""
		If sBrd_sn <> "" Then
			strSql = " DELETE [db_partner].[dbo].tbl_cooperate_file WHERE brd_sn = '" & sBrd_sn & "' "
		End If
		vFileTemp 	= Split(sDoc_File, ",")
		vRFileTemp	= Split(sDoc_RealFile, ",")

		For i = 0 To UBOUND(vFileTemp)
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_cooperate_file " & _
							  "		(file_name, real_name, brd_sn) " & _
							  "	VALUES " & _
							  "		('" & Trim(vFileTemp(i)) & "', '" & Trim(vRFileTemp(i)) & "', '" & sBrd_sn & "') " & vbCrLf
		Next
		dbget.execute strSql
	Else
		If requestCheckVar(Request("isfile"),1) = "o" Then
			dbget.execute " DELETE [db_partner].[dbo].tbl_cooperate_file WHERE brd_sn = '" & sBrd_sn & "' "
		End If
	End If

	'잔디 웹훅 : Y, 직급 전체 : 17 일 때만 웹훅 전송
	If sOpen_team = "Y" and isNotify = "Y" and sPosit_sn = "17" Then
		Call fnJandiCall(sBrd_sn)
	End If

	Response.Write "<script>alert('저장되었습니다.');location.href='/admin/member_board/board_list.asp?menupos="&menupos&"&brd_type=" & sBrd_type & "';</script>"
	dbget.close()
	Response.End
ElseIf sMode = "modify" Then
	Dim fixed, isusing
	fixed 	= requestCheckvar(Request("brd_fixed"),2)
	isusing	= requestCheckvar(Request("isusing"),1)

	strSql = ""
	strSql = strSql & " UPDATE [db_partner].[dbo].tbl_cooperate_board SET " & vbcrlf
	strSql = strSql & " brd_subject = '" & html2db(sBrd_subject) & "' ,brd_content = '" & html2db(sBrd_content) & "', brd_fixed = '" & fixed & "', brd_team = '" & sBrd_team & "', brd_type = '" & sBrd_type & "', brd_isusing = '" & isusing & "', startDate = '" & startDate & "', endDate = '" & endDate & "', lastupdate = getdate() "
	strSql = strSql & " ,isNotify = '" & isNotify & "'  "
	strSql = strSql & " where brd_sn = '"& sBrd_sn &"' "
	if not(C_ADMIN_AUTH or C_PSMngPart) then
		strSql = strSql & " and id='"&sbrd_Id&"'" ''등록자만 수정 가능 2014/03/04 수정
	end if
	'response.write strSql
	dbget.execute strSql, AssignedRow

    If (AssignedRow < 1) Then
        response.write "권한이없습니다."
        response.end
    End If

	strSql = ""
	strSql = strSql & " DELETE FROM [db_partner].[dbo].tbl_cooperate_board_part "
	strSql = strSql & " where brd_sn = '" & sBrd_sn & "'"
	'response.write strSql
	dbget.execute strSql

	For i = 0 to Ubound(department_id)
		strSql = ""
		strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_cooperate_board_part " & vbcrlf
		strSql = strSql & " (part_sn, posit_sn, job_sn, brd_sn,department_id) " & vbcrlf
		strSql = strSql & "	VALUES " & vbcrlf
		strSql = strSql & "	('0', '" & sPosit_sn & "', '" & sJob_sn & "', '" & sBrd_sn & "','"&trim(department_id(i))&"')"
		dbget.execute strSql
	Next

	'####### 첨부파일 저장 #######
	If sDoc_File <> "" Then
		strSql = ""
		If sBrd_sn <> "" Then
			strSql = " DELETE [db_partner].[dbo].tbl_cooperate_file WHERE brd_sn = '" & sBrd_sn & "' "
		End If
		vFileTemp 	= Split(sDoc_File, ",")
		vRFileTemp	= Split(sDoc_RealFile, ",")
		'response.write UBOUND(vFileTemp)
		For i = 0 To UBOUND(vFileTemp)
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_cooperate_file " & _
							  "		(file_name, real_name, brd_sn) " & _
							  "	VALUES " & _
							  "		('" & Trim(vFileTemp(i)) & "', '" & Trim(vRFileTemp(i)) & "', '" & sBrd_sn & "') " & vbCrLf
		Next
		dbget.execute strSql
	Else
		If requestCheckVar(Request("isfile"),1) = "o" Then
			dbget.execute " DELETE [db_partner].[dbo].tbl_cooperate_file WHERE brd_sn = '" & sBrd_sn & "' "
		End If
	End If

	Response.Write "<script>alert('수정되었습니다.');location.href='/admin/member_board/board_list.asp?menupos="&menupos&"&brd_type=" & sBrd_type & "';</script>"
	dbget.close()
	Response.End
ElseIf sMode = "count" Then
	strSql = " UPDATE [db_partner].[dbo].tbl_cooperate_board SET brd_hit = brd_hit + 1 where brd_sn = "& sBrd_sn
	dbget.execute strSql
	response.redirect "board_view.asp?brd_sn="& sBrd_sn&"&menupos="&menupos&"&brd_type="&sBrd_type
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
