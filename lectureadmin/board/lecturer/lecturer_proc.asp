<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  핑거스 강사 게시판
' History : 2010.03.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/board/lecturer/lecturer_cls.asp"-->

<%
Dim strSql, vWorkerTemp, vWorkerViewTemp, vFileTemp, i , g_MenuPos , mode
Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status,sDoc_Type, sDoc_Import, sDoc_Diffi, sDoc_Subj
dim sDoc_Content, sDoc_Worker, sDoc_File, sDoc_WorkerView, sDoc_UseYN
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	sDoc_Id			= session("ssBctId")
	sDoc_UseYN		= NullFillWith(requestCheckVar(Request("doc_useyn"),1),"Y")
	sDoc_Status		= NullFillWith(requestCheckVar(Request("K000"),24),1)	
	sDoc_Type		= NullFillWith(requestCheckVar(Request("G000"),24),0)
	sDoc_Import		= NullFillWith(requestCheckVar(Request("L000"),24),0)
	sDoc_Diffi		= NullFillWith(requestCheckVar(Request("doc_difficult"),2),0)
	sDoc_Worker		= NullFillWith(requestCheckVar(Request("doc_worker"),1000),"")	
	sDoc_Subj		= Request("doc_subject")
	sDoc_Content	= replace(Request("brd_content"),"'","")
	sDoc_File		= NullFillWith(Request("doc_file"),"")
	mode		= RequestCheckvar(Request("mode"),16)
	g_MenuPos = RequestCheckvar(request("menupos"),10)
	'response.write sDoc_Status
	'response.end
	
'On Error Resume Next
'dbacademyget.beginTrans

if mode = "edit" then
	if (checkNotValidHTML(sDoc_Subj) = true) Then
		response.write "<script>alert('제목에는 Script 또는 Action이나 HTML을 사용하실 수 없습니다.');history.back();</script>"
		dbget.Close
		''session.codePage = 949  
		response.End
	End If

	if (checkNotValidHTML(sDoc_Content) = true) Then
		response.write "<script>alert('내용에는 Script 또는 Action을 사용하실 수 없습니다.');history.back();</script>"
		dbget.Close
		''session.codePage = 949  
		response.End
	End If

	'//신규저장	
	If iDoc_Idx = "" Then
		strSql = " INSERT INTO [db_academy].dbo.tbl_lecturer_board_document " & _
				 "		(id, doc_type, doc_important, doc_difficult, doc_subject, doc_content, doc_status) " & _
				 "	VALUES " & _
				 "		('" & sDoc_Id & "','" & sDoc_Type & "', '" & sDoc_Import & "', '" & sDoc_Diffi & "', " & _
				 "		'" & sDoc_Subj & "', '" & html2db(replace(sDoc_Content,vbcrlf,"")) & "', '" & sDoc_Status & "') "
		
		'response.write strSql &"<br>"
		dbacademyget.execute strSql
		
		strSql = ""
		strSql = " SELECT SCOPE_IDENTITY() "
		rsacademyget.Open strSql,dbacademyget
		IF Not rsacademyget.EOF THEN
			iDoc_Idx = rsacademyget(0)
		ELSE
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "") 	
			''session.codePage = 949
		END IF
		rsacademyget.close
	
	'//수정	
	Else
		strSql = " UPDATE [db_academy].dbo.tbl_lecturer_board_document SET " & _
				 "		doc_type = '" & sDoc_Type & "', " & _
				 "		doc_important = '" & sDoc_Import & "', " & _
				 "		doc_difficult = '" & sDoc_Diffi & "', " & _
				 "		doc_subject = '" & sDoc_Subj & "', doc_content = '" & html2db(replace(sDoc_Content,vbcrlf,"")) & "', " & _
				 "		doc_status = '" & sDoc_Status & "', " & _				 
				 "		doc_useyn = '" & sDoc_UseYN & "' " & _
				 "	WHERE " & _
				 "		doc_idx = '" & iDoc_Idx & "' "
		
		'response.write strSql &"<br>"
		dbacademyget.execute strSql
	
	End If

	'####### 첨부파일 저장 #######
	If sDoc_File <> "" Then
		strSql = ""
		If iDoc_Idx <> "" Then
			strSql = " DELETE [db_academy].dbo.tbl_lecturer_board_file WHERE doc_idx = '" & iDoc_Idx & "' "
		End If
		vFileTemp = Split(sDoc_File, ",")
		For i = 0 To UBOUND(vFileTemp)
			strSql = strSql & " INSERT INTO [db_academy].[dbo].tbl_lecturer_board_file " & _
							  "		(file_name, doc_idx) " & _
							  "	VALUES " & _
							  "		('" & vFileTemp(i) & "', '" & iDoc_Idx & "') " & vbCrLf
		Next
		'response.write strSql &"<br>"
		dbacademyget.execute strSql
	Else
		If requestCheckVar(Request("isfile"),1) = "o" Then
			dbget.execute " DELETE [db_academy].dbo.tbl_lecturer_board_file WHERE doc_idx = '" & iDoc_Idx & "' "
		End If
	End If
		
	'0dbacademyget.RollBackTrans
	'dbacademyget.CommitTrans
	'Response.End
	'on error Goto 0

	If Request("gubun") = "write" Then
		Response.Write "<script>alert('OK');location.href='lecturer.asp?menupos="&g_MenuPos&"';</script>"
		''session.codePage = 949
	Else
		Response.Write "<script>alert('OK');location.href='lecturer.asp?menupos="&g_MenuPos&"';</script>"
		''session.codePage = 949
	End If

elseif mode = "view" then

	strSql = " UPDATE [db_academy].dbo.tbl_lecturer_board_document SET " & _
			 " doc_status = '" & sDoc_Status & "'" & _				 
			 " WHERE " & _
			 " doc_idx = '" & iDoc_Idx & "' "
	
	'response.write strSql &"<br>"
	dbacademyget.execute strSql

	Response.Write "<script>alert('OK');location.href='lecturer.asp?menupos="&g_MenuPos&"';</script>"
	''session.codePage = 949

elseif mode = "del" then

	strSql = " UPDATE [db_academy].dbo.tbl_lecturer_board_document SET " & _
			 " doc_useyn = 'N'" & _				 
			 " WHERE " & _
			 " doc_idx = '" & iDoc_Idx & "' "
	
	'response.write strSql &"<br>"
	dbacademyget.execute strSql

	Response.Write "<script>alert('OK'); opener.location.reload(); self.close();</script>"		
	''session.codePage = 949
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

<%
	''session.codePage = 949
%>