<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/programchangeCls.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->

<%
	Dim vQuery, vGubun, vIdx, vTitle, vContent, iCurrentpage, vFileName, vDoc_Idx, vChkList, vSign1Chk, vSign2Chk
	dim tmpStr, i
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	vGubun		= requestCheckVar(Request("gubun"),10)
	vIdx		= requestCheckVar(Request("pidx"),10)
	vTitle		= requestCheckVar(Request("title"),150)
	vContent	= requestCheckVar(Request("contents"),400)
	vFileName	= requestCheckVar(Request("filename"),2000)
	vDoc_Idx	= requestCheckVar(Request("didx"),10)
	vChkList	= Replace(requestCheckVar(Request("programchk"),50)," ","")
	vSign1Chk	= NullFillWith(requestCheckVar(Request("sign1chk"),1),0)
	vSign2Chk	= NullFillWith(requestCheckVar(Request("sign2chk"),1),0)


	If vGubun = "sign" Then
		If session("ssBctId") = "kobula" Then
			vQuery = "SELECT sign1 FROM [db_board].[dbo].[tbl_program_change] WHERE pidx = '" & vIdx & "'"
			rsget.Open vQuery,dbget,1
			If Not rsget.Eof Then
				If rsget(0) = "" Then
					vQuery = "UPDATE [db_board].[dbo].[tbl_program_change] SET sign1 = '" & session("ssBctId") & "', sign1date = getdate(), sign1chk = 1 WHERE pidx = '" & vIdx & "'" & vbCrLf
				Else
					vQuery = ""
				End If
			Else
				vQuery = ""
			End If
			rsget.close()
		End IF

		vQuery = vQuery & "UPDATE [db_board].[dbo].[tbl_program_change] SET "
		If session("ssBctId") = "kobula" Then
			vQuery = vQuery & "sign2 = '" & session("ssBctId") & "', "
			vQuery = vQuery & "sign2date = getdate(), "
			vQuery = vQuery & "chklist = '" & vChkList & "', "
			vQuery = vQuery & "sign2chk = 1 "
		End IF
		vQuery = vQuery & "WHERE pidx = '" & vIdx & "'"
		dbget.execute vQuery

		response.write "<script language='javascript'>alert('결제되었습니다.'); location.href='index.asp?menupos="&request("menupos")&"&iC="&iCurrentpage&"';</script>"
		dbget.close()
	    response.end
	ElseIf vGubun = "allsign" Then
		vDoc_Idx = Left(Request("allpidx"),Len(Request("allpidx"))-1)
		If session("ssBctId") = "kobula" Then
			vQuery = "UPDATE [db_board].[dbo].[tbl_program_change] SET sign1 = '" & session("ssBctId") & "', sign1date = getdate(), sign2 = '" & session("ssBctId") & "', sign2date = getdate(), sign1chk = 1, sign2chk = 1 "
			vQuery = vQuery & "WHERE pidx IN(select pidx from [db_board].[dbo].[tbl_program_change] where sign1 = '' and sign2 = '' and pidx in(" & vDoc_Idx & ")) "
			vQuery = vQuery & "UPDATE [db_board].[dbo].[tbl_program_change] SET sign2 = '" & session("ssBctId") & "', sign2date = getdate(), sign2chk = 1 "
			vQuery = vQuery & "WHERE pidx IN(select pidx from [db_board].[dbo].[tbl_program_change] where sign1 <> '' and sign2 = '' and pidx in(" & vDoc_Idx & ")) "
			dbget.execute vQuery
		End IF

		response.write "<script language='javascript'>alert('결제되었습니다.'); location.href='index.asp?menupos="&request("menupos")&"&iC="&iCurrentpage&"';</script>"
		dbget.close()
	    response.end
	Else
			If vIdx = "" Then
				vFileName = Split(vFileName, vbCrLf)
				for i = LBound(vFileName) to UBound(vFileName)
					tmpStr = Trim(vFileName(i))
					if (tmpStr <> "") then
						vQuery = "INSERT INTO [db_board].[dbo].[tbl_program_change](title, contents, filename, reguserid, doc_idx, chklist) VALUES('" & html2db(vTitle) & "', '" & html2db(vContent) & "', '" & tmpStr & "', '" & session("ssBctId") & "', '" & vDoc_Idx & "', '" & vChkList & "')"
						dbget.execute vQuery

						If session("ssBctId") = "kobula" Then
							vQuery = " SELECT SCOPE_IDENTITY() "
							rsget.Open vQuery,dbget
			 				IF Not rsget.EOF THEN
			 					vIdx = rsget(0)
			 				END IF
			 				rsget.close

							vQuery = "UPDATE [db_board].[dbo].[tbl_program_change] SET sign1 = '" & session("ssBctId") & "', sign1date = getdate(), sign1chk = 1 Where pidx = '" & vIdx & "'"
							dbget.execute vQuery

							If session("ssBctId") = "kobula" Then
								vQuery = "UPDATE [db_board].[dbo].[tbl_program_change] SET sign1 = '" & session("ssBctId") & "', sign1date = getdate(), sign2 = 'kobula', sign2date = getdate(), sign1chk = 1, sign2chk = 1 Where pidx = '" & vIdx & "'"
								dbget.execute vQuery
							End IF
		 				End IF
					end if
				next

				If vDoc_Idx <> "" Then
					response.write "<script language='javascript'>if(confirm('저장되었습니다.\n\n프로그램변경계획리스트 팝업을 열려면 [확인]\n이대로 창을 닫으려면 [취소]') == true){window.open('/admin/cooperate/program/index.asp?didx="&vDoc_Idx&"','ppprogram','width=1100,height=800,scrollbars=yes');window.close();}else{window.close();}opener.location.reload();</script>"
				Else
					response.write "<script language='javascript'>alert('저장되었습니다.'); location.href='index.asp?menupos="&request("menupos")&"&iC="&iCurrentpage&"';</script>"
			    End If
					dbget.close()
				    response.end
			Else
				If isNumeric(vIdx) = False Then
					response.write "<script language='javascript'>alert('잘못된 접근입니다.'); history.back();</script>"
					dbget.close()
				    response.end
				Else
					vQuery = "UPDATE [db_board].[dbo].[tbl_program_change] SET "
					vQuery = vQuery & "title = '" & html2db(vTitle) & "', "
					vQuery = vQuery & "contents = '" & html2db(vContent) & "', "
					vQuery = vQuery & "filename = '" & vFileName & "', "
					vQuery = vQuery & "chklist = '" & vChkList & "' "
					vQuery = vQuery & "WHERE pidx = '" & vIdx & "'"
					dbget.execute vQuery

					response.write "<script language='javascript'>alert('저장되었습니다.'); location.href='index.asp?menupos="&request("menupos")&"&iC="&iCurrentpage&"';</script>"
					dbget.close()
				    response.end
				End IF
			End IF
	End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
