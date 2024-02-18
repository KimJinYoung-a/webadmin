<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// 변수 선언
	dim mode, idx, menupos
	dim page, searchKey, searchString, searchConfirm, confirmMemo
	dim adminId, RetUrl, strMsg, param
	
	dim sqlStr

	'// 파라메터 접수
	mode			= RequestCheckvar(Request("mode"),16)
	idx				= RequestCheckvar(Request("idx"),10)
	menupos			= RequestCheckvar(Request("menupos"),10)
	page			= RequestCheckvar(Request("page"),10)
	searchKey		= RequestCheckvar(Request("searchKey"),16)
	searchString	= html2db(Request("searchString"))
	searchConfirm	= RequestCheckvar(Request("searchConfirm"),1)
	confirmMemo		= html2db(Request("confirmMemo"))
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if confirmMemo <> "" then
		if checkNotValidHTML(confirmMemo) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	'페이지 이동 파라메터
	param = "&searchKey=" & searchKey & "&searchString=" & server.URLencode(searchString) &_
			"&searchConfirm=" & searchConfirm & "&menupos=" & menupos	'페이지 변수


'=========================== 모드별 처리 분기 ============================
Select Case mode
	'========== 강사신청 ==========
	Case "AnsLeturer"
        '// 강사신청 문의 답변
        sqlStr =	" Update db_academy.dbo.tbl_partner_lecturer Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "처리하였습니다."
		RetUrl = "/academy/Partnership/partnerLecture_List.asp?page=" & page & param

	Case "DelLeturer"
        '// 강사신청 문의 삭제
        sqlStr =	" Update db_academy.dbo.tbl_partner_lecturer Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "삭제하였습니다."
		RetUrl = "/academy/Partnership/partnerLecture_List.asp?page=" & page & param
	
	'========== 작가신청 ==========
	Case "AnsWriter"
        '// 작가신청 문의 답변
        sqlStr =	" Update db_academy.dbo.tbl_partner_writer Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "처리하였습니다."
		RetUrl = "/academy/Partnership/partnerWriter_List.asp?menupos="&menupos

	Case "DelWriter"
        '// 작가신청 문의 삭제
        sqlStr =	" Update db_academy.dbo.tbl_partner_writer Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "삭제하였습니다."
		RetUrl = "/academy/Partnership/partnerWriter_List.asp?menupos="&menupos

	'========== 현장강좌 ==========
	Case "AnsField"
        '// 현장강좌 문의 답변
        sqlStr =	" Update db_academy.dbo.tbl_partner_fieldlecture Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "처리하였습니다."
		RetUrl = "/academy/Partnership/partnerFieldLecture_List.asp?page=" & page & param

	Case "DelField"
        '// 현장강좌 문의 삭제
        sqlStr =	" Update db_academy.dbo.tbl_partner_fieldlecture Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "삭제하였습니다."
		RetUrl = "/academy/Partnership/partnerFieldLecture_List.asp?page=" & page & param

	'========== 단체수강 ==========
	Case "AnsGroup"
        '// 단체수강 문의 답변
        sqlStr =	" Update db_academy.dbo.tbl_partner_masslecture Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "처리하였습니다."
		RetUrl = "/academy/Partnership/partnerGroupLecture_List.asp?page=" & page & param

	Case "DelGroup"
        '// 단체수강 문의 삭제
        sqlStr =	" Update db_academy.dbo.tbl_partner_masslecture Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "삭제하였습니다."
		RetUrl = "/academy/Partnership/partnerGroupLecture_List.asp?page=" & page & param


	'========== 제휴광고 ==========
	Case "AnsJoint"
        '// 단체수강 문의 답변
        sqlStr =	" Update db_academy.dbo.tbl_partner_joinadv Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "처리하였습니다."
		RetUrl = "/academy/Partnership/partnerJoinAdv_List.asp?page=" & page & param

	Case "DelJoint"
        '// 단체수강 문의 삭제
        sqlStr =	" Update db_academy.dbo.tbl_partner_joinadv Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "삭제하였습니다."
		RetUrl = "/academy/Partnership/partnerJoinAdv_List.asp?page=" & page & param

end Select

response.write "<script>alert('" & strMsg & ".'); location.href = '" & RetUrl & "';</script>"
dbACADEMYget.close()	:	response.End

Set Request = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->