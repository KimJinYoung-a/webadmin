<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// ���� ����
	dim mode, idx, menupos
	dim page, searchKey, searchString, searchConfirm, confirmMemo
	dim adminId, RetUrl, strMsg, param
	
	dim sqlStr

	'// �Ķ���� ����
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
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if confirmMemo <> "" then
		if checkNotValidHTML(confirmMemo) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	'������ �̵� �Ķ����
	param = "&searchKey=" & searchKey & "&searchString=" & server.URLencode(searchString) &_
			"&searchConfirm=" & searchConfirm & "&menupos=" & menupos	'������ ����


'=========================== ��庰 ó�� �б� ============================
Select Case mode
	'========== �����û ==========
	Case "AnsLeturer"
        '// �����û ���� �亯
        sqlStr =	" Update db_academy.dbo.tbl_partner_lecturer Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "ó���Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerLecture_List.asp?page=" & page & param

	Case "DelLeturer"
        '// �����û ���� ����
        sqlStr =	" Update db_academy.dbo.tbl_partner_lecturer Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "�����Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerLecture_List.asp?page=" & page & param
	
	'========== �۰���û ==========
	Case "AnsWriter"
        '// �۰���û ���� �亯
        sqlStr =	" Update db_academy.dbo.tbl_partner_writer Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "ó���Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerWriter_List.asp?menupos="&menupos

	Case "DelWriter"
        '// �۰���û ���� ����
        sqlStr =	" Update db_academy.dbo.tbl_partner_writer Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "�����Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerWriter_List.asp?menupos="&menupos

	'========== ���尭�� ==========
	Case "AnsField"
        '// ���尭�� ���� �亯
        sqlStr =	" Update db_academy.dbo.tbl_partner_fieldlecture Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "ó���Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerFieldLecture_List.asp?page=" & page & param

	Case "DelField"
        '// ���尭�� ���� ����
        sqlStr =	" Update db_academy.dbo.tbl_partner_fieldlecture Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "�����Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerFieldLecture_List.asp?page=" & page & param

	'========== ��ü���� ==========
	Case "AnsGroup"
        '// ��ü���� ���� �亯
        sqlStr =	" Update db_academy.dbo.tbl_partner_masslecture Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "ó���Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerGroupLecture_List.asp?page=" & page & param

	Case "DelGroup"
        '// ��ü���� ���� ����
        sqlStr =	" Update db_academy.dbo.tbl_partner_masslecture Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "�����Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerGroupLecture_List.asp?page=" & page & param


	'========== ���ޱ��� ==========
	Case "AnsJoint"
        '// ��ü���� ���� �亯
        sqlStr =	" Update db_academy.dbo.tbl_partner_joinadv Set " &_
        			"		  confirmyn= 'Y' " &_
        			"		, confirmMemo = '" & confirmMemo & "' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "ó���Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerJoinAdv_List.asp?page=" & page & param

	Case "DelJoint"
        '// ��ü���� ���� ����
        sqlStr =	" Update db_academy.dbo.tbl_partner_joinadv Set " &_
        			"		  deleteyn = 'Y' " &_
        			" Where idx = " & idx
        dbACADEMYget.Execute(sqlStr)

		strMsg = "�����Ͽ����ϴ�."
		RetUrl = "/academy/Partnership/partnerJoinAdv_List.asp?page=" & page & param

end Select

response.write "<script>alert('" & strMsg & ".'); location.href = '" & RetUrl & "';</script>"
dbACADEMYget.close()	:	response.End

Set Request = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->