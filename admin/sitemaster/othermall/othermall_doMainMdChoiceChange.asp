<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	dim page, menupos, research, sUsing
	dim idx, disporder, isusing
	dim arrIdx, arrDispOrder, arrIsUsing
	dim strSQL, lp

	page = Request("page")
	menupos = Request("menupos")
	research = Request("research")
	sUsing = Request("sUsing")
	idx = Replace(Request("idx"), " ","")
	disporder = Replace(Request("disporder"), " ","")
	isusing = Replace(Request("isusing"), " ","")

	'�迭�� ����
	arrIdx = Split(idx,",")
	arrDispOrder = Split(disporder,",")
	arrIsUsing = Split(isusing,",")

	'���� ���� �ۼ�
	if Ubound(arrIdx)=0 then
		strSQL = "Update [db_contents].[dbo].tbl_othermall_main_mdchoice_flash " &_
				"Set disporder=" & disporder & " " &_
				"	,isusing='" & isusing & "' " &_
				"Where idx=" & idx
	else
		for lp=0 to Ubound(arrIdx)
			strSQL = strSQL & "Update [db_contents].[dbo].tbl_othermall_main_mdchoice_flash " &_
								"Set disporder=" & arrDispOrder(lp) & " " &_
								"	,isusing='" & arrIsUsing(lp) & "' " &_
								"Where idx=" & arrIdx(lp) & ";" & vbCrLf
		next
	end if

	'// DB ���� //
	dbget.beginTrans	'Ʈ������ ����
	dbget.Execute strSQL

	'DB���� �� Ʈ������ ó��
	If Err.Number = 0 Then
		dbget.commitTrans

		response.write "<script>" &_
						"alert('�����Ǿ����ϴ�.');" &_
						"self.location='/admin/sitemaster/othermall/othermall_main_md_recommend_flash.asp?page=" & page & "&menupos=" & menupos & "&research=" & research & "&isusing=" & sUsing & "';" &_
						"</script>"
		dbget.close()	:	response.End
	else
		dbget.RollbackTrans

		response.write "<script>alert('������ ������ �߻��߽��ϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->