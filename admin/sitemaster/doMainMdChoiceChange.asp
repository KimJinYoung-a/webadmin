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
	dim strSQL, lp , realdate
	dim mode , realdatereset

	mode = request("mode")
	page = Request("page")
	menupos = Request("menupos")
	research = Request("research")
	sUsing = Request("sUsing")
	realdate = request("realdate")
	idx = Replace(Request("idx"), " ","")
	disporder = Replace(Request("disporder"), " ","")
	isusing = Replace(Request("isusing"), " ","")
	realdatereset = request("realdatereset")

	'�迭�� ����
	arrIdx = Split(idx,",")
	arrDispOrder = Split(disporder,",")
	arrIsUsing = Split(isusing,",")

	if mode = "" then 
		'���� ���� �ۼ�
		if Ubound(arrIdx)=0 then
			strSQL = "Update [db_sitemaster].[dbo].tbl_main_mdchoice_flash " &_
					"Set disporder=" & disporder & " " &_
					"Where idx=" & idx
		else
			for lp=0 to Ubound(arrIdx)
				strSQL = strSQL & "Update [db_sitemaster].[dbo].tbl_main_mdchoice_flash " & vbCrLf
				strSQL = strSQL & "Set disporder="& arrDispOrder(lp) &" " & vbCrLf
				strSQL = strSQL & "Where idx=" & arrIdx(lp) & ";" & vbCrLf
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
							"self.location='/admin/sitemaster/main_md_recommend_flash.asp?page=" & page & "&menupos=" & menupos & "&research=" & research & "&isusing=" & sUsing & "&realdate="& realdate &"';" &_
							"</script>"
			dbget.close()	:	response.End
		else
			dbget.RollbackTrans

			response.write "<script>alert('������ ������ �߻��߽��ϴ�.');history.back();</script>"
			dbget.close()	:	response.End
		end if

	elseif mode = "del" then 
		strSQL = "DELETE FROM [db_sitemaster].[dbo].tbl_main_mdchoice_flash Where idx=" & idx
		dbget.Execute strSQL
		
		If NOT(Err) then
			response.write "<script>" &_
							"alert('���� �Ǿ����ϴ�.');" &_
							"self.location='/admin/sitemaster/main_md_recommend_flash.asp?page=" & page & "&menupos=" & menupos & "&research=" & research & "&isusing=" & sUsing & "&realdate="& realdate &"&realdatereset="& realdatereset &"';" &_
							"</script>"
			dbget.close()	:	response.End
		end if
	end if 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->