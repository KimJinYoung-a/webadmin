<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim vQuery, vOrderSerial, vLecIdx, vOption, vOptionName
	vOrderSerial = RequestCheckvar(Request("orderserial"),16)
	vLecIdx = RequestCheckvar(Request("lec_idx"),10)
	vOption = RequestCheckvar(Request("option"),128)
  	if vOption <> "" then
		if checkNotValidHTML(vOption) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	If vOrderSerial <> "" AND vLecIdx <> "" AND vOption <> "" Then
		vQuery = "SELECT lecOptionName FROM [db_academy].[dbo].tbl_lec_item_option WHERE lecIdx = '" & vLecIdx & "' AND lecOption = '" & vOption & "'"
		rsACADEMYget.open vQuery,dbACADEMYget,1
		If Not rsACADEMYget.Eof Then
			vOptionName = db2Html(rsACADEMYget("lecOptionName"))
		End If
		rsACADEMYget.close
		
		vQuery = "UPDATE [db_academy].[dbo].[tbl_academy_order_detail] SET itemoption = '" & vOption & "', itemoptionname = '" & vOptionName & "' WHERE orderserial = '" & vOrderSerial & "'"
		dbACADEMYget.Execute(vQuery)
		
		rw "<script language='javascript'>alert('����Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
	Else
		rw "<script language='javascript'>alert('�߸��� ����Դϴ�.');window.close();</script>"
	End IF
%>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->