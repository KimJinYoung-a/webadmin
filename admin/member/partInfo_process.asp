<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim part_sn, part_name, part_sort, part_isDel, mode
	Dim SQL, strMsg

	part_sn = Request("part_sn")
	part_name = Request("part_name")
	part_sort = Request("part_sort")
	part_isDel = Request("part_isDel")
	mode = Request("mode")

	'Ʈ������ ����
	dbget.beginTrans

	'// ó�� �б� //
	Select Case mode
		Case "add"
			strMsg = "�μ������� ��ϵǾ����ϴ�."
			SQL =	"Insert into db_partner.dbo.tbl_partInfo " &_
					" (part_name, part_sort, part_isDel) values " &_
					" ('" & part_name & "'" &_
					" ," & part_sort &_
					" ,'N')"
			dbget.Execute(SQL)
		Case "modi"
			strMsg = "�μ������� �����Ǿ����ϴ�."
			SQL =	"Update db_partner.dbo.tbl_partInfo Set " &_
					"	part_name = '" & part_name & "' " &_
					"	,part_sort = " & part_sort & " " &_
					"Where part_sn=" & part_sn
			dbget.Execute(SQL)
		Case "del"
			strMsg = "ó���� �Ϸ�Ǿ����ϴ�."
			SQL =	"Update db_partner.dbo.tbl_partInfo Set " &_
					"	part_isDel = '" + part_isDel + "' " &_
					"Where part_sn=" & part_sn
			dbget.Execute(SQL)
	End Select

	'�����˻� �� ����
	If Err.Number = 0 Then   
		dbget.CommitTrans				'Ŀ��(����)
	
		response.write	"<script language='javascript'>" &_
						"	alert('" & strMsg & "');" &_
						"	opener.history.go(0);" &_
						"	self.close();" &_
						"</script>"
	Else
	    dbget.RollBackTrans				'�ѹ�(�����߻���)
	
		response.write	"<script language='javascript'>" &_
						"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
						"	history.back();" &_
						"</script>"
	
	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->