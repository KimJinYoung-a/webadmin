<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim job_sn, job_name, job_isDel, mode
	Dim SQL, strMsg

	mode 		= requestCheckvar(request("mode"),32)
	job_sn 		= requestCheckvar(request("job_sn"),32)
	job_name 	= requestCheckvar(request("job_name"),32)
	job_isDel 	= requestCheckvar(request("job_isDel"),32)

	'Ʈ������ ����
	dbget.beginTrans

	'// ó�� �б� //
	Select Case mode
		Case "add"
			strMsg = "��å������ ��ϵǾ����ϴ�."
			SQL =	"Insert into db_partner.dbo.tbl_jobInfo " &_
					" (job_name, job_isDel) values " &_
					" ('" & job_name & "'" &_
					" ,'N')"
			dbget.Execute(SQL)
		Case "modi"
			strMsg = "��å������ �����Ǿ����ϴ�."
			SQL =	"Update db_partner.dbo.tbl_jobInfo Set " &_
					"	job_name = '" & job_name & "' " &_
					"Where job_sn=" & job_sn
			dbget.Execute(SQL)
		Case "del"
			strMsg = "ó���� �Ϸ�Ǿ����ϴ�."
			SQL =	"Update db_partner.dbo.tbl_jobInfo Set " &_
					"	job_isDel = '" + job_isDel + "' " &_
					"Where job_sn=" & job_sn
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