<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim posit_sn, posit_name, posit_isDel, mode
	Dim SQL, strMsg

	posit_sn = Request("posit_sn")
	posit_name = Request("posit_name")
	posit_isDel = Request("posit_isDel")
	mode = Request("mode")

	'Ʈ������ ����
	dbget.beginTrans

	'// ó�� �б� //
	Select Case mode
		Case "add"
			strMsg = "���������� ��ϵǾ����ϴ�."
			SQL =	"Insert into db_partner.dbo.tbl_positInfo " &_
					" (posit_name, posit_isDel) values " &_
					" ('" & posit_name & "'" &_
					" ,'N')"
			dbget.Execute(SQL)
		Case "modi"
			strMsg = "���������� �����Ǿ����ϴ�."
			SQL =	"Update db_partner.dbo.tbl_positInfo Set " &_
					"	posit_name = '" & posit_name & "' " &_
					"Where posit_sn=" & posit_sn
			dbget.Execute(SQL)
		Case "del"
			strMsg = "ó���� �Ϸ�Ǿ����ϴ�."
			SQL =	"Update db_partner.dbo.tbl_positInfo Set " &_
					"	posit_isDel = '" + posit_isDel + "' " &_
					"Where posit_sn=" & posit_sn
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