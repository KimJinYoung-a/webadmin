<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim level_sn, level_no, level_name, level_isDel, strLevel, mode
	Dim SQL, strMsg, strCnt

	level_sn = Request("level_sn")
	level_no = Request("level_no")
	level_name = Request("level_name")
	level_isDel = Request("level_isDel")
	strLevel = Request("strLevel")
	mode = Request("mode")

	'Ʈ������ ����
	dbget.beginTrans

	'// ó�� �б� //
	Select Case mode
		Case "add"
			strMsg = "��������� ��ϵǾ����ϴ�."
			SQL =	"Insert into db_partner.dbo.tbl_level " &_
					" (level_no, level_name, level_isDel) values " &_
					" (" & level_no &_
					" ,'" & level_name & "'" &_
					" ,'N')"
			dbget.Execute(SQL)
		Case "modi"
			strMsg = "��������� �����Ǿ����ϴ�."
			SQL =	"Update db_partner.dbo.tbl_level Set " &_
					"	level_no = " & level_no &_
					"	, level_name = '" & level_name & "' " &_
					"Where level_sn=" & level_sn
			dbget.Execute(SQL)
		Case "del"
			strMsg = "ó���� �Ϸ�Ǿ����ϴ�."
			SQL =	"Update db_partner.dbo.tbl_level Set " &_
					"	level_isDel = '" + level_isDel + "' " &_
					"Where level_sn=" & level_sn
			dbget.Execute(SQL)
		Case "dp_chk"
			SQL =	"Select count(*) From db_partner.dbo.tbl_level " &_
					"Where level_no=" & strLevel
			rsget.Open SQL,dbget,1
				strCnt = rsget(0)
			rsget.Close
			if strCnt>0 then
				response.write	"<script language='javascript'>" &_
								"	alert('�̹� ������� ��޹�ȣ�Դϴ�.\n�ٸ� ��޹�ȣ�� �������ֽʽÿ�.');" &_
								"</script>"
					response.End
			else
				response.write	"<script language='javascript'>" &_
								"	alert('��밡���� ��޹�ȣ�Դϴ�.');" &_
								"	parent.document.frm_level.level_no.value='" & strLevel & "';" &_
								"	parent.document.frm_level.level_name.focus();" &_
								"</script>"
					response.End
			end if
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