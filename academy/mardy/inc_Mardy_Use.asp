<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<%
	'// ���� ���� //
	dim Idx, mode, msg, SQL

	'// �Ķ���� ���� //
	Idx = RequestCheckvar(request("Idx"),10)
	mode = RequestCheckvar(request("mode"),16)

	Select Case mode
		Case "ScrapUse"
			SQL =	"Update db_academy.dbo.tbl_mardyScrap Set " &_
					"	isusing = 'Y' " &_
					"Where scrapId=" & Idx
			msg = "���"

		Case "ScrapDel"
			SQL =	"Update db_academy.dbo.tbl_mardyScrap Set " &_
					"	isusing = 'N' " &_
					"Where scrapId=" & Idx
			msg = "����"

		Case "StoryUse"
			SQL =	"Update db_academy.dbo.tbl_mardyStory Set " &_
					"	isusing = 'Y' " &_
					"Where storyId=" & Idx
			msg = "���"

		Case "StoryDel"
			SQL =	"Update db_academy.dbo.tbl_mardyStory Set " &_
					"	isusing = 'N' " &_
					"Where storyId=" & Idx
			msg = "����"

		Case "TipUse"
			SQL =	"Update db_academy.dbo.tbl_mardyTip Set " &_
					"	isusing = 'Y' " &_
					"Where tipId=" & Idx
			msg = "���"

		Case "TipDel"
			SQL =	"Update db_academy.dbo.tbl_mardyTip Set " &_
					"	isusing = 'N' " &_
					"Where tipId=" & Idx
			msg = "����"

		Case else
			dbget.close()	:	response.End
	End Select

	'Ʈ������ ����
	dbACADEMYget.beginTrans

	'// ��뿩�� ó��
	dbACADEMYget.Execute(SQL)


	'// �����˻� �� �ݿ�
	If Err.Number = 0 Then   
		dbACADEMYget.CommitTrans				'Ŀ��(����)
	
		response.write	"<script language='javascript'>" &_
						"	alert('��������� [" & msg & "]���� �����Ͽ����ϴ�.');" &_
						"	parent.history.go(0);" &_
						"</script>"
	Else
	    dbACADEMYget.RollBackTrans				'�ѹ�(�����߻���)
	
		response.write	"<script language='javascript'>" &_
						"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
						"</script>"
	
	End If
%>