<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/board_cls.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim mode, brdId, adminid
dim ansTitle, ansCont, commCd
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
brdId		= RequestCheckvar(Request("brdId"),10)
mode		= RequestCheckvar(Request("mode"),16)
commCd		= RequestCheckvar(Request("commCd"),16)
ansTitle	= html2db(RequestCheckvar(Request("ansTitle"),128))
ansCont	= html2db(Request("ansCont"))
page		= RequestCheckvar(Request("page"),10)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = RequestCheckvar(Request("searchString"),128)
adminid		= session("ssBctId")

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����


'==============================================================================
'## ���� ����(����) ó��

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "answer"
		'@@ �亯ó��
		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	  ansTitle= '" & ansTitle & "'" &_
				"	, ansCont = '" & ansCont & "'" &_
				"	, ansUserId = '" & Session("ssBctId") & "'" &_
				"	, ansDate = getdate() " &_
				"	, isanswer = 'Y' " &_
				" Where brdId = " & brdId

		dbACADEMYget.Execute(SQL)

		msg = "�亯ó���Ͽ����ϴ�."

		'���ư� ������
		retURL = "lec_board_view.asp?menupos=" & menupos & "&brdId=" & brdId & param

	Case "change"
		'@@ ���� ����

		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	commCd = '" & commCd & "'" &_
				" Where brdId = " & brdId
		dbACADEMYget.Execute(SQL)

		msg = "������ �����Ͽ����ϴ�."

		'���ư� ������
		retURL = "lec_board_list.asp?menupos=" & menupos & param

	Case "delete"
		'@@ ���� ����

		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	isusing = 'N'" &_
				" Where brdId = " & brdId
		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "lec_board_list.asp?menupos=" & menupos & param

End Select


'�����˻� �� �ݿ�
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'Ŀ��(����)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbACADEMYget.RollBackTrans				'�ѹ�(�����߻���)

	response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->