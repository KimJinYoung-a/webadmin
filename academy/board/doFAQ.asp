<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim mode, faqid, userid
dim title, contents, commCd
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
faqid		= RequestCheckvar(Request("faqid"),10)
mode		= RequestCheckvar(Request("mode"),16)
commCd		= RequestCheckvar(Request("commCd"),16)
title		= html2db(RequestCheckvar(Request("title"),128))
contents	= html2db(Request("contents"))
page		= RequestCheckvar(Request("page"),10)
searchDiv	= RequestCheckvar(Request("searchDiv"),16)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = RequestCheckvar(Request("searchString"),128)
userid		= session("ssBctId")

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����


'==============================================================================
'## ���� ����(����) ó��

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ ���� ����
		SQL =	"Insert into db_academy.dbo.tbl_faq " &_
				"	(title, contents, commCd, userid) values " &_
				"	('" & title & "'" &_
				"	,'" & contents & "'" &_
				"	,'" & commCd & "'" &_
				"	,'" & userid & "')"
		dbACADEMYget.Execute(SQL)

		'��� �޽���
		msg = "�����Ͽ����ϴ�."
		
		'���ư� ������
		retURL = "faq_list.asp?menupos=" & menupos & param


	Case "modify"
		'@@ ���� ����
		SQL =	"Update db_academy.dbo.tbl_faq Set " &_
				"	  title= '" & title & "'" &_
				"	, contents = '" & contents & "'" &_
				"	, commCd = '" & commCd & "'" &_
				" Where faqid = " & faqid

		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "faq_view.asp?menupos=" & menupos & "&faqid=" & faqid & param

	Case "delete"
		'@@ ���� ����

		'# ���� ����
		SQL =	"Update db_academy.dbo.tbl_faq Set " &_
				"	  title= '" & title & "'" &_
				"	, isusing = 'N'" &_
				" Where faqid = " & faqid

		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "faq_list.asp?menupos=" & menupos & param

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