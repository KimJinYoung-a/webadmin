<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim mode, prfid, userid
dim title, prfCont, groupCd, commCd, isusing
dim SQL
dim page, searchDiv, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
prfid		= RequestCheckvar(Request("prfid"),10)
mode		= RequestCheckvar(Request("mode"),16)
groupCd		= RequestCheckvar(Request("groupCd"),16)
commCd		= RequestCheckvar(Request("commCd"),16)
prfCont	= html2db(Request("prfCont"))
isusing		= RequestCheckvar(Request("isusing"),2)
page		= RequestCheckvar(Request("page"),10)
searchDiv	= RequestCheckvar(Request("searchDiv"),16)
searchString = RequestCheckvar(Request("searchString"),128)

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchString=" & searchString	'������ ����


'==============================================================================
'## ���� ����(����) ó��

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ ���� ����
		SQL =	"Insert into db_academy.dbo.tbl_preface " &_
				"	(prfCont, groupCd, commCd) values " &_
				"	('" & prfCont & "'" &_
				"	,'" & groupCd & "'" &_
				"	,'" & commCd & "')"
		dbACADEMYget.Execute(SQL)

		'��� �޽���
		msg = "�����Ͽ����ϴ�."
		
		'���ư� ������
		retURL = "Preface_list.asp?menupos=" & menupos & param


	Case "modify"
		'@@ ���� ����
		SQL =	"Update db_academy.dbo.tbl_preface Set " &_
				"	  prfCont = '" & prfCont & "'" &_
				"	, groupCd = '" & groupCd & "'" &_
				"	, commCd = '" & commCd & "'" &_
				"	, isusing = '" & isusing & "'" &_
				" Where prfid = " & prfid

		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "Preface_modi.asp?menupos=" & menupos & "&prfid=" & prfid & param

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