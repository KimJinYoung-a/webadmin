<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim mode, cplid, userid
dim title, cplCont, commCd, isusing
dim SQL
dim page, searchDiv, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
cplid		= RequestCheckvar(Request("cplid"),10)
mode		= RequestCheckvar(Request("mode"),16)
commCd		= RequestCheckvar(Request("commCd"),16)
cplCont	= html2db(Request("cplCont"))
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
		SQL =	"Insert into db_academy.dbo.tbl_Compliment " &_
				"	(cplCont, commCd) values " &_
				"	('" & cplCont & "'" &_
				"	,'" & commCd & "')"
		dbACADEMYget.Execute(SQL)

		'��� �޽���
		msg = "�����Ͽ����ϴ�."
		
		'���ư� ������
		retURL = "Compliment_list.asp?menupos=" & menupos & param


	Case "modify"
		'@@ ���� ����
		SQL =	"Update db_academy.dbo.tbl_Compliment Set " &_
				"	  cplCont = '" & cplCont & "'" &_
				"	, commCd = '" & commCd & "'" &_
				"	, isusing = '" & isusing & "'" &_
				" Where cplid = " & cplid

		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "Compliment_modi.asp?menupos=" & menupos & "&cplid=" & cplid & param

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