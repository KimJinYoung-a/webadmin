<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim mode
dim groupCd, commCd, commNm, isusing
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
mode		= RequestCheckvar(Request("mode"),16)
groupCd		= RequestCheckvar(Request("groupCd"),10)
commCd		= RequestCheckvar(Request("commCd"),10)
commNm		= html2db(RequestCheckvar(Request("commNm"),32))
isusing		= RequestCheckvar(Request("isusing"),2)
page		= RequestCheckvar(Request("page"),10)
searchDiv	= RequestCheckvar(Request("searchDiv"),16)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = RequestCheckvar(Request("searchString"),128)

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����


'==============================================================================
'## ���� ����(����) ó��

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ �űԵ��
		'�ߺ��˻�
		SQL = "Select count(commCd) as cnt From db_academy.dbo.tbl_CommCd where commCd='" & commCd & "'"
		rsACADEMYget.Open sql, dbACADEMYget, 1
			if rsACADEMYget("cnt")>0 then
				response.write	"<script language='javascript'>" &_
								"	alert('�ߺ��� �ڵ带 �Է��Ͽ����ϴ�.');" &_
								"	history.back();" &_
								"</script>"
				dbget.close()	:	response.End
			end if
		rsACADEMYget.close

		'����
		SQL =	"Insert into db_academy.dbo.tbl_CommCd (groupCd, commCd, commNm) " &_
				"	Values " &_
				"	( '" & groupCd & "'" &_
				"	, '" & commCd & "'" &_
				"	, '" & commNm & "') "

		dbACADEMYget.Execute(SQL)

		msg = "�ű� ����Ͽ����ϴ�."

		'���ư� ������
		retURL = "CommCd_list.asp?menupos=" & menupos & param

	Case "modify"
		'@@ ����ó��

		SQL =	"Update db_academy.dbo.tbl_CommCd Set " &_
				"	commNm = '" & commNm & "'" &_
				"	,isUsing = '" & isusing & "'" &_
				" Where CommCd = '" & CommCd & "'"
		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "CommCd_list.asp?menupos=" & menupos & param

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