<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim mode, ntcId, userid
dim title, contents, commCd
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
ntcId		= RequestCheckvar(Request("ntcId"),10)
mode		= RequestCheckvar(Request("mode"),16)
commCd		= RequestCheckvar(Request("commCd"),16)
title		= html2db(RequestCheckvar(Request("title"),128))
contents	= html2db(Request("contents"))
page		= RequestCheckvar(Request("page"),10)
searchDiv	= RequestCheckvar(request("searchDiv"),16)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = RequestCheckvar(Request("searchString"),128)
userid		= session("ssBctId")

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����


'==============================================================================
'## ���� ����(����) ó��

if (checkNotValidHTML(title) = true) Then
	response.write "<script>alert('�������� ���񿡴� HTML�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	dbget.Close
	response.End
End If

'' imgsrc / ahref �� üũ�ϴ� ����?	checkNotValidHTML = > checkNotValidHTMLcritical
''if (checkNotValidHTMLcritical(sBrd_content) = true) Then			'// img �±� ������� ���� > �˻��׸� ����ȭ
if (checkNotValidHTML(contents) = true) Then
	response.write "<script>alert('�������� ���뿡�� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	dbget.Close
	response.End
End If


'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ ���� ����
		SQL =	"Insert into db_academy.dbo.tbl_Notice " &_
				"	(title, contents, commCd, userid) values " &_
				"	('" & title & "'" &_
				"	,'" & contents & "'" &_
				"	,'" & commCd & "'" &_
				"	,'" & userid & "')"
		dbACADEMYget.Execute(SQL)

		'��� �޽���
		msg = "�����Ͽ����ϴ�."
		
		'���ư� ������
		retURL = "Notice_list.asp?menupos=" & menupos & param


	Case "modify"
		'@@ ���� ����
		SQL =	"Update db_academy.dbo.tbl_Notice Set " &_
				"	  title= '" & title & "'" &_
				"	, contents = '" & contents & "'" &_
				"	, commCd = '" & commCd & "'" &_
				" Where ntcId = " & ntcId

		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "Notice_view.asp?menupos=" & menupos & "&ntcId=" & ntcId & param

	Case "delete"
		'@@ ���� ����

		'# ���� ����
		SQL =	"Update db_academy.dbo.tbl_Notice Set " &_
				"	  title= '" & title & "'" &_
				"	, isusing = 'N'" &_
				" Where ntcId = " & ntcId

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "Notice_list.asp?menupos=" & menupos & param

End Select


'�����˻� �� �ݿ�
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'Ŀ��(����)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"</script>"
'"	self.location='" & retURL & "';" &_
Else
    dbACADEMYget.RollBackTrans				'�ѹ�(�����߻���)

	response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

End If

IF application("Svr_Info") = "Dev" THEN
	Response.Redirect "http://test.thefingers.co.kr/chtml/make_index_notice.asp?retURL=http://testwebadmin.10x10.co.kr/academy/board/notice_list.asp?menupos=784"
ELSE
	Response.Redirect "http://www.thefingers.co.kr/chtml/make_index_notice.asp?retURL=http://webadmin.10x10.co.kr/academy/board/notice_list.asp?menupos=784"
END IF

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->