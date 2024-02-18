<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/board_cls.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim mode, brdId, lecUserId
dim qstTitle, qstCont, commCd
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= requestCheckVar(Request("menupos"),10)
brdId		= requestCheckVar(Request("brdId"),10)
mode		= requestCheckVar(Request("mode"),16)
commCd		= requestCheckVar(Request("commCd"),4)
qstTitle	= html2db(Request("qstTitle"))
qstCont	= html2db(Request("qstCont"))
page		= requestCheckVar(Request("page"),10)
searchDiv	= requestCheckVar(Request("searchDiv"),10)
searchKey	= requestCheckVar(Request("searchKey"),10)
searchString = requestCheckVar(Request("searchString"),128)
lecUserId	= session("ssBctId")

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

  	if qstTitle <> "" then
		if checkNotValidHTML(qstTitle) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end If
  	if qstCont <> "" then
		if checkNotValidHTML(qstCont) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end If
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if

'==============================================================================
'## ���� ����(����) ó��

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ ���� ����
		SQL =	"Insert into db_academy.dbo.tbl_lec_board (qstTitle, qstCont, lecUserId, commCd) values " &_
				"	 ('" & qstTitle & "'" &_
				"	, '" & qstCont & "'" &_
				"	, '" & lecUserId & "'" &_
				"	, '" & commCd & "')"

		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "board_list.asp?menupos=" & menupos & param

	Case "modify"
		'@@ ���� ����

		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	qstTitle = '" & qstTitle & "'" &_
				"	qstCont = '" & qstCont & "'" &_
				"	commCd = '" & commCd & "'" &_
				" Where brdId = " & brdId
		dbACADEMYget.Execute(SQL)

		msg = "������ �����Ͽ����ϴ�."

		'���ư� ������
		retURL = "board_view.asp?menupos=" & menupos & "&brdId=" & brdId & param

	Case "delete"
		'@@ ���� ����

		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	isusing = 'N'" &_
				" Where brdId = " & brdId
		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "board_list.asp?menupos=" & menupos & param

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
<!-- #include virtual="/lib/db/dbclose.asp" -->