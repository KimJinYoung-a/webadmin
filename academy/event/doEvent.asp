<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̺�Ʈ
' History : 2010.09.17 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim msg, lp, menupos ,evtTitle, evtCont, evtSdate, evtEdate, isComment
dim mode, evtId ,SQL ,page, searchKey, searchString, param, retURL
	menupos		= RequestCheckvar(Request("menupos"),10)
	evtId		= RequestCheckvar(Request("evtId"),10)
	mode		= RequestCheckvar(Request("mode"),10)
	evtTitle		= html2db(RequestCheckvar(Request("evtTitle"),64))
	evtCont	= html2db(Request("evtCont"))
	page		= RequestCheckvar(Request("page"),10)
	searchKey	= RequestCheckvar(Request("searchKey"),16)
	searchString = RequestCheckvar(Request("searchString"),128)
	evtSdate = Request("syy") & "-" & Request("smm") & "-" & Request("sdd")
	evtEdate = Request("eyy") & "-" & Request("emm") & "-" & Request("edd")
	isComment = Request("isComment")

  	if evtCont <> "" then
		if checkNotValidHTML(evtCont) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
  	if isComment <> "" then
		if checkNotValidHTML(isComment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
param = "&page=" & page & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

'==============================================================================
'## ���� ����(����) ó��

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ ���� ����
		SQL =	"Insert into db_academy.dbo.tbl_eventInfo " &_
				"	(evtTitle, evtCont, evtSdate, evtEdate, isComment) values " &_
				"	('" & evtTitle & "'" &_
				"	,'" & evtCont & "'" &_
				"	,'" & evtSdate & "'" &_
				"	,'" & evtEdate & "'" &_
				"	,'" & isComment & "')"
		dbACADEMYget.Execute(SQL)

		'��� �޽���
		msg = "�����Ͽ����ϴ�."
		
		'���ư� ������
		retURL = "Event_list.asp?menupos=" & menupos & param

	Case "modify"
		'@@ ���� ����
		SQL =	"Update db_academy.dbo.tbl_eventInfo Set " &_
				"	  evtTitle= '" & evtTitle & "'" &_
				"	, evtCont = '" & evtCont & "'" &_
				"	, evtSdate = '" & evtSdate & "'" &_
				"	, evtEdate = '" & evtEdate & "'" &_
				"	, isComment = '" & isComment & "'" &_
				" Where evtId = " & evtId

		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "Event_view.asp?menupos=" & menupos & "&evtId=" & evtId & param

	Case "delete"
		'@@ ���� ����

		'# ���� ����
		SQL =	"Update db_academy.dbo.tbl_eventInfo Set " &_
				"	  evtTitle= '" & evtTitle & "'" &_
				"	, isusing = 'N'" &_
				" Where evtId = " & evtId

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "Event_list.asp?menupos=" & menupos & param

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

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->