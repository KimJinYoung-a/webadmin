<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� faq����������
' Hieditor : 2009.11.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->
<%
'// ���� ����
dim lp , isusing , msg
dim mode, ntcId, userid ,title, contents ,SQL , retURL
	ntcId		= Request("ntcId")
	mode		= Request("mode")
	title		= Request("title")
	isusing		= Request("isusing")
	contents	= Request("contents")
	userid		= session("ssBctId")

'==============================================================================
'## ���� ����(����) ó��

'Ʈ������ ����
dbget.beginTrans

Select Case mode
	Case "edit"
	
		'//�ű�����
		if ntcId = "" then 
			
			SQL =	"Insert into db_momo.dbo.tbl_Notice " &_
					"	(title, contents, commCd, isusing,userid) values " &_
					"	('" & html2db(title) & "'" &_
					"	,'" & html2db(contents) & "'" &_
					"	,2" &_
					"	,'Y'" &_
					"	,'" & userid & "')"
			
			'response.write SQL &"<br>"		
			dbget.Execute(SQL)
	
			'��� �޽���
			msg = "�����Ͽ����ϴ�."
		
		'//����
		else
			
			SQL =	"Update db_momo.dbo.tbl_Notice Set " &_
					"	  title= '" & html2db(title) & "'" &_
					"	, contents = '" & html2db(contents) & "'" &_
					"	, isusing = '" & isusing & "'" &_
					" Where ntcId = " & ntcId
			
			'response.write SQL &"<br>"
			dbget.Execute(SQL)
	
			msg = "�����Ͽ����ϴ�."
		end if
		
	Case "delete"
		'@@ ���� ����
		SQL =	"Update db_momo.dbo.tbl_Notice Set " &_
				" isusing = 'N'" &_
				" Where ntcId = " & ntcId

		'response.write SQL &"<br>"
		dbget.Execute(SQL)
		
		msg = "�����Ͽ����ϴ�."

End Select


'�����˻� �� �ݿ�
If Err.Number = 0 Then   
	dbget.CommitTrans				'Ŀ��(����)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	location.href='faq_list.asp';" &_
					"</script>"

Else
    dbget.RollBackTrans				'�ѹ�(�����߻���)

	response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

End If
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->