<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
response.write "������"
response.end

'// ���� ����
dim msg, lp, menupos
dim mode
dim comm_group, comm_cd, comm_name, comm_isDel, comm_color, sortno
dim SQL
dim page, groupCd, searchKey, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= Request("menupos")
mode		= Request("mode")
comm_group	= Request("comm_group")
comm_cd		= Request("comm_cd")
comm_name	= html2db(Request("comm_name"))
comm_isDel	= Request("comm_isDel")
page		= Request("page")
groupCd		= Request("groupCd")
searchKey	= Request("searchKey")
searchString = Request("searchString")
comm_color   = Request("menucolor")
sortno		= Request("sortno")


param = "&page=" & page & "&groupCd=" & groupCd & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����
if sortno="" then sortno=0


'==============================================================================
'## ���� ����(����) ó��

'Ʈ������ ����
dbget.beginTrans

Select Case mode
	Case "write"
		'@@ �űԵ��
		'�ߺ��˻�
		SQL = "Select count(comm_cd) as cnt From db_cs.dbo.tbl_cs_comm_code where comm_cd='" & comm_cd & "'"
		rsget.Open sql, dbget, 1
			if rsget("cnt")>0 then
				response.write	"<script language='javascript'>" &_
								"	alert('�ߺ��� �ڵ带 �Է��Ͽ����ϴ�.');" &_
								"	history.back();" &_
								"</script>"
				dbget.close()	:	response.End
			end if
		rsget.close

		'����
		SQL =	"Insert into db_cs.dbo.tbl_cs_comm_code (comm_group, comm_cd, comm_name, comm_color, sortno) " &_
				"	Values " &_
				"	( '" & comm_group & "'" &_
				"	, '" & comm_cd & "'" &_
				"	, '" & comm_name & "'" &_
				"	, '" & comm_color & "'" &_
				"	, '" & sortno & "') "

		dbget.Execute(SQL)

		msg = "�ű� ����Ͽ����ϴ�."

		'���ư� ������
		retURL = "commCd_List.asp?menupos=" & menupos & param

	Case "modify"
		'@@ ����ó��

		SQL =	"Update db_cs.dbo.tbl_cs_comm_code Set " &_
				"	comm_name = '" & comm_name & "'" &_
				"	,comm_isDel = '" & comm_isDel & "'" &_
				"	,comm_color = '" & comm_color & "'" &_
				"	,sortno = '" & sortno & "'" &_
				" Where comm_cd = '" & comm_cd & "'"
		dbget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "commCd_List.asp?menupos=" & menupos & param

End Select


'�����˻� �� �ݿ�
If Err.Number = 0 Then   
	dbget.CommitTrans				'Ŀ��(����)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbget.RollBackTrans				'�ѹ�(�����߻���)

	response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->