<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : [ON]��ǰ����>>�ǸŴ���ǰLIST ó��������
' History : �̻� ����
'			2023.10.4 �ѿ�� ����(�����α��߰�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim strSql, i, vChangeContents, vSCMChangeSQL, mode, itemidArr, refer, menupos
	mode     		= requestCheckVar(Request("mode"), 32)
	menupos     		= requestCheckVar(getNumeric(Request("menupos")), 10)
	itemidArr     	= requestCheckVar(Request("itemidArr"), 4000)

refer = request.ServerVariables("HTTP_REFERER")

Select Case mode
	Case "modisellY"
		strSql = " update [db_item].[dbo].tbl_item "
		strSql = strSql + " set sellyn = 'Y', lastupdate = getdate() "
		strSql = strSql + " where itemid in (" & itemidArr & ") and sellyn <> 'Y' "

		'response.write strSql & "<Br>"
		dbget.execute strSql

		vChangeContents = "- �Ǹſ��� : sellyn = Y" & vbCrLf

		'### ���� �α� ����(item)
		vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log] (userid, gubun, pk_idx, menupos, contents, refip)"
		vSCMChangeSQL = vSCMChangeSQL & "   select"
		vSCMChangeSQL = vSCMChangeSQL & "   '" & session("ssBctId") & "', 'item', i.itemid, '" & menupos & "', '" & vChangeContents & "'"
		vSCMChangeSQL = vSCMChangeSQL & "   , '" & Request.ServerVariables("REMOTE_ADDR") & "'"
		vSCMChangeSQL = vSCMChangeSQL & "   from db_item.dbo.tbl_item i with (nolock)"
		vSCMChangeSQL = vSCMChangeSQL & "   where itemid in (" & itemidArr & ") and sellyn <> 'Y' "

		'response.write vSCMChangeSQL & "<Br>"
		dbget.execute vSCMChangeSQL

		response.write	"<script type='text/javascript'>"
		response.write	"	alert('ó���Ǿ����ϴ�.');"
		response.write	"	location.replace('" + CStr(refer) + "');"
		response.write	"</script>"
	Case Else
		response.write("�߸��� ��� �Դϴ�.")
End Select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
