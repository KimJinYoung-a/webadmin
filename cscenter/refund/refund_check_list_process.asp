<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim mode
dim asidList
dim AssignedRow


mode		= trim(request("mode"))
asidList	= trim(request("asidList"))


dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr


Select Case mode
	Case "ipjumRefund"
		'' ���޸� ����Ȯ�� �� ȯ��
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = '���޸� ����Ȯ�� �� ȯ��' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case "ipjumDiffRefund"
		'' ���Ա� ����ȯ��
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = '���Ա� ����ȯ��' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case "prdDiffRefund"
		'' ��ǰ��� ����ȯ��
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = '��ǰ��� ����ȯ��' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case "csDelivRefund"
		'' CS���� - ��ۺ�ȯ��
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = 'CS���� - ������ ȯ��(��ۺ�)' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case "upcheJungsanRefund"
		'' ��ü���� �� ��ȯ��
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = '��ü���� �� ��ȯ��' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case Else
		''
End Select

%>
<script language='javascript'>
alert('���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
