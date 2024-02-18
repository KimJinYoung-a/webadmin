<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim vQuery, vOrderserial, v6MonthAgo, vOrderTable, vYN_Gubun
	vOrderserial 	= Request("forderserial")
	v6MonthAgo		= requestCheckvar(Request("o6monthago"),1)
	vYN_Gubun		= requestCheckvar(Request("yn_gubun"),1)

	If v6MonthAgo = "o" Then
		vOrderTable = "[db_log].[dbo].[tbl_old_order_master_2003]"
	Else
		vOrderTable = "[db_order].[dbo].[tbl_order_master]"
	End IF

	If vOrderserial <> "" Then
		vQuery = "UPDATE " & vOrderTable & " Set userDisplayYn = '" & vYN_Gubun & "' WHERE orderserial IN(" & vOrderserial & ")"
		dbget.Execute vQuery
	End IF
%>
<script language="javascript">
alert("처리되었습니다.");
parent.document.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->