<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/offshop/staff/offshop_employee_managementCls.asp"-->

<%
	Dim i, vQuery, vAction, vWorkCode, vStartWork, vEndWork, vIsOK
	vAction = Request("action")
	vWorkCode = Request("workcode")
	vStartWork = Request("startwork")
	vEndWork = Request("endwork")
	
	If InStr(vStartWork,":") <> 0 AND vEndWork = "" Then
		Response.Write "<script>alert('��ٽð��� �Է��ϼ���.');history.back();</script>"
		dbget.close()
		Response.End
	End IF
	
	If InStr(vStartWork,":") = 0 AND vEndWork <> "" Then
		Response.Write "<script>alert('��ٽð��� �ð������� �ƴ� ���� ��ٽð��� ����μ���.');history.back();</script>"
		dbget.close()
		Response.End
	End IF
	
	If vAction = "" OR (vAction <> "update" AND vAction <> "insert") Then
		Response.Write "<script>alert('�߸��� ����Դϴ�.');history.back();</script>"
		dbget.close()
		Response.End
	End IF
	
	If InStr(vStartWork,":") <> 0 Then
		vStartWork = (CInt(Split(vStartWork,":")(0))*60) + CInt(Split(vStartWork,":")(1))
		vEndWork = (CInt(Split(vEndWork,":")(0))*60) + CInt(Split(vEndWork,":")(1))
	End If
	
	If vAction = "insert" Then
		vWorkCode = UCase(vWorkCode)
		
		vQuery = "IF NOT EXISTS(SELECT workcode FROM [db_partner].[dbo].[tbl_offshop_employee_workcode] WHERE workcode = '" & vWorkCode & "') " & vbCrLf
		vQuery = vQuery & "	BEGIN " & vbCrLf
		vQuery = vQuery & "		SELECT 'o' " & vbCrLf
		vQuery = vQuery & "	END " & vbCrLf
		vQuery = vQuery & "ELSE " & vbCrLf
		vQuery = vQuery & "	BEGIN " & vbCrLf
		vQuery = vQuery & "		SELECT 'x' " & vbCrLf
		vQuery = vQuery & "	END "
		rsget.open vQuery,dbget,1

		vIsOK = rsget(0)
		rsget.close()

		If vIsOK = "x" Then
			Response.Write "<script>alert('�Ȱ��� �ٹ��ڵ尡 �����մϴ�.');history.back();</script>"
			dbget.close()
			Response.End
		Else
			vQuery = "INSERT INTO [db_partner].[dbo].[tbl_offshop_employee_workcode](workcode, startwork, endwork) VALUES('" & vWorkCode & "', '" & vStartWork & "', '" & vEndWork & "') " & vbCrLf
			dbget.Execute vQuery
		End IF
	ElseIf vAction = "update" Then
		vQuery = "UPDATE [db_partner].[dbo].[tbl_offshop_employee_workcode] SET startwork = '" & vStartWork & "', endwork = '" & vEndWork & "' WHERE workcode = '" & vWorkCode & "' " & vbCrLf
		dbget.Execute vQuery
	End IF
%>

<script type="text/javascript">
<!--
	alert("����Ǿ����ϴ�.");
	location.href = "/common/offshop/staff/offshop_employee_workcode.asp";
//-->
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->