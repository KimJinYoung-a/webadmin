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
		Response.Write "<script>alert('퇴근시간을 입력하세요.');history.back();</script>"
		dbget.close()
		Response.End
	End IF
	
	If InStr(vStartWork,":") = 0 AND vEndWork <> "" Then
		Response.Write "<script>alert('출근시간이 시간형식이 아닌 경우는 퇴근시간을 비워두세요.');history.back();</script>"
		dbget.close()
		Response.End
	End IF
	
	If vAction = "" OR (vAction <> "update" AND vAction <> "insert") Then
		Response.Write "<script>alert('잘못된 경로입니다.');history.back();</script>"
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
			Response.Write "<script>alert('똑같은 근무코드가 존재합니다.');history.back();</script>"
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
	alert("저장되었습니다.");
	location.href = "/common/offshop/staff/offshop_employee_workcode.asp";
//-->
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->