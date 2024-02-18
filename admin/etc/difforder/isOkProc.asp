<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim chk, idx, sqlStr, mode, getOrderDate
Dim cksel : cksel = request("cksel")
chk	= request("chk")
idx = request("idx")
mode = request("mode")
getOrderDate = request("getOrderDate")

If mode = "CHK" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE [db_etcmall].[dbo].[tbl_outmall_diffOrder] SET " & vbcrlf
	If chk = "Y" Then
		sqlStr = sqlStr & " isOk = 'Y' "	 & vbcrlf
	Else
		sqlStr = sqlStr & " isOk = null " & vbcrlf
	End If
	sqlStr = sqlStr & " WHERE idx = '"&idx&"'  "
	dbget.Execute sqlStr
ElseIf mode = "getOrder" Then
	sqlStr = ""
	If getOrderDate = "" Then
		sqlStr = sqlStr & " exec [db_etcmall].[dbo].[sp_Ten_outmall_DiffOrder]" 
	Else
		sqlStr = sqlStr & " exec [db_etcmall].[dbo].[sp_Ten_outmall_DiffOrder] '"&getOrderDate&"' "	
	End If
	dbget.Execute sqlStr, 1
ElseIf mode = "CHK2" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_etcmall.[dbo].[tbl_outmall_margin_check] SET " & vbcrlf
	If chk = "Y" Then
		sqlStr = sqlStr & " isOk = 'Y' "	 & vbcrlf
	Else
		sqlStr = sqlStr & " isOk = 'N' " & vbcrlf
	End If
	sqlStr = sqlStr & " WHERE idx = '"&idx&"'  "
	dbget.Execute sqlStr
ElseIf mode = "ALL" Then
	cksel = Trim(cksel)
	If Right(cksel,1) = "," Then cksel = Left(cksel, Len(cksel) - 1)

	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_etcmall.[dbo].[tbl_outmall_margin_check] SET " & vbcrlf
	sqlStr = sqlStr & " isOk = 'Y' "	 & vbcrlf
	sqlStr = sqlStr & " WHERE idx in (" & cksel & ")" & VbCrlf
	dbget.Execute sqlStr
End If

If (session("ssBctID")<>"kjy8517") Then
	response.write "<script>parent.location.reload();</script>"
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->