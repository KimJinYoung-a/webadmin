<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/kaffa/kaffaCls.asp"-->
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<%
	Dim vQuery, vTenCDL, vTenCDM, vTenCDS, vCate1, vCate2, vCate3

	If Request("tencode") = "" Then
		dbget.close()
		Response.End
	End If

	vTenCDL = Left(Request("tencode"),3)
	vTenCDM = Mid(Request("tencode"),4,3)
	vTenCDS = Right(Request("tencode"),3)
	vCate1 = Request("cate1")
	If vCate1 = "x" Then
		vCate1 = "0"
	End If
	vCate2 = Request("cate2")
	If vCate2 = "x" Then
		vCate2 = "0"
	End If
	vCate3 = Request("cate3")
	If vCate3 = "x" Then
		vCate3 = "0"
	End If

	vQuery = "IF EXISTS(select tencdl from [db_item].dbo.tbl_kaffa_category_mapping where tencdl = '" & vTenCDL & "' and tencdm = '" & vTenCDM & "' and tencds = '" & vTenCDS & "') " & _
			 "	BEGIN " & _
			 "		UPDATE [db_item].dbo.tbl_kaffa_category_mapping " & _
			 "			SET kaffacate1 = '" & vCate1 & "', kaffacate2 = '" & vCate2 & "', kaffacate3 = '" & vCate3 & "' " & _
			 "		WHERE tencdl = '" & vTenCDL & "' and tencdm = '" & vTenCDM & "' and tencds = '" & vTenCDS & "' " & _
			 "	END " & _
			 "ELSE " & _
			 "	BEGIN " & _
			 "		INSERT INTO [db_item].dbo.tbl_kaffa_category_mapping(tencdl, tencdm, tencds, kaffacate1, kaffacate2, kaffacate3) " & _
			 "		VALUES('" & vTenCDL & "', '" & vTenCDM & "', '" & vTenCDS & "', '" & vCate1 & "', '" & vCate2 & "', '" & vCate3 & "') " & _
			 "END "

	dbget.execute vQuery
%>
<script>
parent.$("#result-<%=Request("tencode")%>").empty().append("<br><font color=blue>[ÀúÀå]</font>");
parent.frm1.tencode.value = "";
parent.frm1.cate1.value = "0";
parent.frm1.cate2.value = "0";
parent.frm1.cate3.value = "0";
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->