<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
	Dim vQuery, vCode, vRealCode, vMidCode, vMidCodeName, vOrderNo, vIsUsing, vAdminID
	vCode = Trim(Request("code"))
	vMidCode = Trim(Request("midcode"))
	vMidCodeName = Trim(Request("midcodename"))
	vOrderNo = Trim(Request("orderno"))
	vIsUsing = Request("isusing")
	vAdminID = session("ssBctId")
	
	If vMidCode <> "" Then
		vQuery = "UPDATE [db_giftplus].[dbo].[tbl_stylepick_cate_cd3] SET catename = '" & vMidCodeName & "', isusing = '" & vIsUsing & "', orderno = '" & vOrderNo & "', lastadminid = '" & vAdminID & "' WHERE cd3 = '" & vMidCode & "'"
		dbget.execute vQuery
	Else
		vQuery = "SELECT TOP 1 cd3 From [db_giftplus].[dbo].[tbl_stylepick_cate_cd3] WHERE Left(cd3, 1) = '" & vCode & "' ORDER BY cd3 DESC"
		rsget.Open vQuery, dbget, 1
		If rsget.Eof Then
			vRealCode = vCode & "01"
		Else
			vRealCode = CStr(CInt(rsget("cd3")) + 1)
		End IF
		rsget.Close
		
		vQuery = "INSERT INTO [db_giftplus].[dbo].[tbl_stylepick_cate_cd3](cd3,catename,isusing,orderno,lastadminid) VALUES('" & vRealCode & "', '" & vMidCodeName & "', '" & vIsUsing & "', '" & vOrderNo & "', '" & vAdminID & "')"
		dbget.execute vQuery
	End IF
	
%>

<script language="javascript">
document.location.href = "stylepick_midcate.asp?code=<%=vCode%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->