<%@ language = vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode
dim barcodeArr
dim strSql

mode = requestCheckVar(request("mode"), 32)
barcodeArr = requestCheckVar(request("barcodeArr"), 4096)


Select Case mode
	Case "setuseyn"
		barcodeArr = Mid(barcodeArr, 2, Len(barcodeArr) - 1)

		strSql = " update i "
		strSql = strSql + " set i.isusing = 'N' "
		strSql = strSql + " from "
		strSql = strSql + " 	[db_item].[dbo].tbl_item i "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "
		strSql = strSql + " 	and i.isusing = 'Y' "
		strSql = strSql + " 	and i.sellyn <> 'Y' "
		strSql = strSql + " 	and i.itemid in (" + CStr(barcodeArr) + ") "
		''response.write strSql
		dbget.execute strSql

		response.write "<script>alert('저장되었습니다'); history.back();</script>"
		dbget.Close
		response.end
	Case Else
		response.write "ERR"
		response.end
End Select

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<script language="javascript">
	alert('저장되었습니다');
	// opener.opener.location.reload();
	// opener.frm.drawitemid.value = '';
	self.close();
</script>
