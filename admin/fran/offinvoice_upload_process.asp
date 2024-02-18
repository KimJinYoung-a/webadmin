<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%

'==============================================================================
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim masteridx, filename, realfilename,ino



'==============================================================================
masteridx		= request("masteridx")
filename		= request("filename")
realfilename	= html2db(request("realfilename"))
ino				= request("ino")


'==============================================================================
dim sqlStr, i, iid


if ino = "2" then
	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_offline_invoice_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	exportdeclarefilename2 = '" + CStr(filename) + "' "
	sqlStr = sqlStr + " 	, realfilename2 = '" + CStr(realfilename) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	idx = " + CStr(masteridx) + "	 "
	dbget.Execute sqlStr
elseif ino = "3" then
	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_offline_invoice_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	exportdeclarefilename3 = '" + CStr(filename) + "' "
	sqlStr = sqlStr + " 	, realfilename3 = '" + CStr(realfilename) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	idx = " + CStr(masteridx) + "	 "
	dbget.Execute sqlStr
else		
	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_offline_invoice_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	exportdeclarefilename = '" + CStr(filename) + "' "
	sqlStr = sqlStr + " 	, realfilename = '" + CStr(realfilename) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	idx = " + CStr(masteridx) + "	 "
	dbget.Execute sqlStr
end if
%>

<script language="javascript">
	alert('저장 되었습니다.');
	opener.focus();
	window.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->