<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	dim strSql, sIdx, sSortNo, sIsOpen, i

	'@���Ĺ�ȣ �ϰ�����
	for i=1 to request.form("chkIdx").count
		sIdx = request.form("chkIdx")(i)
		sSortNo = request.form("sort"&sIdx)
		sIsOpen = request.form("open"&sIdx)
		if sSortNo="" then sSortNo="0"
		if sIsOpen="" then sIsOpen="N"

		strSql = strSql & "Update db_board.dbo.tbl_giftShop_theme Set "
		strSql = strSql & " sortNo='" & sSortNo & "'"
		strSql = strSql & " ,isOpen='" & sIsOpen & "'"
		strSql = strSql & " Where themeIdx='" & sIdx & "';" & vbCrLf
	next

	if strSql<>"" then
		dbget.Execute strSql
	else
		Call Alert_return("������ ������ �����ϴ�.")
		dbget.Close: Response.End
	end if

	dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script>alert('����Ǿ����ϴ�.');</script>"
	response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->