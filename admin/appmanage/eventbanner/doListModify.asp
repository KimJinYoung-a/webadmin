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
	dim strSql, sIdx, sSortNo, sIsUsing, i

	'@정렬번호 일괄저장
	for i=1 to request.form("chkIdx").count
		sIdx = request.form("chkIdx")(i)
		sSortNo = request.form("sort"&sIdx)
		sIsUsing = request.form("use"&sIdx)
		if sSortNo="" then sSortNo="0"
		if sIsUsing="" then sIsUsing="N"

		strSql = strSql & "Update [db_contents].[dbo].tbl_app_eventBanner Set "
		strSql = strSql & " sortNo='" & sSortNo & "'"
		strSql = strSql & " ,isUsing='" & sIsUsing & "'"
	    strSql = strSql & " ,lastUpdateUser='" & session("ssBctId") & "'"
	    strSql = strSql & " ,lastUpdate=getdate()"
		strSql = strSql & " Where idx='" & sIdx & "';" & vbCrLf
	next

	if strSql<>"" then
		dbget.Execute strSql
	else
		Call Alert_return("저장할 내용이 없습니다.")
		dbget.Close: Response.End
	end if

	dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->