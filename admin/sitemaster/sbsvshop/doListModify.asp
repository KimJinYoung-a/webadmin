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
	dim strSql, mode, sIdx, sSortNo, sIsUsing, i , sItemname , sImageno

	mode = request.form("mode")

	'@정렬번호 일괄저장
	for i=1 to request.form("chkIdx").count
		sIdx = request.form("chkIdx")(i)
		sSortNo = request.form("sort"&sIdx)
		sIsUsing = request.form("use"&sIdx)
		sItemname = request.form("itemname"&sIdx)
		sImageno = request.form("imageno")
		if sSortNo="" then sSortNo="0"
		if sIsUsing="" then sIsUsing="N"

		Select Case mode
			Case "main"
				strSql = strSql & "Update [db_sitemaster].[dbo].tbl_sbsvshop_items Set "
				strSql = strSql & " sortnum='" & sSortNo & "'"
				strSql = strSql & " ,isusing='" & sIsUsing & "'"		'사이트 메인: 사용여부 > 선노출로 변경
				strSql = strSql & " ,itemname='" & sItemname & "'"
				strSql = strSql & " Where subidx='" & sIdx & "';" & vbCrLf
			Case "sub"
				strSql = strSql & "Update [db_sitemaster].[dbo].tbl_sbsvshop_items Set "
				strSql = strSql & " sortnum='" & sSortNo & "'"
				strSql = strSql & " ,isusing='" & sIsUsing & "'"
				strSql = strSql & " ,itemname='" & sItemname & "'"
				strSql = strSql & " Where subidx='" & sIdx & "';" & vbCrLf
			Case "itemdel"
				strSql = strSql & "delete from [db_sitemaster].[dbo].tbl_sbsvshop_items "
				strSql = strSql & " Where subidx='" & sIdx & "';" & vbCrLf
			Case "imagedel"
				strSql = strSql & "update [db_sitemaster].[dbo].tbl_sbsvshop_list set subimage"& sImageno &" = ''"
				strSql = strSql & " Where listidx = "& sIdx &";" & vbCrLf
		end Select
	next

	if strSql<>"" then
		dbget.Execute strSql
	else
		Call Alert_return("저장할 내용이 없습니다.")
		dbget.Close: Response.End
	end if

	dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	If mode = "itemdel" Then
		response.write "<script>alert('삭제되었습니다.');</script>"
	Else
		response.write "<script>alert('저장되었습니다.');</script>"
	End If 
	response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->