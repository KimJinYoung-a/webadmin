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
	dim strSql, mode, sIdx, sSortNo, sIsUsing, i , sStandardprice

	mode = request.form("mode")

	If mode = "new" Then '// �����ͻ����� �ʱ� �� �ҷ�����
	
		strSql = " DELETE FROM db_sitemaster.dbo.tbl_mobile_todaymore_category; "			
		strSql = strSql & " insert into db_sitemaster.dbo.tbl_mobile_todaymore_category "
		strSql = strSql & " (dispcate , catename , sorting , standardprice) "
		strSql = strSql & " select catecode , catename , (ROW_NUMBER() OVER(ORDER BY sortNo ASC)) as sorting , 0 as standardprice "
		strSql = strSql & " from db_item.dbo.tbl_display_cate "
		strSql = strSql & "	where depth = 1 and useyn = 'Y' and catecode not in('123') "

		dbget.Execute strSql

	ElseIf mode = "edit" then
		'@���Ĺ�ȣ �ϰ�����
		for i=1 to request.form("chkIdx").count
			sIdx = request.form("chkIdx")(i)
			sSortNo = request.form("sort"&sIdx)
			sStandardprice = request.form("standardprice"&sIdx)

			if sSortNo="" then sSortNo="1"

			strSql = strSql & "Update db_sitemaster.dbo.tbl_mobile_todaymore_category Set "
			strSql = strSql & " sorting='" & sSortNo & "'"
			strSql = strSql & " ,standardprice='" & Trim(sStandardprice) & "'"
			strSql = strSql & " Where dispcate='" & sIdx & "';" & vbCrLf
		Next

		if strSql<>"" then
			dbget.Execute strSql
		else
			Call Alert_return("������ ������ �����ϴ�.")
			dbget.Close: Response.End
		end If

	End If 

	dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->