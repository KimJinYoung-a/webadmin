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
	dim strSql, mode, sIdx, sSortNo, sIsUsing, i , sItemname
	Dim mourl , appurl , label , ldv

	mode = request.form("mode")

	'@���Ĺ�ȣ �ϰ�����
	for i=1 to request.form("chkIdx").count
		sIdx = request.form("chkIdx")(i)
		sSortNo = request.form("sort"&sIdx)
		sIsUsing = request.form("use"&sIdx)
		sItemname = request.form("itemnamewish"&sIdx)
		label = request.form("label"&sIdx)
		ldv = request.form("ldv"&sIdx)
		if sSortNo="" then sSortNo="0"
		if sIsUsing="" then sIsUsing="N"

		Select Case mode
			Case "main"
				strSql = strSql & "Update [db_sitemaster].[dbo].tbl_pc_main_wishbest_item Set "
				strSql = strSql & " sortnum='" & sSortNo & "'"
				strSql = strSql & " ,isusing='" & sIsUsing & "'"		'����Ʈ ����: ��뿩�� > ������� ����
				strSql = strSql & " Where subidx='" & sIdx & "';" & vbCrLf
			Case "sub"
				strSql = strSql & "Update [db_sitemaster].[dbo].tbl_pc_main_wishbest_item Set "
				strSql = strSql & " sortnum='" & sSortNo & "'"
				strSql = strSql & " ,isusing='" & sIsUsing & "'"
				strSql = strSql & " ,wishitemname='" & sItemname & "'"
				strSql = strSql & " Where subidx='" & sIdx & "';" & vbCrLf
		end Select
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