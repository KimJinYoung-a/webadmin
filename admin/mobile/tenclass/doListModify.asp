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

	mode = request.form("mode")

	'@���Ĺ�ȣ �ϰ�����
	for i=1 to request.form("chkIdx").count
		sIdx = request.form("chkIdx")(i)
		sSortNo = request.form("sort"&sIdx)
		sIsUsing = request.form("use"&sIdx)
		sItemname = request.form("itemname"&sIdx)
		if sSortNo="" then sSortNo="0"
		if sIsUsing="" then sIsUsing=1

		Select Case mode
			Case "main"
				strSql = strSql & "UPDATE db_sitemaster.dbo.tbl_mobile_class_items SET "
				strSql = strSql & " sortno='" & sSortNo & "'"
				strSql = strSql & " ,isusing=" & sIsUsing & ""		'����Ʈ ����: ��뿩�� > ������� ����
				strSql = strSql & " ,itemname='" & html2db(sItemname) & "'"
				strSql = strSql & " WHERE didx='" & sIdx & "';" & vbCrLf
			Case "sub"
				strSql = strSql & "UPDATE db_sitemaster.dbo.tbl_mobile_class_items SET "
				strSql = strSql & " sortno='" & sSortNo & "'"
				strSql = strSql & " ,isusing=" & sIsUsing & ""
				strSql = strSql & " ,itemname='" & html2db(sItemname) & "'"
				strSql = strSql & " WHERE didx='" & sIdx & "';" & vbCrLf
			Case "itemdel"
				strSql = strSql & "DELETE FROM db_sitemaster.dbo.tbl_mobile_class_items "
				strSql = strSql & " WHERE didx='" & sIdx & "';" & vbCrLf
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
	If mode = "itemdel" Then
		response.write "<script>alert('�����Ǿ����ϴ�.');</script>"
	Else
		response.write "<script>alert('����Ǿ����ϴ�.');</script>"
	End If
	response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->