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
	dim strSql, mode, i
	Dim dispCate
	dim referer
	referer = request.ServerVariables("HTTP_REFERER")

	mode = request.form("mode")
	strSql = ""

	dispCate		= request.form("catecode_b")

	'// ó�� ��� �б�
	Select Case mode
		Case "add", "modi"
			'ī�װ��Ӽ� �űԵ��
			if Not(dispCate="") then
				strSql = strSql & "Delete from db_item.dbo.tbl_itemAttrib_dispCate Where catecode='" & dispCate & "'"

				for i=1 to request.form("attribDiv").count
					if request.form("attribDiv")(i)<>"" then
						strSql = strSql & "Insert into db_item.dbo.tbl_itemAttrib_dispCate values "
						strSql = strSql & "('" & request.form("attribDiv")(i) & "'"
						strSql = strSql & ",'" & dispCate & "')" & vbCrLf
					end if
				next
			end if

		Case "del"
			'ī�װ��Ӽ� ����
			if Not(dispCate="") then
				strSql = "Delete from db_item.dbo.tbl_itemAttrib_dispCate Where catecode='" & dispCate & "'"
			end if

	end Select

	if strSql<>"" then
		dbget.Execute strSql
	else
		Call Alert_return("������ ������ �����ϴ�.")
		dbget.Close: Response.End
	end if

	response.write "<script>alert('����Ǿ����ϴ�.');</script>"
	response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->