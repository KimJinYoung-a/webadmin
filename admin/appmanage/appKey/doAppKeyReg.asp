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
Dim sqlStr
Dim idx, mode, vType, vOsType, vAppVersion, vValidationKey, vIsUsing


idx				= requestCheckVar(request("idx"),10)
mode			= requestCheckVar(request("mode"),4)
vType			= requestCheckVar(request("type"),200)
vOsType			= requestCheckVar(request("ostype"),50)
vAppVersion		= requestCheckVar(request("appversion"), 100)
vValidationKey	= requestCheckVar(request("validationkey"),800)
vIsUsing		= requestCheckVar(request("isusing"), 30)

'// ó�� �б�
Select Case mode
	Case "add"
	'�ű� ���
	sqlStr = "Insert into [db_sitemaster].[dbo].tbl_AppValidationCheckKey "
	sqlStr = sqlStr & "(type, ostype, appVersion, validationKey, regdate, lastupdate, adminid, adminname, isusing) "
	sqlStr = sqlStr & "values ("
	sqlStr = sqlStr & "'" & vType & "'"
	sqlStr = sqlStr & ",'" & vOsType & "'"
	sqlStr = sqlStr & ",'" & vAppVersion & "'"
	sqlStr = sqlStr & ",'" & vValidationKey & "'"
	sqlStr = sqlStr & ",getdate()"
	sqlStr = sqlStr & ",getdate()"	
	sqlStr = sqlStr & ",'" & session("ssBctId") & "'"
	sqlStr = sqlStr & ",'" & session("ssBctCname") & "'"
	sqlStr = sqlStr & ",'" & vIsUsing & "'"
	sqlStr = sqlStr & ")"

	dbget.Execute(sqlStr)

	Case "modi"
	'����
	if Not(idx="" or isNull(idx)) then

		sqlStr = "Update [db_sitemaster].[dbo].tbl_AppValidationCheckKey Set "
		sqlStr = sqlStr & "type ='" & vType &"'"
		sqlStr = sqlStr & ", osType='" & vOsType &"'"
		sqlStr = sqlStr & ", appVersion ='" & vAppVersion & "'"
		sqlStr = sqlStr & ", validationKey='" & vValidationKey & "'"
		sqlStr = sqlStr & ", lastupdate = getdate()"
		sqlStr = sqlStr & ", adminid  ='" & session("ssBctId") & "'"
		sqlStr = sqlStr & ", adminName ='" & session("ssBctCname") & "'"
		sqlStr = sqlStr & ", isUsing   ='" & vIsUsing & "'"
		sqlStr = sqlStr & " Where idx=" & idx
		dbget.Execute(sqlStr)

	end if
end Select

''response.write "<script>alert('����Ǿ����ϴ�.');</script>"
response.write "<script>opener.history.go(0); self.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->