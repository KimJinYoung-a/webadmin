<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim mode
dim companyid, userid, username, useyn
dim sqlStr

mode        = requestCheckVar(request("mode"),32)
userid   	= requestCheckVar(request("userid"),32)
companyid   = requestCheckVar(request("companyid"),32)
username 	= html2db(requestCheckVar(request("username"),32))
useyn       = requestCheckVar(request("useyn"),32)

select case mode
	case "modi"
		sqlStr = ""
		sqlStr = sqlStr & " update [db_threepl].[dbo].[tbl_user] "
		sqlStr = sqlStr & " set lastupdt = getdate(), "
		sqlStr = sqlStr & " username = '" & username & "', "
		sqlStr = sqlStr & " companyid = '" & companyid & "', "
		sqlStr = sqlStr & " useyn = '" & useyn & "' "
		sqlStr = sqlStr & " where userid = '" & userid & "'"
		dbget_TPL.Execute sqlStr

		response.write "<script>alert('수정 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case "ins"
		sqlStr = ""
		sqlStr = sqlStr & " insert into [db_threepl].[dbo].[tbl_user](userid, companyid, username, useyn, userpass, isadmin) "
		sqlStr = sqlStr & " values('" & userid & "', '" & companyid & "', '" & username & "', '" & useyn & "', 'not used', 'N')"
		dbget_TPL.Execute sqlStr

		'// userid, companyid, defaultlocationid, userpass, username, isadmin, userdiv, useyn, lastupdt, regdate

		response.write "<script>alert('저장 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case else
		response.write "에러"
end select

%>
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
