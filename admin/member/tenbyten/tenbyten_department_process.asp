<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim SQL
dim menupos
dim mode
dim pid, cid, departmentName, dispOrderNo, useYN

mode = requestCheckVar(request("mode"), 32)

pid = requestCheckVar(request("pid"), 32)
cid = requestCheckVar(request("cid"), 32)
departmentName = requestCheckVar(request("departmentName"), 64)
dispOrderNo = requestCheckVar(request("dispOrderNo"), 32)
useYN = requestCheckVar(request("useYN"), 32)


if (mode = "depart_modi") then
	if (pid <> "") then
		SQL = " insert into db_partner.dbo.tbl_user_department(pid, departmentName, dispOrderNo) "
		SQL = SQL + " values(" + CStr(pid) + ", '" + html2db(CStr(departmentName)) + "', " + CStr(dispOrderNo) + ") "
		rsget.Open SQL, dbget, 1
	elseif (cid <> "") then
		'
		SQL = " update db_partner.dbo.tbl_user_department "
		SQL = SQL + " set departmentName = '" + html2db(CStr(departmentName)) + "', dispOrderNo = " + CStr(dispOrderNo) + ", useYN = '" + CStr(useYN) + "', lastupdate = getdate() "
		SQL = SQL + " where cid = " + CStr(cid) + " "
		rsget.Open SQL, dbget, 1
	end if

	response.write	"<script language='javascript'>" &_
					"	alert('저장되었습니다.'); opener.focus(); opener.location.reload(); window.close(); " &_
					"</script>"

else

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
