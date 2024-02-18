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
dim mode
dim mdTime

mode = requestCheckVar(request("mode"), 32)
mdTime = requestCheckVar(request("mdTime"), 32)

if (mode = "RefreshData") then

	application(mdTime) = DateAdd("d", -1, now)

	response.write	"<script language='javascript'>" &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
