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

dim sqlStr
dim menupos
dim mode
dim extTime

menupos = requestCheckVar(request("menupos"), 32)
mode = requestCheckVar(request("mode"), 32)
extTime = requestCheckVar(request("extTime"), 32)

if (mode = "RefreshData") then

	application(extTime) = ""

	response.write	"<script language='javascript'>" &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
