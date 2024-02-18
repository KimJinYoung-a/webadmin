<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim userid
dim shiftid
dim groupid

userid  = session("ssBctId")
groupid = session("ssGroupid")
shiftid = requestCheckVar(request("shiftid"),32)

dim ref
ref = request.ServerVariables("HTTP_REFERER")

dim IsValidShiftID
dim curr_Cuserdiv, pre_Cuserdiv


''로그인
dim sqlStr
sqlStr = "select top 1 p.id,p.company_name,p.tel,p.fax,p.url,p.email,p.bigo,p.userdiv,p.groupid, c.userdiv as cuserdiv " + vbCrlf
sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
sqlStr = sqlStr + "     Join db_user.dbo.tbl_user_c c on p.id=c.userid"
sqlStr = sqlStr + " where p.id = '" + shiftid + "'" + vbCrlf
sqlStr = sqlStr + " and p.groupid='" + groupid + "'"  + vbCrlf
sqlStr = sqlStr + " and p.isusing='Y'"
rsget.Open sqlStr,dbget,1


if  not rsget.EOF  then
    session("ssBctId") = rsget("id")
    session("ssBctDiv") = rsget("userdiv")
    session("ssBctBigo") = rsget("bigo")
    session("ssBctCname") = db2html(rsget("company_name"))
	session("ssBctEmail") = db2html(rsget("email"))
	session("ssGroupid") = rsget("groupid")

	response.Cookies("partner").domain = "10x10.co.kr"
    response.Cookies("partner")("userid") = session("ssBctId")
    response.Cookies("partner")("userdiv") = session("ssBctDiv")

    curr_Cuserdiv = rsget("cuserdiv")

    IsValidShiftID = true
else
	IsValidShiftID = false
end if
rsget.close

''무조건 Main Notics로 보냄 20080516 서동석
ref = getSCMURL&"/designer/notics/notics.asp?menupos=52"

'''강사에서 <=> 브랜드 로 스위칭 한경우=====================================
sqlStr = "select top 1 userdiv as precuserdiv " + vbCrlf
sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c where userid = '" + userid + "'" + vbCrlf
rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    pre_Cuserdiv = rsget("precuserdiv")
end if
rsget.close

dim topchange : topchange = false
if (curr_Cuserdiv<>pre_Cuserdiv) then
    topchange = true
    ref = "/designer/index.asp"
end if

if (curr_Cuserdiv="14") then
    session("isAgreeReq")="" ''계약서 관련 세션 (2016/08/11)
    topchange = true
    ref = "/lectureadmin/index.asp"
end if
''==========================================================================
%>

<% if Not IsValidShiftID then %>
<script language='javascript'>alert('Not Valid ID');</script>
<%
session.abandon
%>
<% end if %>

<% if topchange then %>
<script language='javascript'>top.location.replace('<%= ref %>')</script>
<% else %>
<script language='javascript'>location.replace('<%= ref %>')</script>
<% end if %>


<!-- #include virtual="/lib/db/dbclose.asp" -->