<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  언어정보 불러오기
' History : 2016.06.01 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbhelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<%
dim sqlstr, i, loginsite, countrylangcd, tmp_str
	loginsite = requestcheckvar(request("loginsite"),32)
	countrylangcd = requestcheckvar(request("countrylangcd"),32)

if loginsite="" then
	dbget.close() : response.end
end if

sqlstr = " select"
sqlstr = sqlstr & " countrylangcd"
sqlstr = sqlstr & " from db_item.dbo.tbl_exchangeRate"
sqlstr = sqlstr & " where 1=1"

if loginsite<>"" and loginsite<>"SCM" then
	sqlstr = sqlstr & " and sitename='"& loginsite &"'"
end if

sqlstr = sqlstr & " group by countrylangcd"
sqlstr = sqlstr & " order by countrylangcd asc"

'response.write sqlstr &"<Br>"
rsget.Open sqlstr,dbget,1
%>
<select class="select" name="tmpcountrylangcd" onchange='selectedcountrylangcd(this.value);'>
	<option value='' <%if countrylangcd="" then response.write " selected"%>>CHOICE</option>
<%
if not rsget.EOF then
	rsget.Movefirst

	do until rsget.EOF
	'if Lcase(countrylangcd) = Lcase(rsget("countrylangcd")) then
		'tmp_str = " selected"	' 선택시키지 말것.
	'end if
	response.write("<option value='"&rsget("countrylangcd")&"' "&tmp_str&">"&rsget("countrylangcd")&"</option>")
	tmp_str = ""
	rsget.MoveNext
	loop
end if
rsget.close

	response.write("</select>")
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
