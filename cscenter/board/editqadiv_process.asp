<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/noreplyboardcls.asp" -->
<%
dim referer
dim id,qadiv,itemid

id = request.form("id")
qadiv = request.form("qadiv")
itemid = request.form("itemid")

if (itemid = "") then
    itemid = "null"
end if


referer = request.ServerVariables("HTTP_REFERER")

 dim sql

	sql = "update [db_cs].[dbo].tbl_myqna " + VbCRlf
	sql = sql + " set qadiv = '" + qadiv + "'" + VbCRlf
	sql = sql + " , itemid = " + CStr(itemid) + "" + VbCRlf
	sql = sql + " where (id = " + id + ") "

	rsget.Open sql, dbget, 1


	response.write "<script>alert('수정되었습니다.')</script>"
	response.write "<script>location.replace('" + referer + "')</script>"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->