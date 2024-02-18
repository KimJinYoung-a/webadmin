<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim sDate		: sDate = req("sDate", "")
Dim eDate		: eDate = req("eDate", "")

Dim yyyymmdd	: yyyymmdd = req("yyyymmdd", "")
Dim tenUserID	: tenUserID = req("tenUserID", "")
Dim workdays	: workdays = req("workdays", "")

Dim strSql
strSql = " db_datamart.dbo.sp_Ten_Call_Person_WorkDay_Proc ('" & yyyymmdd & "', '" & tenUserID & "', '" & workdays & "')"

db3_dbget.execute (strSql)

response.redirect "operatorSummaryReport.asp?sDate=" & sDate & "&eDate=" & eDate & "&tenUserID=" & tenUserID
%>
<!-- #include virtual="/lib/db/db3close.asp" -->