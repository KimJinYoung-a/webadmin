<% option Explicit
Response.CharSet = "euc-kr"
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "Expires","0"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/report/category_reportcls.asp"-->
<%
Dim gubun, sitename, selGpRdsite
gubun = request("gubun")
sitename = request("sitename")
selGpRdsite = request("GpRdsite")
If gubun = "nvshop" Then
	Call RdsiteGubunList(gubun, "GpRdsite", selGpRdsite)
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->