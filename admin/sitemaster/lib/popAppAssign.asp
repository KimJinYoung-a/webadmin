<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim allrefreshVal, idx
allrefreshVal = request("allrefresh")
idx = requestCheckvar(request("idx"),10)
%>
<table width="100%" >
<tr>
    <td width="50%"><iframe src="http://www1.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300" ></iframe></td>
    <td width="50%"><iframe src="http://www2.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300" ></iframe></td>
</tr>
<tr>
    <td width="50%"><iframe src="http://www3.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300" ></iframe></td>
    <td width="50%"><iframe src="http://www4.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300" ></iframe></td>
</tr>
<tr>
    <td width="50%"><iframe src="http://www5.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300" ></iframe></td>
    <td width="50%"><iframe src="http://www6.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300" ></iframe></td>
</tr>
<tr>
    <td width="50%"><iframe src="http://www7.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300"></iframe></td>
    <td width="50%"><iframe src="http://www8.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300"></iframe></td>
</tr>
<!--
<tr>
    <td width="50%"><iframe src="http://www9.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300"></iframe></td>
    <td width="50%"></td>
</tr>

<tr>
    <td width="50%"><iframe src="http://www9.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300"></iframe></td>
    <td width="50%"><iframe src="http://www10.10x10.co.kr/chtml/make_mainApp_refresh.asp?idx=<%= idx %>&allrefresh=<%=allrefreshVal %>" width="50%" height="300"></iframe></td>
</tr>
-->
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->