<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/customer_board_reportcls.asp" -->
<%
dim oreport,avgvalue,ix
dim nextdateStr,startdateStr

startdateStr = request("startdate")
nextdateStr = request("enddate")

set oreport = new CReportMaster
oreport.FRectStart = startdateStr
oreport.FRectEnd =  nextdateStr
oreport.SearchReport


for ix = 0 to oreport.FResultCount - 1
avgvalue = (oreport.FMasterItemList(ix).Fcount/oreport.Ftotalcount) * 100
response.write "high" & ix & "=" & Clng(avgvalue) * 2 & "&"
next

set oreport = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->