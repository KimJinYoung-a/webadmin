<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/resending_reportcls.asp"-->
<%
dim oreport,avgvalue,ix
dim nextdateStr,startdateStr,divcd

divcd = request("divcd")
startdateStr = request("startdate")
nextdateStr = request("enddate")

set oreport = new CReportMaster
oreport.FRectStart = startdateStr
oreport.FRectEnd =  nextdateStr
oreport.FRectDivcd =  divcd
oreport.SearchReport


for ix = 0 to oreport.FResultCount - 1
avgvalue = (oreport.FMasterItemList(ix).Fcount/oreport.Ftotalcount) * 100
response.write "high" & ix & "=" & Clng(avgvalue) * 2 & "&cause" & ix & "=" & oreport.FMasterItemList(ix).Fcausedetail & "&"
next

set oreport = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->