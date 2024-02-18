<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim countvb
countvb = 40
const Maxlines = 10
dim totalpage, totalnum, q, ix
dim totalcount


dim oreport
dim fromDate,toDate,itemidlist,settle2

totalcount = 0
itemidlist = request("itemidlist")
settle2 = request("settle2")
fromDate = request("fromDate")
toDate = request("toDate")

if (settle2="") then settle2= "d"

set oreport = new CReportMaster

oreport.FRectRegStart = fromDate
oreport.FRectRegEnd = toDate
oreport.FRectSettle2 = settle2
oreport.FRectItemList = itemidlist
oreport.SearchEachItemReport

response.write "&row=" & oreport.FResultCount & "&"
for ix = 0 to oreport.FResultCount -1

 response.write "date" & ix + 1 & "=" & oreport.FMasterItemList(ix).Fselldate  & "&tea" & ix + 1 & "=" & oreport.FMasterItemList(ix).Fsellcnt  & "&tmoney" & ix + 1 & "=" & FormatNumber(oreport.FMasterItemList(ix).Fselltotal,0) &  "¿ø&<br>"

next

set oreport = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->