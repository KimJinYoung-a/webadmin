<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","-1"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->

<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim countvb
countvb = 40
const Maxlines = 10
dim totalpage, totalnum, q, i
dim totalcount


dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim gotopage,ojumun
dim fromDate,toDate,itemidlist,settle2

totalcount = 0
itemidlist = request("itemidlist")
settle2 = request("settle2")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (settle2="") then settle2= "d"

	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)

set ojumun = new CReportMaster

ojumun.FRectRegStart = fromDate
ojumun.FRectRegEnd = toDate
ojumun.FRectSettle2 = settle2
ojumun.FRectItemList = itemidlist
ojumun.SearchEachItemReport

%>
<table width="400" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="100" class="a"><font color="#FFFFFF">구분</font></td>
          <td class="a" width="150"><font color="#FFFFFF">아이템번호</font></td>
          <td class="a" width="150"><font color="#FFFFFF">총판매가</font></td>
          <td class="a" width="150"><font color="#FFFFFF">판매갯수</font></td>
        </tr>
        <% for i=0 to ojumun.FResultCount - 1 %>
        <tr bgcolor="#FFFFFF" height="10">
          <td height="10" class="a"><%= ojumun.FMasterItemList(i).Fselldate %></td>
		  <td height="10" class="a"><%= ojumun.FMasterItemList(i).Fitemid %></td>
		  <td class="a">
		  <% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
		  	<%= FormatNumber(FormatCurrency(ojumun.FMasterItemList(i).Fselltotal),0) & "원"%>&nbsp;
		  <% end if %>
		  </td>
		  <td class="a"><%= ojumun.FMasterItemList(i).Fsellcnt & "(EA)"%></td>
        </tr>
		<% totalcount = totalcount + ojumun.FMasterItemList(i).Fsellcnt %>
        <% next %>
</table>

<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"  codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" ID=graph_report WIDTH=700 HEIGHT=400>
 <PARAM NAME=movie VALUE="graph_report.swf">
 <PARAM NAME=quality VALUE=high>
 <PARAM NAME=bgcolor VALUE=#FFFFFF>
 <EMBED src="graph_report.swf" quality=high bgcolor=#FFFFFF  WIDTH=700 HEIGHT=400	swLiveConnect=true NAME=graph_report TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED>
</OBJECT>
<SCRIPT LANGUAGE="JavaScript">
<!--

		var movie = window.document.graph_report ;
		movie.SetVariable("startdate","<% = fromDate  %>");
		movie.SetVariable("enddate","<% = toDate %>");
		movie.SetVariable("settle","<% = settle2  %>");
		movie.SetVariable("itemlist","<% = itemidlist  %>");
//-->
</SCRIPT>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
