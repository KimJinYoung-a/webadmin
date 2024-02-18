<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim cdl,cdm,cds,oldlist


yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")
oldlist = request("oldlist")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CJumunMaster
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
oreport.FRectCD1 = cdl
oreport.FRectCD2 = cdm
oreport.FRectCD3 = cds
oreport.FRectOldJumun = oldlist
oreport.SearchChannalDailySellRePort

dim i,p1,p2
%>
<script language="JavaScript">
<!--
function changecontent(){
	frm.submit();
}
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
		<br>
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>&nbsp;
		카테고리선택 :
		<% DrawSelectBoxCategoryLarge "cdl", cdl %>
		<% DrawSelectBoxCategoryMid "cdm", cdl,cdm %>
		<% DrawSelectBoxCategorySmall "cds", cdl,cdm,cds %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">기간</font></td>
          <td class="a" width="600"><font color="#FFFFFF"></font></td>
          <td class="a" width="120"><font color="#FFFFFF">내용</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
			end if

			if oreport.maxc<>0 then
				p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
			end if
		%>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
		  <td width="120" height="10">
          	<%= oreport.FMasterItemList(i).Fsitename %>(<%= oreport.FMasterItemList(i).GetDpartName %>)
          </td>
          <td  height="10" width="600">
			<div align="left"> <img src="/images/dot1.gif" height="4" width="<%= p1 %>%"></div><br>
          	<div align="left"> <img src="/images/dot2.gif" height="4" width="<%= p2 %>%"></div>
          </td>
		  <td class="a" width="160" align="right">
		    <%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>원 <br>
          	<%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>건
		  </td>
        </tr>
        <% next %>
</table>
<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->