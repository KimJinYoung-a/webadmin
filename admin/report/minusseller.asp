<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/checkupcls.asp"-->
<%
''수정요망.
'dbget.close()	:	response.End



dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nextdateStr,searchnextdate,searchstartpredate
dim designer,page
dim totalcnt
totalcnt = 0

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
designer = request("designer")
page = request("page")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

searchstartpredate = Left(CStr(dateserial(yyyy1,mm1-1,dd1)),10)
searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

if page="" then page=1

dim ominus
set ominus = new CCheckUp
ominus.FPageSize = 50
ominus.FCurrPage = page
ominus.FRectDesignerID = designer
ominus.FRectRegStartPre = searchstartpredate
ominus.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
ominus.FRectRegEnd = searchnextdate

ominus.getMinusItemList

dim i,ix,p
%>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		디자이너 :
		<% drawSelectBoxDesigner "designer",designer %>
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td width="60">ItemID</td>
    	<td width="60">Designer</td>
    	<td >아이템명</td>
    	<td width="100">옵션명</td>
    	<td width="100">판매갯수</td>
    	<td width="100">반품갯수</td>
    	<td width="100">반품율</td>
    	<td width="100">총금액</td>
    </tr>
    <% for i=0 to ominus.FResultCount-1 %>
    <%
    if ominus.FItemList(i).FRealSell<>0 then
    	p = CLng(ominus.FItemList(i).FCount/ominus.FItemList(i).FRealSell*100)
    else
    	p = "???"
    end if
    %>
    <tr bgcolor="#FFFFFF">
    	<td ><%= ominus.FItemList(i).FItemID %></td>
    	<td ><%= ominus.FItemList(i).FDesignerID %></td>
    	<td ><%= ominus.FItemList(i).FItemName %></td>
    	<td ><%= ominus.FItemList(i).FItemOptionName %></td>
    	<td ><%= ForMatNumber(ominus.FItemList(i).FRealSell,0) %></td>
    	<td ><%= ForMatNumber(ominus.FItemList(i).FCount,0) %></td>
    	<% if p = "???" then %>
    	<td ><%= p %></td>
    	<% elseif IsNumeric(p) and (p>10) then %>
    	<td ><font color="red"><%= p %>%</font></td>
    	<% else %>
    	<td ><%= p %>%</td>
    	<% end if %>
    	<td ><%= ForMatNumber(ominus.FItemList(i).FPriceSum,0) %></td>
    </tr>
	<% totalcnt = totalcnt  + ominus.FItemList(i).FCount  %>
    <% next %>
    <tr bgcolor="#FFFFFF" >
    	<td colspan="8" align="right">
			총반품갯수 : <font color="red"><% = Cint(totalcnt) %></font>개&nbsp;&nbsp;
    	</td>
    </tr>
    <tr bgcolor="#FFFFFF" >
    	<td colspan="8" align="center">
    	<% if ominus.HasPreScroll then %>
			<a href="?page=<%= ominus.StarScrollPage-1 %>&designer=<%= designer %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + ominus.StarScrollPage to ominus.StarScrollPage + ominus.FScrollCount - 1 %>
			<% if (ix > ominus.FTotalpage) then Exit for %>
			<% if CStr(ix) = CStr(ominus.FCurrPage) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="?page=<%= ix %>&designer=<%= designer %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ominus.HasNextScroll then %>
			<a href="?page=<%= ix %>&designer=<%= designer %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
    	</td>
    </tr>
<%
set ominus = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->