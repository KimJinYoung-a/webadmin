<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order_category_saacls2.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr, page
dim rpttype,addstand
dim gubun
dim i

addstand = request("addstand")
if addstand = "" then addstand = 1
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
rpttype = request("rpttype")
page = request("page")
gubun = request("gubun")

if page="" then page=1

nowdateStr = CStr(now())


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

if (rpttype="") then rpttype="day"

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

const MAXBARSIZE = 500
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>&nbsp;
		<select name="gubun">
			<option value="">선택</option>
			<option value="01" <% if gubun = "01" then response.write "selected" %>>디자인</option>
			<option value="02" <% if gubun = "02" then response.write "selected" %>>플라워</option>
			<option value="03" <% if gubun = "03" then response.write "selected" %>>패션,쥬얼리</option>
			<option value="04" <% if gubun = "04" then response.write "selected" %>>애견</option>
			<option value="05" <% if gubun = "05" then response.write "selected" %>>뷰티케어</option>
			<option value="06" <% if gubun = "06" then response.write "selected" %>>보드게임</option>
		</select>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<%
dim orderreport2
set orderreport2 = new UserJoinClass
orderreport2.FRectStart = startdateStr
orderreport2.FRectEnd =  nextdateStr
orderreport2.FRectGroup = rpttype
orderreport2.FRectGubun = gubun
orderreport2.GetUserJoinByNai2
%>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td colspan="5">2. 연령별 매출</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="100">전체</td>
    	<td width="100" align="right">
    		<%= FormatNumber(orderreport2.FNaiMaster.FManTotal + orderreport2.FNaiMaster.FWoManTotal,0) %>
    	</td>
    	<td width="50" align="right">
    		<%= orderreport2.FNaiMaster.GetManTotalPercent %> (%)<br>
    		<%= orderreport2.FNaiMaster.GetWoManTotalPercent %> (%)
    	</td>
    	<td width="50" align="right">100 (%)</td>
    	<td></td>
    </tr>
    <% for i=0 to orderreport2.FNaiMaster.FItemCount - 1  %>
    <tr bgcolor="#FFFFFF">
    	<td width="100"><%= orderreport2.FNaiMaster.FItemList(i).FNaiStr %></td>
    	<td width="100" align="right">
    		<%= orderreport2.FNaiMaster.FItemList(i).FManCount + orderreport2.FNaiMaster.FItemList(i).FWoManCount %>
    	</td>
    	<td width="50" align="right">
    		<%= orderreport2.FNaiMaster.GetManPercent(i) %> (%)<br>
    		<%= orderreport2.FNaiMaster.GetWoManPercent(i) %> (%)
    	</td>
    	<td width="50" align="right"><%= orderreport2.FNaiMaster.GetTotPercent(i) %> (%)</td>
    	<td>
    		<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%= CInt(MAXBARSIZE * orderreport2.FNaiMaster.GetManPercent(i) / 100) %>"><br>
    		<img src="http://partner.10x10.co.kr/images/dot2.gif" height="4" width="<%= CInt(MAXBARSIZE * orderreport2.FNaiMaster.GetWoManPercent(i) / 100) %>">
    	</td>
    </tr>
    <% next %>
</table>
<%
set orderreport2 = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->