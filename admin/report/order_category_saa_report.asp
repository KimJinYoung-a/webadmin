<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/order_category_saacls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr, page
dim rpttype,addstand,oldDataYn
dim cdl
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
cdl = request("cdl")
oldDataYn=request("oldDataYn")

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

'response.write startdateStr
nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)


dim orderreport
set orderreport = new UserJoinClass
orderreport.FRectStart = startdateStr
orderreport.FRectEnd =  nextdateStr
orderreport.FRectGroup = rpttype
orderreport.FRectGubun = cdl
orderreport.FoldDataYn = oldDataYn
orderreport.GetUserJoinByNai

const MAXBARSIZE = 500
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		<input type="checkbox" name="oldDataYn" <% if oldDataYn="on" then response.write "checked" %>>6���� ���� ����
		�Ⱓ :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>&nbsp;
		<% SelectBoxCategoryLarge cdl %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td colspan="5">1. ���ɺ� ������</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="100">��ü</td>
    	<td width="100" align="right">
    		<%= FormatNumber(orderreport.FNaiMaster.FManTotal,0) %><br>
    		<%= FormatNumber(orderreport.FNaiMaster.FWoManTotal,0) %>
    	</td>
    	<td width="50" align="right">
    		<%= orderreport.FNaiMaster.GetManTotalPercent %> (%)<br>
    		<%= orderreport.FNaiMaster.GetWoManTotalPercent %> (%)
    	</td>
    	<td width="50" align="right">100 (%)</td>
    	<td></td>
    </tr>
    <% for i=0 to orderreport.FNaiMaster.FItemCount - 1  %>
    <tr bgcolor="#FFFFFF">
    	<td width="100"><%= orderreport.FNaiMaster.FItemList(i).FNaiStr %></td>
    	<td width="100" align="right">
    		<%= FormatNumber(orderreport.FNaiMaster.FItemList(i).FManCount,0) %><br>
    		<%= FormatNumber(orderreport.FNaiMaster.FItemList(i).FWoManCount,0) %>
    	</td>
    	<td width="50" align="right">
    		<%= orderreport.FNaiMaster.GetManPercent(i) %> (%)<br>
    		<%= orderreport.FNaiMaster.GetWoManPercent(i) %> (%)
    	</td>
    	<td width="50" align="right"><%= orderreport.FNaiMaster.GetTotPercent(i) %> (%)</td>
    	<td>
    		<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%= CInt(MAXBARSIZE * orderreport.FNaiMaster.GetManPercent(i) / 100) %>"><br>
    		<img src="http://partner.10x10.co.kr/images/dot2.gif" height="4" width="<%= CInt(MAXBARSIZE * orderreport.FNaiMaster.GetWoManPercent(i) / 100) %>">
    	</td>
    </tr>
    <% next %>
</table>
<%
set orderreport = Nothing
%>
<br>
<%
dim orderreport2
set orderreport2 = new UserJoinClass
orderreport2.FRectStart = startdateStr
orderreport2.FRectEnd =  nextdateStr
orderreport2.FRectGroup = rpttype
orderreport2.FRectGubun = cdl
orderreport2.FoldDataYn = oldDataYn

orderreport2.GetUserJoinByNai2
%>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td colspan="5">2. ���ɺ� ����</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="100">��ü</td>
    	<td width="100" align="right">
    		<%= FormatNumber(orderreport2.FNaiMaster.FManTotal,0) %><br>
    		<%= FormatNumber(orderreport2.FNaiMaster.FWoManTotal,0) %>
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
    		<%= FormatNumber(orderreport2.FNaiMaster.FItemList(i).FManCount,0) %><br>
    		<%= FormatNumber(orderreport2.FNaiMaster.FItemList(i).FWoManCount,0) %>
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