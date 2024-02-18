<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/order_saacls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr, page
dim rpttype,addstand,oldDataYn

addstand = request("addstand")
if addstand = "" then addstand = 1
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
rpttype = request("rpttype")
oldDataYn=request("oldDataYn")
page = request("page")

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

dim orderreport
set orderreport = new UserJoinClass
orderreport.FRectStart = startdateStr
orderreport.FRectEnd =  nextdateStr
orderreport.FRectGroup = rpttype
orderreport.FoldDataYn = oldDataYn
orderreport.GetUserJoinBySex

dim i
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" ><input type="checkbox" name="oldDataYn" <% if oldDataYn="on" then response.write "checked" %>>6개월 이전 내역
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="radio" name="addstand" value="1" <% if addstand = "1" then response.write "checked" %>> 주문지 <input type="radio" name="addstand" value="2" <% if addstand = "2" then response.write "checked" %>> 배송지
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<%
const MAXBARSIZE = 500
dim totno, MsexPercent,WsexPercent

totno = orderreport.FManNo + orderreport.FWoManNo

if totno<>0 then
	MsexPercent = CInt(orderreport.FManNo/totno*100)
	WsexPercent = CInt(orderreport.FWoManNo/totno*100)
else
	MsexPercent = 0
	WsexPercent = 0
end if
%>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td colspan="4">1. 성별 구매비율 / 주문수 </td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="100">전체</td>
    	<td width="100" align="right"><%= FormatNumber(totno,0) %></td>
    	<td width="100" align="right">100 (%)</td>
    	<td></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td>남성</td>
    	<td align="right"><%= FormatNumber(orderreport.FManNo,0) %></td>
    	<td align="right"><%= MsexPercent %> (%)</td>
    	<td><img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%= CInt(MAXBARSIZE * MsexPercent / 100) %>"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td>여성</td>
    	<td align="right"><%= FormatNumber(orderreport.FWoManNo,0) %></td>
    	<td align="right"><%= WsexPercent %> (%)</td>
    	<td><img src="http://partner.10x10.co.kr/images/dot2.gif" height="4" width="<%= CInt(MAXBARSIZE * WsexPercent / 100) %>"></td>
    </tr>
</table>
<br>
<%
orderreport.GetUserJoinByNai
%>

<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td colspan="5">2. 연령별 구매비율 / 주문수 (현재 나이기준)</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="100">전체</td>
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
<br>
<%
dim tmppercent
orderreport.FRectBeasongArea = addstand
orderreport.GetUserJoinByArea
%>

<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td colspan="5">3. 지역별 구매비율 / 주문수 (정렬:주문수 순)
    		<br />※ 상품 주문시 주문자 주소는 저장되지 않아 회원DB를 사용할 수 밖에 없는데 회원정보도 몇년전부터 주소를 받지 않고 있습니다.
    		<br />&nbsp;&nbsp;&nbsp;&nbsp;그래서 주문지 검색은 되지않고 배송지로만 검색이 됩니다.
    	</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="120">전체</td>
    	<td width="100" align="right"><%= FormatNumber(orderreport.FTotalUsercount,0) %></td>
    	<td width="100" align="right">100 (%)</td>
    	<td></td>
    </tr>
    <% for i=0 to orderreport.FResultCount -1 %>
    <%
    if orderreport.FTotalUsercount=0 then
    	tmppercent = 0
    else
    	tmppercent = CInt(orderreport.FItemList(i).FCount/orderreport.FTotalUsercount*100)
    end if
    %>
    <tr bgcolor="#FFFFFF">
    	<td width="120"><%= CHKIIF(orderreport.FItemList(i).FArea="X","예외",orderreport.FItemList(i).FArea) %> </td>
    	<td width="100" align="right"><%= FormatNumber(orderreport.FItemList(i).FCount,0) %></td>
    	<td width="100" align="right"><%= tmppercent %> (%)</td>
    	<td><img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%= CInt(MAXBARSIZE * tmppercent / 100) %>"></td>
    </tr>
    <% next %>
</table>
<%
set orderreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->