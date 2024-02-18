<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/analysiscls.asp"-->
<%
dim yyyy1,mm1, oldlist
yyyy1 = request("yyyy1")
mm1 = request("mm1")
oldlist= request("oldlist")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now),1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim oanal
set oanal = new CAnalysis
oanal.FRectYYYYMMDD = yyyy1 + "-" + mm1 + "-" + "01"
oanal.FRectYYYYMMDD2 = CStr(DateAdd("m",1,oanal.FRectYYYYMMDD))
oanal.FRectOldList = oldlist
oanal.getOnLineDailyGainSum

dim i

dim orgitemcost
dim totalsum
dim miletotalprice
dim tencardspend
dim allatdiscountprice
dim spendmembership
dim subtotalprice
dim ErrSubTotal
dim itemtotalsum
dim itembuysum
dim deliverytotalsum
dim tenbeaCount
dim GainSum


%>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="">
	<tr>
		<td class="a" >
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
		검색대상년월:<% DrawYMBox yyyy1,mm1 %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>



<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#CCCCCC" class="a" >
<tr align="center" bgcolor="#E6E6E6" >
    <td>날짜</td>
    <td>소비자가</td>
    <td>주문합계</td>
    <td>마일리지</td>
    <td>쿠폰</td>
    <td>올@카드</td>
    <td>SKT</td>
    <td>실매출</td>
    <td>ERR</td>
    <td>상품매출</td>
    <td>상품매입</td>
    <td>배송비매출액</td>
    <td>텐배송건</td>
    <td>일별수익</td>
</tr>
<% for i=0 to oanal.FResultCount -1 %>
<%
orgitemcost        =    orgitemcost             +               oanal.FItemList(i).Forgitemcost
totalsum           =    totalsum                +               oanal.FItemList(i).Ftotalsum
miletotalprice     =    miletotalprice          +               oanal.FItemList(i).Fmiletotalprice
tencardspend       =    tencardspend            +               oanal.FItemList(i).Ftencardspend
allatdiscountprice =    allatdiscountprice      +               oanal.FItemList(i).Fallatdiscountprice
spendmembership    =    spendmembership         +               oanal.FItemList(i).Fspendmembership
subtotalprice      =    subtotalprice           +               oanal.FItemList(i).Fsubtotalprice
ErrSubTotal        =    ErrSubTotal             +               oanal.FItemList(i).GetErrSubTotal
itemtotalsum       =    itemtotalsum            +               oanal.FItemList(i).Fitemtotalsum
itembuysum         =    itembuysum              +               oanal.FItemList(i).Fitembuysum
deliverytotalsum   =    deliverytotalsum        +               oanal.FItemList(i).Fdeliverytotalsum
tenbeaCount        =    tenbeaCount             +               oanal.FItemList(i).FtenbeaCount
GainSum            =    GainSum                 +               oanal.FItemList(i).GetGainSum
%>
<tr bgcolor="#FFFFFF" >
    <td><%= oanal.FItemList(i).Fyyyymmdd %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Forgitemcost,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Ftotalsum,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Fmiletotalprice,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Ftencardspend,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Fallatdiscountprice,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Fspendmembership,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Fsubtotalprice,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).GetErrSubTotal,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Fitemtotalsum,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Fitembuysum,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).Fdeliverytotalsum,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).FtenbeaCount,0) %></td>
    <td align="right"><%= FormatNumber(oanal.FItemList(i).GetGainSum,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" >
    <td>Total</td>
    <td align="right"><%= FormatNumber(orgitemcost,0) %></td>
    <td align="right"><%= FormatNumber(totalsum,0) %></td>
    <td align="right"><%= FormatNumber(miletotalprice,0) %></td>
    <td align="right"><%= FormatNumber(tencardspend,0) %></td>
    <td align="right"><%= FormatNumber(allatdiscountprice,0) %></td>
    <td align="right"><%= FormatNumber(spendmembership,0) %></td>
    <td align="right"><%= FormatNumber(subtotalprice,0) %></td>
    <td align="right"><%= FormatNumber(ErrSubTotal,0) %></td>
    <td align="right"><%= FormatNumber(itemtotalsum,0) %></td>
    <td align="right"><%= FormatNumber(itembuysum,0) %></td>
    <td align="right"><%= FormatNumber(deliverytotalsum,0) %></td>
    <td align="right"><%= FormatNumber(tenbeaCount,0) %></td>
    <td align="right"><%= FormatNumber(GainSum,0) %></td>
</tr>
</table>

<%
set oanal = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
