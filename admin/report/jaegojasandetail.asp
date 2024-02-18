<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jaegocls.asp"-->
<%
dim yyyy1,mm1,designer,mwdiv
yyyy1 = request("yyyy1")
mm1 = request("mm1")
mwdiv = request("mwdiv")
designer = request("designer")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim ojaego, yyyymm
yyyymm = yyyy1 + "-" + mm1

set ojaego = new CJaegoEval
ojaego.FRectYYYYMM   = yyyymm
ojaego.FRectMwDiv = mwdiv
ojaego.FRectDesigner = designer
ojaego.GetMonthJeagoDetail


dim totno, totbuy, totsell,i
%>
<h2>수정중</h2>
브랜드 | 상품코드 | 상품명 | 옵션 | 소비자가 | 매입가 | 재고수량 | 월매출액 | 월매입액
<br>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<form name="frm" method="get" action="">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	월말 : <% DrawYMBox yyyy1,mm1 %>
        		브랜드 :	<% drawSelectBoxDesignerwithName "designer", designer %>
        		<input type="radio" name="mwdiv" value="M" <% if mwdiv="M" then response.write "checked" %> >매입
        		<input type="radio" name="mwdiv" value="W" <% if mwdiv="W" then response.write "checked" %> >위탁
	        </td>
	        <td valign="top" align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#000000">
    <tr bgcolor="DDDDFF" align="center">
    	<td width="80">브랜드</td>
    	<td width="25">구분</td>
    	<td width="50">상품코드</td>
    	<td>상품명[옵션]</td>
    	<td width="60">월말<br>재고수량</td>
    	<td width="60">재고총액<br>(소비자가)</td>
    	<td width="60">재고총액<br>(매입가)</td>
    	<td width="80">금월매입총액</td>
    	<td width="80">금월매출총액</td>
    	<td width="40">회전율</td>
    	<td width="80">3개월매출총액</td>
    	<td width="40">회전율</td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    %>
    <tr bgcolor="#FFFFFF">
    	<td align="center"><%= ojaego.FItemList(i).Fmakerid %></td>
    	<td></td>
    	<td align="center"><%= ojaego.FItemList(i).Fitemid %></td>
    	<td><%= ojaego.FItemList(i).Fitemname %><br><font color="blue"><%= ojaego.FItemList(i).Fitemoptionname %></font></td>
    	<td align="center"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
    	<td align="center">총계</td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td align="center" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    </tr>
</table>


<%
set ojaego = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->