<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%
dim yyyy1,mm1,isusing,sysorreal,mwgubun,makerid
yyyy1 = request("yyyy1")
mm1 = request("mm1")
isusing = request("isusing")
sysorreal = request("sysorreal")
mwgubun = request("mwgubun")
makerid = request("makerid")

if sysorreal="" then sysorreal="real"
dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


dim ojaego
set ojaego = new CMonthlyStock
ojaego.FRectYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectIsUsing = isusing
ojaego.FRectGubun = sysorreal
ojaego.FRectMakerid = makerid
ojaego.FRectMwDiv = mwgubun

if makerid<>"" then
	ojaego.GetMonthlyRealJeagoDetailByMaker
else
	ojaego.GetMonthlyRealJeagoDetail
end if

dim i
dim totno, totbuy, totsell
%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %> 말일자 재고자산
			&nbsp;&nbsp;
			<font color="#CC3333">브랜드 :</font> <% drawSelectBoxDesignerwithName "makerid",makerid %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">재고구분:</font>
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >시스템재고
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >실사재고
        	&nbsp;&nbsp;&nbsp;
        	<font color="#CC3333">상품사용구분:</font>
        	<input type="radio" name="isusing" value="" <% if isusing="" then response.write "checked" %> >전체
        	<input type="radio" name="isusing" value="Y" <% if isusing="Y" then response.write "checked" %> >사용함
        	<input type="radio" name="isusing" value="N" <% if isusing="N" then response.write "checked" %> >사용안함
        	&nbsp;&nbsp;&nbsp;
        	<font color="#CC3333">매입구분:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >전체
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >매입
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >위탁
        	<input type="radio" name="mwgubun" value="U" <% if mwgubun="U" then response.write "checked" %> >업체
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<% if makerid<>"" then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">상품코드</td>
    	<td>상품명</td>
    	<td>옵션명</td>
    	<td width="60">매입구분</td>
    	<td width="60">재고수량</td>
    	<td width="70">소비자가계</td>
    	<td width="60">평균마진</td>
    	<td width="70">매입가계</td>
    	<td>비고</td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    %>
    <% if (ojaego.FItemList(i).FIsUsing="N") or (ojaego.FItemList(i).FOptionUsing="N") then %>
    <tr align="center" bgcolor="#CCCCCC">
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    <% end if %>
    	<td><a href="javascript:TnPopItemStock('<%= ojaego.FItemList(i).FItemID %>');"><%= ojaego.FItemList(i).FItemID %></a></td>
    	<td align="left"><%= ojaego.FItemList(i).FItemName %></td>
    	<td><%= ojaego.FItemList(i).FItemOptionName %></td>
    	<td><%= fncolor(ojaego.FItemList(i).FMaeIpGubun,"mw") %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td>
    		<% if ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td></td>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td colspan="4">총계</td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>

    	<td></td>
    </tr>
</table>

<% else %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2">브랜드ID</td>
    	<td rowspan="2" width="50">총재고<br>수량</td>
    	<td rowspan="2" width="70">총재고액<br>(공급가)<br>(S)</td>
    	
    	<td colspan="4">이전3개월 회전율</td>
    	<td colspan="4"><%= yyyy1 %>년 <%= mm1 %>월 회전율</td>
    	
    	<td rowspan="2" width="80"><%= yyyy1 %>년 <%= mm1 %>월<br>매입(입고)액<br>(공급가)</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="70">ON매출액<br>(판매가)(O)</td>
    	<td width="70">출고액<br>(판매가)(C)</td>
    	<td width="90"><b>판매/출고총액<br>(O+C)</b></td>
    	<td width="90"><font color="blue"><b>회전율<br>(R=(O+C)/S)</b></font></td>
    	
    	<td width="70">ON매출액<br>(판매가)(O)</td>
    	<td width="70">출고액<br>(판매가)(C)</td>
    	<td width="90"><b>판매/출고총액<br>(O+C)</b></td>
    	<td width="90"><font color="blue"><b>회전율<br>(R=(O+C)/S)</b></font></td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    %>
    <% if ojaego.FItemList(i).FMakerUsing="Y" then %>
    <tr align="center" bgcolor="#FFFFFF">
    <% else %>
    <tr align="center" bgcolor="#CCCCCC">
    <% end if %>
    	<td align="left"><a href="monthlystock_detail.asp?menupos=<%= menupos %>&mwgubun=<%= ojaego.FItemList(i).FMaeIpGubun %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&sysorreal=<%= sysorreal %>&isusing=<%= isusing %>&makerid=<%= ojaego.FItemList(i).FMakerid %>" ><%= ojaego.FItemList(i).FMakerid %></a></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><b><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %><b></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>총계</td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
		<td></td>
		<td></td>
		<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    </tr>
</table>
<% end if %>

<%
set ojaego = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->