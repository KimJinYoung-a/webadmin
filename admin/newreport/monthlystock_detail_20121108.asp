<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%
Const isShowIpgoPrice = FALSE
Const isOnlySys = FALSE
Const isShowOffreturn = FALSE
Dim isViewUser : isViewUser = (session("ssAdminPsn")="17")

dim yyyy1,mm1,isusing,sysorreal,mwgubun,makerid,newitem,itemgubun,vatyn
dim research,offrt2on
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1         = requestCheckvar(request("mm1"),10)
isusing     = requestCheckvar(request("isusing"),10)
sysorreal   = requestCheckvar(request("sysorreal"),10)
mwgubun     = requestCheckvar(request("mwgubun"),10)
makerid     = requestCheckvar(request("makerid"),32)
newitem     = requestCheckvar(request("newitem"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
vatyn       = requestCheckvar(request("vatyn"),10)
offrt2on    = requestCheckvar(request("offrt2on"),10)
research    = requestCheckvar(request("research"),10)

if sysorreal="" then sysorreal="sys" ''real
if (research="") or (not isShowOffreturn) then offrt2on="on"

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


dim ojaego
set ojaego = new CMonthlyStock
ojaego.FRectYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
ojaego.FRectIsUsing  = isusing
ojaego.FRectGubun    = sysorreal
ojaego.FRectMakerid  = makerid
ojaego.FRectMwDiv    = mwgubun
ojaego.FRectNewItem  = newitem
ojaego.FRectVatYn    = vatyn
ojaego.FRectItemGubun = itemgubun
ojaego.FRectOFFReturn2OnStock = offrt2on


if makerid<>"" then
    ojaego.GetMonthlyRealJeagoDetailByMakerNew
else
	ojaego.GetMonthlyRealJeagoDetailNew
	
end if

dim i
dim totno, totbuy, totsell, totavgIpgoPrice
dim iURL
%>
<script language='javascript'>
<% if (Not isViewUser) then %>
function TnPopItemStockWithGubun(itemgubun,itemid,itemoption){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
<% end if %>
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %> 말일자 재고자산
			&nbsp;&nbsp;
			<font color="#CC3333">브랜드 :</font> <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;&nbsp;
			<input type="radio" name="newitem" value="" <% if newitem="" then response.write "checked" %> >전체상품
        	<input type="radio" name="newitem" value="new" <% if newitem="new" then response.write "checked" %> >신상품
        	
        	&nbsp;&nbsp;|&nbsp;&nbsp;
	        	과세구분
	        	<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >전체
	        	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >과세
	        	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >면세
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <% IF not (isOnlySys) THEN %>
			<font color="#CC3333">재고구분:</font>
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >시스템재고
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >실사재고
        	&nbsp;&nbsp;
        	<% end if %>
        	<font color="#CC3333">상품사용구분:</font>
        	<input type="radio" name="isusing" value="" <% if isusing="" then response.write "checked" %> >전체
        	<input type="radio" name="isusing" value="Y" <% if isusing="Y" then response.write "checked" %> >사용함
        	<input type="radio" name="isusing" value="N" <% if isusing="N" then response.write "checked" %> >사용안함
        	&nbsp;&nbsp;
        	<font color="#CC3333">매입구분:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >전체
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >매입
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >특정
        	<!-- <input type="radio" name="mwgubun" value="U" <% if mwgubun="U" then response.write "checked" %> >업체 -->
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >미지정
        	<font color="#CC3333">상품구분:</font>
        	<select name="itemgubun">
        	<option value="10" <%= CHKIIF(itemgubun="10","selected" ,"") %> >일반(10)
        	<option value="70" <%= CHKIIF(itemgubun="70","selected" ,"") %> >소포품(70)
        	<option value="80" <%= CHKIIF(itemgubun="80","selected" ,"") %> >사은품(80)
        	<option value="90" <%= CHKIIF(itemgubun="90","selected" ,"") %> >오프전용(90)
        	</select>
        	<% if (isShowOffreturn) then %>
            <br><input type="checkbox" name="offrt2on" <%= CHKIIF(offrt2on="on","checked","") %> >매장반품On재고로
            <% end if %>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>


<% if makerid<>"" then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (Not isViewUser) then %>
    	<td width="65">등록일</td>
    	<td width="40"></td>
    <% end if %>
    	<td>구분</td>
    	<td width="50">상품코드</td>
    	<td width="50">옵션코드</td>
    	<td>상품명</td>
    	<td>옵션명</td>
    	<td width="50">구분</td>
    	<td width="50">재고수량</td>
    	<td width="80">소비자가</td>
    	<td width="50">매입마진</td>
    	<td width="80">매입가</td>
    	<% IF(isShowIpgoPrice)THEN %><td width="90">매입가<br>(매입시기준)</td><% end if %>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    
    if not IsNULL(ojaego.FItemList(i).FavgIpgoPriceSum) then
        totavgIpgoPrice = totavgIpgoPrice + ojaego.FItemList(i).FavgIpgoPriceSum
    end if
    
    %>
    <% if (Not isViewUser) and ((ojaego.FItemList(i).FIsUsing="N") or (ojaego.FItemList(i).FOptionUsing="N")) then %>
    <tr align="center" bgcolor="<%= adminColor("dgray") %>">
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    <% end if %>
        <% if (Not isViewUser) then %>
    	<td><%= left(ojaego.FItemList(i).Fregdate,10) %></td>
    	<td>
    		<% if (datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD)) <= 3 then %>
    		<font color="red"><%= datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD) %>개월</font>
    		<% else %>
    		<%= datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD) %>개월
    		<% end if %>
    	</td>
    	<% end if %>
    	<td><%= ojaego.FItemList(i).FItemGubun %></td>
    	<% if ( isViewUser) then %>
    	<td><%= ojaego.FItemList(i).FItemID %></td>
    	<% else %>
    	<td><a href="javascript:TnPopItemStockWithGubun('<%= ojaego.FItemList(i).FItemGubun %>','<%= ojaego.FItemList(i).FItemID %>','<%= ojaego.FItemList(i).FItemOption %>');"><%= ojaego.FItemList(i).FItemID %></a></td>
    	<% end if %>
    	<td><%= ojaego.FItemList(i).FItemOption %></td>
    	<td align="left"><%= ojaego.FItemList(i).FItemName %></td>
    	<td><%= ojaego.FItemList(i).FItemOptionName %></td>
    	<td><%= ojaego.FItemList(i).getMaeipGubunName %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td>
    		<% if ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<% IF(isShowIpgoPrice)THEN %>
    	<td align="right">
    	<% if IsNULL(ojaego.FItemList(i).FavgIpgoPriceSum) then %>
    	-
    	<% else %>
    	<%= FormatNumber(ojaego.FItemList(i).FavgIpgoPriceSum,0) %>
    	<% end if %>
    	</td>
    	<% end if %>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
        <% if ( isViewUser) then %>
        <td colspan="6">총계</td>
        <% else %>
    	<td colspan="8">총계</td>
    	<% end if %>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<% IF(isShowIpgoPrice)THEN %><td align="right" ><%= FormatNumber(totavgIpgoPrice,0) %></td><% end if %>
    </tr>
</table>

<% else %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">브랜드ID</td>
    	<td width="80">총재고수량</td>
    	<td width="100">소비자가<br>(현기준)</td>
    	<td width="80">매입마진</td>
    	<td width="100">매입가<br>(현기준)</td>
    	<% IF(isShowIpgoPrice)THEN %><td width="100">매입가<br>(매입시기준)</td><% end if %>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    
    if not IsNULL(ojaego.FItemList(i).FavgIpgoPriceSum) then
        totavgIpgoPrice = totavgIpgoPrice + ojaego.FItemList(i).FavgIpgoPriceSum
    end if
    
    iURL = "monthlystock_detail.asp?menupos="& menupos &"&mwgubun="& mwgubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&makerid="& ojaego.FItemList(i).FMakerid &"&newitem="& newitem &"&itemgubun="& itemgubun&"&vatyn="&vatyn
    if Not(isOnlySys) THEN iURL=iURL&"&sysorreal="& sysorreal 
    %>
    <% if  (isViewUser) or (ojaego.FItemList(i).FMakerUsing="Y") then %>
    <tr align="center" bgcolor="#FFFFFF">
    <% else %>
    <tr align="center" bgcolor="#CCCCCC">
    <% end if %>
    	<td align="left"><a href="<%= iURL %>" ><%= ojaego.FItemList(i).FMakerid %></a></td>
    	<td align="center"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td>
    		<% if ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<% IF(isShowIpgoPrice)THEN %>
    	<td align="right">
    	<% if IsNULL(ojaego.FItemList(i).FavgIpgoPriceSum) then %>
    	-
        <% else %>
    	<%= FormatNumber(ojaego.FItemList(i).FavgIpgoPriceSum,0) %>
    	<% end if %>
    	</td>
    	<% end if %>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>총계</td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<% IF(isShowIpgoPrice)THEN %><td align="right" ><%= FormatNumber(totavgIpgoPrice,0) %></td><% end if %>
    </tr>
</table>

<% end if %>

<%
set ojaego = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->