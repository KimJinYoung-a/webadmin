<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 브랜드 매장별 매출
' History : 2012.03.28 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim page,shopid ,yyyymmdd1,yyymmdd2 ,offgubun ,oldlist ,fromDate,toDate ,yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim i, sum1, sum2, sum3 ,makerid ,datefg , parameter ,CurrencyUnit, CurrencyChar, ExchangeRate ,FmNum
dim menupos, vPurchaseType ,reload, buyergubun, inc3pl
	menupos = requestCheckVar(request("menupos"),10)
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
	if page="" then page=1
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	offgubun = requestCheckVar(request("offgubun"),10)
	oldlist = requestCheckVar(request("oldlist"),10)
	makerid = requestCheckVar(request("makerid"),32)
	datefg = requestCheckVar(request("datefg"),32)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	reload = requestCheckVar(request("reload"),2)
	buyergubun = requestCheckVar(request("buyergubun"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if datefg = "" then datefg = "maechul"	
if reload <> "on" and offgubun = "" then offgubun = "95"
	
sum1 =0
sum2 =0
sum3 =0

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/매장
if (C_IS_SHOP) then
	
	'//직영점일때
	if C_IS_OWN_SHOP then
		
		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if		
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

if shopid<>"" then offgubun=""

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopID = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOffgubun = offgubun
	ooffsell.FRectOldData = oldlist
	ooffsell.frectmakerid = makerid
	ooffsell.frectdatefg = datefg
	ooffsell.FCurrPage = page
	ooffsell.Fpagesize=1000
	ooffsell.FRectBrandPurchaseType = vPurchaseType
	ooffsell.frectbuyergubun = buyergubun
	ooffsell.FRectInc3pl = inc3pl	
	ooffsell.GetBrandshopSell

Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
FmNum = CHKIIF(CurrencyUnit="WON" or CurrencyUnit="KRW",0,2)

parameter = "menupos="& menupos &"&datefg="& datefg &"&offgubun="& offgubun &"&oldlist="& oldlist &"&purchasetype="& vPurchaseType &"&buyergubun="& buyergubun
%>

<script language="javascript">
	
function pop_category(makerid,yyyy1,mm1,dd1,yyyy2,mm2,dd2,shopid){
	var pop_category = window.open('/admin/offshop/offshop_categorysellsum.asp?shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&makerid='+makerid+'&<%=parameter%>','pop_category','width=1024,height=768,scrollbars=yes,resizable=yes');
    pop_category.focus();
}

function pop_detail(makerid,yyyy1,mm1,dd1,yyyy2,mm2,dd2,shopid){
	var pop_detail = window.open('/admin/offshop/todayselldetail.asp?shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&makerid='+makerid+'&<%=parameter%>','pop_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
	//var pop_detail = window.open('/admin/offshop/brandselldetail.asp?shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&designer='+makerid+'&<%=parameter%>','pop_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
    pop_detail.focus();
}

//function pop_stock(makerid,shopid){
//	var pop_stock = window.open('/admin/offshop/jaegolist.asp?shopid='+shopid+'&makerid='+makerid+'&<%=parameter%>','pop_stock','width=1024,height=768,scrollbars=yes,resizable=yes');
//    pop_stock.focus();
//}

function frmsubmit(){
	frm.submit();
}

</script>
	
<!-- 표 상단바 시작-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="reload" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 : <% drawmaechul_datefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3년이전
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>	
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
				<p>
				* 매장 구분 :<% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='frmsubmit();'" %>
				&nbsp;&nbsp;
				* 구매유형 : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
				&nbsp;&nbsp;
				* 국적구분: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
				<p>
				* 브랜드:<% drawSelectBoxDesignerwithName "makerid",makerid %>
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>				
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onclick="frmsubmit();">
	</td>
</tr>
</table>
<!-- 표 상단바 끝-->
<br>
<!-- 표 중간바 시작-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">        	
    </td>
    <td align="right">
    </td>        
</tr>
</form>
</table>
<!-- 표 중간바 끝-->

<table width="100%" cellspacing="1" cellpadding="3" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ooffsell.FResultCount %></b> ※ 최대 1000건 까지 검색됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>매장</td>
	<td>브랜드</td>
	<td>아이템건수</td>
	<td>총매출액</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>총매입가</td>
		<td>수익</td>
		<td>마진율</td>
	<% end if %>
	
	<td>비고</td>
</tr>
<% 
for i=0 to ooffsell.FresultCount-1

sum1 = sum1 + ooffsell.FItemList(i).FSum

if ooffsell.FItemList(i).FChargeDiv="6" then
	sum2 = sum2 + ooffsell.FItemList(i).FSum
else
	sum3 = sum3 + ooffsell.FItemList(i).FSum
end if
%>
<tr bgcolor="#FFFFFF" align="center">
	<td align="center"><%= ooffsell.FItemList(i).fshopname %></td>
	
	<% if ooffsell.FItemList(i).FChargeDiv="6" then %>
		<td><b><font color="#3333CC"><%= ooffsell.FItemList(i).FMakerid %></font></b></td>
	<% else %>
		<td><%= ooffsell.FItemList(i).FMakerid %></td>
	<% end if %>
	
	<td align="center"><%= ooffsell.FItemList(i).FCount %></td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell.FItemList(i).FSum,0) %></td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).fsuplyprice,0) %></td>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).fprofit,0) %></td>
		<td align="right">
			<% if ooffsell.FItemList(i).fsuplyprice > 0 and ooffsell.FItemList(i).FSum > 0 then %>
				<%= FormatNumber(100-ooffsell.FItemList(i).fsuplyprice/ooffsell.FItemList(i).FSum*100,0) %>%
			<% else %>
				0%
			<% end if %>
		</td>
	<% end if %>
	
	<td width=180>
		<input type="button" onclick="pop_detail('<%= ooffsell.FItemList(i).FMakerid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%= ooffsell.FItemList(i).fshopid %>');" value="상품별" class="button">
		<input type="button" onclick="pop_category('<%= ooffsell.FItemList(i).FMakerid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%= ooffsell.FItemList(i).fshopid %>');" value="카테고리별" class="button">

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<!--<input type="button" onclick="pop_stock('<%'= ooffsell.FItemList(i).FMakerid %>','<%'= ooffsell.FItemList(i).fshopid %>');" value="예상재고" class="button">-->
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height=24 align="center">
	<td colspan=4>합계</td>
	<td colspan="10" align="right">
		<b><font color="#3333CC">업체위탁 : </font></b><%= FormatNumber(sum2,0) %>
		<br>일반 : <%= FormatNumber(sum3,0) %>
		<br>Total : <%= FormatNumber(sum1,0) %>
	</td>
</tr>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->