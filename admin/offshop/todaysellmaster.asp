<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출 주문건 상세 공용페이지 NO 페이징 버전
' History : 2009.04.07 서동석 생성
'			2010.03.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim shopid, oldlist , datefg , prejumunno , makerid , menupos ,yyyy1,mm1,dd1,yyyy2,mm2,dd2, toDate,fromDate
dim cardMinusTotal, cashMinusTotal, cardMinusCnt, cashMinusCnt, buyergubun, inc3pl
dim etcTotal, etcCnt, etcMinusTotal, etcMinusCnt ,i,totalsum ,cardtotal, cashtotal, cardcnt, cashcnt
dim extTotal, extCnt, extMinusTotal, extMinusCnt
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	menupos = requestCheckVar(request("menupos"),10)
	shopid = requestCheckVar(request("shopid"),32)
	oldlist = requestCheckVar(request("oldlist"),10)
	datefg = requestCheckVar(request("datefg"),32)
	makerid = requestCheckVar(request("makerid"),32)
	buyergubun = requestCheckVar(request("buyergubun"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"

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
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

''두타쪽 매출조회 권한 
Dim isFixShopView
IF (session("ssBctID")="doota01") then 
    shopid="streetshop014"
    C_IS_SHOP = TRUE
    isFixShopView = TRUE
ENd If
		
dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopid = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectOldData = oldlist
    ooffsell.FRectTerms = ""
    ooffsell.FRectStartDay = fromDate
    ooffsell.frectdatefg = datefg
    ooffsell.FRectDesigner = makerid
    ooffsell.FRectEndDay = toDate
    ooffsell.FRectbuyergubun = buyergubun
	ooffsell.FRectInc3pl = inc3pl    
	ooffsell.GetDaylySellJumunList

totalsum =0
cardtotal =0
cashtotal =0
cardcnt   =0
cashcnt   =0
cardMinusTotal =0
cashMinusTotal =0
cardMinusCnt   =0
cashMinusCnt   =0
etcTotal        =0
etcCnt          =0
etcMinusTotal   =0
etcMinusCnt     =0
extTotal        =0
extCnt          =0
extMinusTotal   =0
extMinusCnt     =0
%>

<script type='text/javascript'>

function frmsubmit(){

	frm.submit();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 :
				<% drawmaechuldatefg "datefg" ,datefg ,""%> 
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>	
					<% if (not C_IS_OWN_SHOP and shopid <> "") or (isFixShopView) then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>				
				<p>
				<% if C_IS_Maker_Upche then %>
					* 브랜드 : <%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
				&nbsp;&nbsp;
				* 국적구분: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>				
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>	
</form>
</table>
<!-- 표 상단바 끝-->

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">        	
    </td>
    <td align="right">	       
    </td>        
</tr>	
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		검색결과 : <b><%=ooffsell.FresultCount%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>
		<% if datefg = "maechul" then %>
			<%= chkIIF(shopid="cafe002","매출일","주문번호") %>
		<% else %>
			<%= chkIIF(shopid="cafe002","주문일","주문번호") %>
		<% end if %>	
	</td>
	<td></td>
	<td>상품명</td>
	<td>판매가</td>
	<td>결제금액</td>
	<td>갯수</td>
	
	<% if shopid<>"cafe002" then %>
		<td>샾주문일</td>
	<% end if %>
</tr>
<%
if ooffsell.FresultCount > 0 then
	
for i=0 to ooffsell.FresultCount-1

if prejumunno<>ooffsell.FItemList(i).ForderNo then
	
	totalsum = totalsum + ooffsell.FItemList(i).Frealsum
	if (ooffsell.FItemList(i).Fcardsum>0) then
        cardtotal = cardtotal + ooffsell.FItemList(i).Fcardsum
        cardcnt   = cardcnt + 1
    elseif (ooffsell.FItemList(i).Fcardsum<0) then
        cardMinusTotal = cardMinusTotal + ooffsell.FItemList(i).Fcardsum
        cardMinusCnt   =cardMinusCnt + 1
    end if
    
    if (ooffsell.FItemList(i).Fcashsum>0) then
        cashtotal = cashtotal + ooffsell.FItemList(i).Fcashsum
        cashcnt   = cashcnt + 1
    elseif (ooffsell.FItemList(i).Fcashsum<0) then
        cashMinusTotal = cashMinusTotal + ooffsell.FItemList(i).Fcashsum
        cashMinusCnt   =cashMinusCnt + 1
    end if
    
    if (ooffsell.FItemList(i).FgiftcardPaysum>0) then
        etcTotal = etcTotal + ooffsell.FItemList(i).FgiftcardPaysum
        etcCnt   = etcCnt + 1
    elseif (ooffsell.FItemList(i).FgiftcardPaysum<0) then
        etcMinusTotal = etcMinusTotal + ooffsell.FItemList(i).FgiftcardPaysum
        etcMinusCnt   =etcMinusCnt + 1
    end if

    if (ooffsell.FItemList(i).FextPaysum>0) then
        extTotal = extTotal + ooffsell.FItemList(i).FextPaysum
        extCnt   = extCnt + 1
    elseif (ooffsell.FItemList(i).FextPaysum<0) then
        extMinusTotal = extMinusTotal + ooffsell.FItemList(i).FextPaysum
        extMinusCnt   =extMinusCnt + 1
    end if

prejumunno = ooffsell.FItemList(i).ForderNo
%>
<tr bgcolor="#EEEEEE" align="center">
	<td align="center"><%= chkIIF(shopid="cafe002",ooffsell.FItemList(i).Fshopregdate,ooffsell.FItemList(i).ForderNo) %></td>
	<td><font color="<%= ooffsell.FItemList(i).JumunMethodColor %>"><%= ooffsell.FItemList(i).JumunMethodName %></font></td>
	<td><%= ooffsell.FItemList(i).Fpointuserno %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Ftotalsum,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Frealsum,0) %></td>
	<td align="center"></td>
	
	<% if shopid<>"cafe002" then %>
		<td><%= ooffsell.FItemList(i).Fshopregdate %></td>
	<% end if %>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" align="center">
	<td></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>
	<td><%= ooffsell.FItemList(i).FItemName %> <%= ooffsell.FItemList(i).FItemOptionName %></td>
	
	<% if ooffsell.FItemList(i).FItemNo<0 then %>
		<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></font></td>
		<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></font></td>
		<td align="center"><font color=red><%= ooffsell.FItemList(i).FItemNo %></font></td>
	<% else %>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></td>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></td>
		<td align="center"><%= ooffsell.FItemList(i).FItemNo %></td>
	<% end if %>
	
	<% if shopid<>"cafe002" then %>
		<td align="right"></td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2"><b>총계</b></td>
	<td colspan="6" align="right">
		<table width=440 border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
		    <td>현금 :</td>
		    <td align="right"><%= FormatNumber(cashtotal,0) %> 원</td>
		    <td align="center">(<%= FormatNumber(cashcnt,0) %> 건)</td>
		    <td width=10></td>
		    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal,0) %> 원</td>
		    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt,0) %> 건)</font></td>
		    <td align="right"><%= FormatNumber(cashtotal + cashMinusTotal,0) %> 원</td>
		</tr>
		<tr>
		    <td>카드 :</td>
		    <td align="right"><%= FormatNumber(cardtotal,0) %> 원</td>
		    <td align="center">(<%= FormatNumber(cardcnt,0) %> 건)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(cardMinusTotal,0) %> 원</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(cardMinusCnt,0) %> 건)</font></td>
		    <td align="right"><%= FormatNumber(cardtotal + cardMinusTotal,0) %> 원</td>
		</tr>
		<tr>
		    <td>상품권 :</td>
		    <td align="right"><%= FormatNumber(etcTotal,0) %> 원</td>
		    <td align="center">(<%= FormatNumber(etccnt,0) %> 건)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(etcMinusTotal,0) %> 원</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(etcMinusCnt,0) %> 건)</font></td>
		    <td align="right"><%= FormatNumber(etcTotal + etcMinusTotal,0) %> 원</td>
		</tr>
		<tr>
		    <td>기타결제 :</td>
		    <td align="right"><%= FormatNumber(extTotal,0) %> 원</td>
		    <td align="center">(<%= FormatNumber(extcnt,0) %> 건)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(extMinusTotal,0) %> 원</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(extMinusCnt,0) %> 건)</font></td>
		    <td align="right"><%= FormatNumber(extTotal + extMinusTotal,0) %> 원</td>
		</tr>
		<tr>
		    <td>합계 :</td>
		    <td align="right"><%= FormatNumber(cashtotal + cardtotal + etcTotal + extTotal,0) %> 원</td>
		    <td align="center">(<%= FormatNumber(cashcnt + cardcnt + etccnt + extcnt,0) %> 건)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal + cardMinusTotal + etcMinusTotal + extMinusTotal,0) %> 원</td>
		    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt + cardMinusCnt + etcMinusCnt + extMinusCnt,0) %> 건)</font></td>
		    <td align="right"><%= FormatNumber(totalsum,0) %> 원</td>
		</tr>
		</table>
	</td>
</tr>
<% else %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="20">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->