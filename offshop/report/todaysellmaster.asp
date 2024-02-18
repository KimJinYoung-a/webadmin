<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim chargeid, shopid, terms, datefg, oldlist
chargeid = request("chargeid")
shopid = request("shopid")
terms = request("terms")
datefg = request("datefg")
oldlist = request("oldlist")
if datefg = "" then datefg = "maechul"
	
shopid = session("ssBctID") ''강제지정
if (shopid="doota01") then shopid="streetshop014"

dim ooffsell
set ooffsell = new COffShopSellReport
'ooffsell.FRectShopid = shopid
'ooffsell.FRectNormalOnly = "on"
'ooffsell.FRectTerms = terms

''ooffsell.GetDaylySellJumunList

ooffsell.FRectShopid = shopid
ooffsell.FRectNormalOnly = "on"
ooffsell.FRectOldData = oldlist
ooffsell.FRectTerms = ""
ooffsell.FRectStartDay = terms
ooffsell.frectdatefg = datefg
ooffsell.FRectEndDay = CStr(dateAdd("d",1,terms))
ooffsell.GetDaylySellJumunList

dim i,totalsum
dim cardtotal, cashtotal, cardcnt, cashcnt
dim cardMinusTotal, cashMinusTotal, cardMinusCnt, cashMinusCnt
dim etcTotal, etcCnt, etcMinusTotal, etcMinusCnt
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
dim debittotal,debitcnt, debitMinusTotal,debitMinusCnt
    debittotal  =0
    debitcnt    =0
    debitMinusTotal=0
    debitMinusCnt=0
    

dim prejumunno

Dim CurrencyUnit, CurrencyChar, ExchangeRate
Dim FmNum, IsTaxAddCharge
Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
FmNum = CHKIIF(CurrencyUnit="WON" or CurrencyUnit="KRW",0,2)

IsTaxAddCharge = CHKIIF(CurrencyUnit<>"WON" and CurrencyUnit<>"KRW",true,false)
%>
<table width="100%" cellspacing="1" cellpadding="0" class="a" bgcolor=#3d3d3d>
<tr>
	<td width="100" bgcolor="#DDDDFF">기간</td>
	<td bgcolor="#FFFFFF"><%= terms %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">샾 구분</td>
	<td bgcolor="#FFFFFF"><%= shopid %></td>
</tr>
</table>
<br>
<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align="center">
	<td width="86">주문번호</td>
	<td width="90"></td>
	<td>상품명</td>
	<td width="60">판매가</td>
	<td width="60">결제금액</td>
	<% if IsTaxAddCharge then %>
	<td width="60">TAX</td>
	<% end if %>
	<td width="40">갯수</td>
	<td width="100">샾주문일</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>

<% if prejumunno<>ooffsell.FItemList(i).ForderNo then %>
<%
	totalsum = totalsum + ooffsell.FItemList(i).Frealsum
	if (ooffsell.FItemList(i).Fcardsum>0) then
        cardtotal = cardtotal + ooffsell.FItemList(i).Fcardsum
        cardcnt   = cardcnt + 1
        
        if (ooffsell.FItemList(i).Fjumunmethod="06") or (ooffsell.FItemList(i).Fjumunmethod="07") then
            debittotal = debittotal + ooffsell.FItemList(i).Fcardsum
            debitcnt   = debitcnt+ 1
        end if
        
    elseif (ooffsell.FItemList(i).Fcardsum<0) then
        cardMinusTotal = cardMinusTotal + ooffsell.FItemList(i).Fcardsum
        cardMinusCnt   =cardMinusCnt + 1
        
        if (ooffsell.FItemList(i).Fjumunmethod="06") or (ooffsell.FItemList(i).Fjumunmethod="07") then
            debitMinusTotal = debitMinusTotal + ooffsell.FItemList(i).Fcardsum
            debitMinusCnt   = debitMinusCnt+ 1
        end if
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
%>
<tr bgcolor="#EEEEEE">
	<td><%= ooffsell.FItemList(i).ForderNo %></td>
	<td><font color="<%= ooffsell.FItemList(i).JumunMethodColor %>"><%= ooffsell.FItemList(i).JumunMethodName %></font></td>
	<td></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Ftotalsum,FmNum) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Frealsum,FmNum) %></td>
	<% if IsTaxAddCharge then %>
	<td align="right"></td>
	<% end if %>
	<td align="center"></td>
	<td align="right"><%= ooffsell.FItemList(i).Fshopregdate %></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>
	<td><%= ooffsell.FItemList(i).FItemName %> <%= ooffsell.FItemList(i).FItemOptionName %></td>
	<% if ooffsell.FItemList(i).FItemNo<0 then %>
	<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,FmNum) %></font></td>
	<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,FmNum) %></font></td>
	<% if IsTaxAddCharge then %>
	<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FaddTaxCharge,FmNum) %></font></td>
	<% end if %>
	<td align="center"><font color=red><%= ooffsell.FItemList(i).FItemNo %></font></td>
	<% else %>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,FmNum) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,FmNum) %></td>
	<% if IsTaxAddCharge then %>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FaddTaxCharge,FmNum) %></td>
	<% end if %>
	<td align="center"><%= ooffsell.FItemList(i).FItemNo %></td>
	<% end if %>
	<td align="right"></td>
</tr>
<%
prejumunno=ooffsell.FItemList(i).ForderNo
%>
<% next %>
<tr bgcolor="#FFFFFF">
	<td><b>총계</b></td>
	<td colspan="7" align="right">
	<table width=440 border=0 cellspacing=0 cellpadding=0 class="a">
	<tr>
	    <td>현금 :</td>
	    <td align="right"><%= FormatNumber(cashtotal,FmNum) %> <%= CurrencyChar %></td>
	    <td align="center">(<%= FormatNumber(cashcnt,0) %> 건)</td>
	    <td width=10></td>
	    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal,FmNum) %> <%= CurrencyChar %></td>
	    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt,0) %> 건)</font></td>
	    <td align="right"><%= FormatNumber(cashtotal + cashMinusTotal,FmNum) %> <%= CurrencyChar %></td>
	</tr>
	<tr>
	    <td>카드 :</td>
	    <td align="right"><%= FormatNumber(cardtotal,FmNum) %> <%= CurrencyChar %></td>
	    <td align="center">(<%= FormatNumber(cardcnt,0) %> 건)</td>
	    <td></td>
	    <td align="right"><font color="red"><%= FormatNumber(cardMinusTotal,FmNum) %> <%= CurrencyChar %></font></td>
	    <td align="center"><font color="red">(<%= FormatNumber(cardMinusCnt,0) %> 건)</font></td>
	    <td align="right"><%= FormatNumber(cardtotal + cardMinusTotal,FmNum) %> <%= CurrencyChar %></td>
	</tr>
	<tr>
	    <td align="right">CreditCard :</td>
	    <td align="right"><%= FormatNumber(cardtotal-debittotal,FmNum) %> <%= CurrencyChar %></td>
	    <td align="center">(<%= FormatNumber(cardcnt-debitcnt,0) %> 건)</td>
	    <td></td>
	    <td align="right"><font color="red"><%= FormatNumber(cardMinusTotal-debitMinusTotal,FmNum) %> <%= CurrencyChar %></font></td>
	    <td align="center"><font color="red">(<%= FormatNumber(cardMinusCnt-debitMinusCnt,0) %> 건)</font></td>
	    <td align="right"><%= FormatNumber(cardtotal + cardMinusTotal - (debittotal+debitMinusTotal),FmNum) %> <%= CurrencyChar %></td>
	</tr>
	<tr>
	    <td align="right">Debit :</td>
	    <td align="right"><%= FormatNumber(debittotal,FmNum) %> <%= CurrencyChar %></td>
	    <td align="center">(<%= FormatNumber(debitcnt,0) %> 건)</td>
	    <td></td>
	    <td align="right"><font color="red"><%= FormatNumber(debitMinusTotal,FmNum) %> <%= CurrencyChar %></font></td>
	    <td align="center"><font color="red">(<%= FormatNumber(debitMinusCnt,0) %> 건)</font></td>
	    <td align="right"><%= FormatNumber(debittotal + debitMinusTotal,FmNum) %> <%= CurrencyChar %></td>
	</tr>
	<tr>
	    <td>상품권 :</td>
	    <td align="right"><%= FormatNumber(etcTotal,FmNum) %> <%= CurrencyChar %></td>
	    <td align="center">(<%= FormatNumber(etccnt,0) %> 건)</td>
	    <td></td>
	    <td align="right"><font color="red"><%= FormatNumber(etcMinusTotal,FmNum) %> <%= CurrencyChar %></font></td>
	    <td align="center"><font color="red">(<%= FormatNumber(etcMinusCnt,0) %> 건)</font></td>
	    <td align="right"><%= FormatNumber(etcTotal + etcMinusTotal,0) %> <%= CurrencyChar %></td>
	</tr>
	<tr>
	    <td>합계 :</td>
	    <td align="right"><%= FormatNumber(cashtotal + cardtotal + etcTotal,FmNum) %> <%= CurrencyChar %></td>
	    <td align="center">(<%= FormatNumber(cashcnt + cardcnt + etccnt,0) %> 건)</td>
	    <td></td>
	    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal + cardMinusTotal + etcMinusTotal,FmNum) %> <%= CurrencyChar %></td>
	    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt + cardMinusCnt + etcMinusCnt,0) %> 건)</font></td>
	    <td align="right"><%= FormatNumber(totalsum,FmNum) %> <%= CurrencyChar %></td>
	</tr>
	</td>
</tr>
</table>
<%
set ooffsell = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->