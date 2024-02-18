<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출
' History : 2009.04.07 서동석 생성
'			2010.03.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim shopid ,oldlist ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,yyyymmdd1,yyymmdd2 ,fromDate,toDate ,datefg
dim i ,totalsum, totalcount ,totalmileage, totalgainmileage ,TenGiftCardPaySum , totsellsumpro
dim fromDateolddatay, toDateolddatay ,olddatay ,ooffsell2 ,offgubun , reload , parameter, totmaechul
dim inc3pl
	olddatay = RequestCheckVar(request("olddatay"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	oldlist = requestCheckVar(request("oldlist"),10)
	datefg = requestCheckVar(request("datefg"),32)
	offgubun = requestCheckVar(request("offgubun"),10)
	reload = requestCheckVar(request("reload"),2)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"
if reload <> "on" and offgubun = "" then offgubun = "95"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now())))
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

fromDateolddatay = dateadd("m",-12,fromDate)
toDateolddatay = dateadd("m",-12,dateadd("d",-1,toDate))
	
dim ooffsell
set ooffsell = new COffShopSell
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOldData = oldlist
	ooffsell.frectdatefg = datefg
	ooffsell.frectoffgubun = offgubun
	ooffsell.FRectInc3pl = inc3pl
	ooffsell.GetOffSellByShop

totalsum = 0
totalcount = 0
totalmileage = 0
totalgainmileage = 0
TenGiftCardPaySum = 0
totsellsumpro = 0
totmaechul = 0

parameter = "oldlist="&oldlist&"&datefg="&datefg&"&offgubun="&offgubun&"&inc3pl="&inc3pl&"&menupos="&menupos
%>

<script language='javascript'>

function PopbrandSellSum(shopid,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
	var PopbrandSellSum = window.open('dailysellreport_detailbrand.asp?shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&<%=parameter%>','PopbrandSellSum','width=1024,height=768,scrollbars=yes,resizable=yes');
	PopbrandSellSum.focus();
}

function popitemdetail(yyyy1,mm1,dd1,yyyy2,mm2,dd2,shopid){
	var popitemdetail = window.open('/admin/offshop/todayselldetail.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&shopid='+shopid+'&<%=parameter%>','popitemdetail','width=1024,height=768,scrollbars=yes,resizable=yes');
	popitemdetail.focus();
}

function popjumundetail(yyyy1,mm1,dd1,yyyy2,mm2,dd2,shopid){
	var popjumundetail = window.open('/admin/offshop/todaysellmaster.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&shopid='+shopid+'&<%=parameter%>','popjumundetail','width=1024,height=768,scrollbars=yes,resizable=yes');
	popjumundetail.focus();
}

function frmsubmit(cholddatay){
	
	if(cholddatay=='RESETOLDDATAY'){
		frm.olddatay.value = '';
	}

	frm.submit();
}

function cholddatay(){
	//cholddatay = document.getElementsByName("cholddatay")
	
	if(frm.olddatay.value==''){
		frm.olddatay.value = 'ON';
	} else {
		frm.olddatay.value = '';
	}
	
	frmsubmit('');
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="reload" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="olddatay" value="<%= olddatay %>">
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
				* 매장구분 : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='frmsubmit(""RESETOLDDATAY"");'" %>
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>				
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('RESETOLDDATAY');">
	</td>
</tr>	
</form>
</table>
<!-- 표 상단바 끝-->
<Br>
<!-- 표 중간바 시작-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">        	
    </td>
    <td align="right">	       
    </td>        
</tr>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ooffsell.FResultCount %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>샾구분</td>
	<td>샾명</td>
	<td>주문건수</td>
	<td>매출액</td>
	<!--<td>마일리지사용</td>-->
	<td>매출액<Br>(마일리지포함)</td>
	<td>%</td>

	<% if not(C_IS_Maker_Upche) then %>
		<td>객단가</td>
	<% end if %>

	<!--<td>마일리지<br>적립</td>
	<td>기프트카드</td>-->
	<td>비고</td>
</tr>
<%
if ooffsell.FResultCount > 0 then

for i=0 to ooffsell.FResultCount -1

totalcount = totalcount + ooffsell.FItemList(i).FCount
totalsum = totalsum + ooffsell.FItemList(i).Fsellsum
totalmileage = totalmileage + ooffsell.FItemList(i).FSpendMile
totalgainmileage  = totalgainmileage + ooffsell.FItemList(i).FGainMile
TenGiftCardPaySum  = TenGiftCardPaySum + ooffsell.FItemList(i).fTenGiftCardPaySum
totmaechul = totmaechul + (ooffsell.FItemList(i).Fsellsum + ooffsell.FItemList(i).FSpendMile)

if ooffsell.FItemList(i).Fsellsum <> 0 and ooffsell.FItemList(i).Fsellsum <> "" and ooffsell.maxt <> 0 and ooffsell.maxt <> "" then
	totsellsumpro = totsellsumpro + round(ooffsell.FItemList(i).Fsellsum/ooffsell.maxt*100,1)
else
	totsellsumpro = 0
end if
%>
<tr bgcolor="#FFFFFF" height=24 align="center">
	<td><%= ooffsell.FItemList(i).Fshopid %></td>
	<td><%= ooffsell.FItemList(i).Fshopname %></td>
	<td><%= FormatNumber(ooffsell.FItemList(i).FCount,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Fsellsum,0) %></td>
	<!--<td align="right"><%'= FormatNumber(ooffsell.FItemList(i).FSpendMile,0) %></td>-->
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell.FItemList(i).Fsellsum + ooffsell.FItemList(i).FSpendMile,0) %></td>
	<td>
		<% if ooffsell.FItemList(i).Fsellsum <> 0 and ooffsell.FItemList(i).Fsellsum <> "" then %>
			<%= round(ooffsell.FItemList(i).Fsellsum/ooffsell.maxt*100,1) %> %
		<% else %>
			0 %
		<% end if %>
	</td>
	<!--<td align="right"><%'= FormatNumber(ooffsell.FItemList(i).FGainMile,0) %></td>
	<td align="right"><%'= FormatNumber(ooffsell.FItemList(i).fTenGiftCardPaySum,0) %></td>-->

	<% if not(C_IS_Maker_Upche) then %>
		<td align="right">
			<%
			if ooffsell.FItemList(i).Fsellsum <> 0 and ooffsell.FItemList(i).FCount <> 0 then
				response.write  FormatNumber(ooffsell.FItemList(i).Fsellsum / ooffsell.FItemList(i).FCount,0)
			else
				response.write "0"
			end if
			%>
		</td>
	<% end if %>

	<td width="300">
		<input type="button" onclick="popitemdetail('<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>','<%= ooffsell.FItemList(i).FShopid %>');" value="상품상세" class="button">
		
		<% if not(C_IS_Maker_Upche) then %> 
			<input type="button" onclick="popjumundetail('<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>','<%= ooffsell.FItemList(i).FShopid %>');" value="주문상세" class="button">
		<% end if %>
				
		<input type="button" onclick="PopbrandSellSum('<%= ooffsell.FItemList(i).Fshopid %>','<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>');" value="브랜드상세" class="button">
	</td>
</tr>
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td colspan=2>Total</td>
	<td><%= FormatNumber(totalcount,0) %></td>
	<td align="right"><%= FormatNumber(totalsum,0) %></td>
	<!--<td align="right"><%'= FormatNumber(totalmileage,0) %></td>-->
	<td align="right"><%= FormatNumber(totmaechul,0) %></td>
	<td><%= round(totsellsumpro,0) %> %</td>
	<!--<td align="right"><%'= FormatNumber(totalgainmileage,0) %></td>
	<td align="right"><%'= FormatNumber(TenGiftCardPaySum,0) %></td>-->

	<% if not(C_IS_Maker_Upche) then %>
		<td align="right">
			<%= FormatNumber(totalsum/totalcount,0) %>
		</td>
	<% end if %>

	<td>
	</td>
</tr>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="15">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>

<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="checkbox" name="cholddatay" <% if olddatay="ON" then response.write " checked" %> onclick='cholddatay();'>전년도 비교내역 표시( <%= fromDateolddatay%> - <%=toDateolddatay%> )
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<%
if olddatay = "ON" then

set ooffsell2 = new COffShopSell
	ooffsell2.FRectStartDay = fromDateolddatay
	ooffsell2.FRectEndDay = dateadd("d",+1,toDateolddatay)
	ooffsell2.FRectOldData = oldlist
	ooffsell2.frectdatefg = datefg
	ooffsell2.frectoffgubun = offgubun
	ooffsell2.FRectInc3pl = inc3pl	
	ooffsell2.GetOffSellByShop

totalsum = 0
totalcount = 0
totalmileage = 0
totalgainmileage = 0
TenGiftCardPaySum = 0
totsellsumpro = 0
totmaechul = 0
%>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ooffsell2.FResultCount %></b>
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td>샾구분</td>
		<td>샾명</td>
		<td>주문건수</td>
		<td>매출액</td>
		<!--<td>마일리지사용</td>-->
		<td>매출액<Br>(마일리지포함)</td>
		<td>%</td>
		<!--<td>마일리지<br>적립</td>
		<td>기프트카드</td>-->

		<% if not(C_IS_Maker_Upche) then %>
			<td>객단가</td>
		<% end if %>

		<td>비고</td>
	</tr>
	<%
	if ooffsell2.FResultCount > 0 then
	
	for i=0 to ooffsell2.FResultCount -1
	
	totalcount = totalcount + ooffsell2.FItemList(i).FCount
	totalsum = totalsum + ooffsell2.FItemList(i).Fsellsum
	totalmileage = totalmileage + ooffsell2.FItemList(i).FSpendMile
	totalgainmileage  = totalgainmileage + ooffsell2.FItemList(i).FGainMile
	TenGiftCardPaySum  = TenGiftCardPaySum + ooffsell2.FItemList(i).fTenGiftCardPaySum
	totmaechul = totmaechul + (ooffsell2.FItemList(i).Fsellsum + ooffsell2.FItemList(i).FSpendMile)
	
	if ooffsell2.FItemList(i).Fsellsum <> 0 and ooffsell2.FItemList(i).Fsellsum <> "" and ooffsell2.maxt <> 0 and ooffsell2.maxt <> "" then
		totsellsumpro = totsellsumpro + round(ooffsell2.FItemList(i).Fsellsum/ooffsell2.maxt*100,1)
	else
		totsellsumpro = 0
	end if
	%>
	<tr bgcolor="#FFFFFF" height=24 align="center">
		<td><%= ooffsell2.FItemList(i).Fshopid %></td>
		<td><%= ooffsell2.FItemList(i).Fshopname %></td>
		<td><%= FormatNumber(ooffsell2.FItemList(i).FCount,0) %></td>
		<td align="right"><%= FormatNumber(ooffsell2.FItemList(i).Fsellsum,0) %></td>
		<!--<td align="right"><%'= FormatNumber(ooffsell2.FItemList(i).FSpendMile,0) %></td>-->
		<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell2.FItemList(i).Fsellsum + ooffsell2.FItemList(i).FSpendMile,0) %></td>
		<td>
			<% if ooffsell2.FItemList(i).Fsellsum <> 0 and ooffsell2.FItemList(i).Fsellsum <> "" then %>
				<%= round(ooffsell2.FItemList(i).Fsellsum/ooffsell2.maxt*100,1) %> %
			<% else %>
				0 %
			<% end if %>
		</td>
		<!--<td align="right"><%'= FormatNumber(ooffsell2.FItemList(i).FGainMile,0) %></td>
		<td align="right"><%'= FormatNumber(ooffsell2.FItemList(i).fTenGiftCardPaySum,0) %></td>-->

		<% if not(C_IS_Maker_Upche) then %>
			<td align="right">
				<%
				if ooffsell2.FItemList(i).Fsellsum <> 0 and ooffsell2.FItemList(i).FCount <> 0 then
					response.write  FormatNumber(ooffsell2.FItemList(i).Fsellsum / ooffsell2.FItemList(i).FCount,0)
				else
					response.write "0"
				end if
				%>
			</td>
		<% end if %>

		<td width="300">
			<input type="button" onclick="popitemdetail('<%= left(fromDateolddatay,4) %>','<%= Mid(fromDateolddatay,6,2) %>','<%= Mid(fromDateolddatay,9,2) %>','<%= left(toDateolddatay,4) %>','<%= Mid(toDateolddatay,6,2) %>','<%= Mid(toDateolddatay,9,2) %>','<%= ooffsell2.FItemList(i).FShopid %>');" value="상품상세" class="button">
			
			<% if not(C_IS_Maker_Upche) then %> 
				<input type="button" onclick="popjumundetail('<%= left(fromDateolddatay,4) %>','<%= Mid(fromDateolddatay,6,2) %>','<%= Mid(fromDateolddatay,9,2) %>','<%= left(toDateolddatay,4) %>','<%= Mid(toDateolddatay,6,2) %>','<%= Mid(toDateolddatay,9,2) %>','<%= ooffsell2.FItemList(i).FShopid %>');" value="주문상세" class="button">
			<% end if %>
					
			<input type="button" onclick="PopbrandSellSum('<%= ooffsell2.FItemList(i).Fshopid %>','<%= left(fromDateolddatay,4) %>','<%= Mid(fromDateolddatay,6,2) %>','<%= Mid(fromDateolddatay,9,2) %>','<%= left(toDateolddatay,4) %>','<%= Mid(toDateolddatay,6,2) %>','<%= Mid(toDateolddatay,9,2) %>');" value="브랜드상세" class="button">			
		</td>
	</tr>
	<% next %>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td colspan=2>Total</td>
		<td><%= FormatNumber(totalcount,0) %></td>
		<td align="right"><%= FormatNumber(totalsum,0) %></td>
		<!--<td align="right"><%'= FormatNumber(totalmileage,0) %></td>-->
		<td align="right"><%= FormatNumber(totmaechul,0) %></td>
		<td><%= round(totsellsumpro,0) %> %</td>
		<!--<td align="right"><%'= FormatNumber(totalgainmileage,0) %></td>
		<td align="right"><%'= FormatNumber(TenGiftCardPaySum,0) %></td>-->

		<% if not(C_IS_Maker_Upche) then %>
			<td align="right">
				<%= FormatNumber(totalsum/totalcount,0) %>
			</td>
		<% end if %>

		<td>
		</td>
	</tr>
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="15">등록된 내용이 없습니다.</td>
	</tr>
	<% end if %>
	</table>
<% end if %>

<%
set ooffsell = Nothing
set ooffsell2 = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->