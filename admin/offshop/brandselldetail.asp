<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출
' History : 2009.04.07 서동석 생성
'			2010.04.27 한용민 수정
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
dim shopid, designer, page ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,yyyymmdd1,yyymmdd2 ,fromDate,toDate ,offgubun , oldlist
dim menupos ,ooffsell ,i, totalsum, totcnt ,datefg, vPurchaseType
	shopid = requestCheckVar(request("shopid"),32)
	designer = requestCheckVar(request("designer"),32)
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
	menupos = requestCheckVar(request("menupos"),10)
	datefg = requestCheckVar(request("datefg"),32)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)

if datefg = "" then datefg = "maechul"
			
if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-14)
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
	
	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		designer = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

set ooffsell = new COffShopSellReport
	ooffsell.FRectShopid = shopid
	ooffsell.FRectDesigner = designer
	ooffsell.FRectNormalOnly = "on"
	ooffsell.frectdatefg = datefg
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOffgubun = offgubun
	ooffsell.FRectOldData = oldlist
	ooffsell.FRectBrandPurchaseType = vPurchaseType		
	ooffsell.GetDaylySellItemList

totalsum = 0
totcnt = 0
%>
<br>

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

<table width="100%" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ooffsell.FResultCount %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td width="86">바코드</td>
	<td width="90">브랜드</td>
	<td width="100">상품명</td>
	<td width="100">옵션명</td>
	<td width="70">소비자가</td>
	<td width="70">판매가</td>
	<td width="60">갯수</td>
	<td width="80">합계</td>
</tr>
<% 
for i=0 to ooffsell.FresultCount-1

totcnt = totcnt + ooffsell.FItemList(i).FItemNo
totalsum = totalsum + ooffsell.FItemList(i).FSubTotal
%>
<tr bgcolor="#FFFFFF" height=24>
	<td><%= ooffsell.FItemList(i).GetBarCode %></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>
	<td><%= ooffsell.FItemList(i).FItemName %></td>
	<td><%= ooffsell.FItemList(i).FItemOptionName %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></td>
	<td align="center"><%= ooffsell.FItemList(i).FItemNo %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSubTotal,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="6"><b>총계</b></td>
	<td><%= totcnt %></td>
	<td align="right"><b><%= FormatNumber(totalsum,0) %></b></td>
</tr>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->