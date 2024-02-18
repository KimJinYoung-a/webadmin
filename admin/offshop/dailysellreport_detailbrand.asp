<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출
' History : 2010.06.08 한용민 생성
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
dim shopid, oldlist, yyyy1, mm1, dd1, yyyy2, mm2, dd2, yyyymmdd1, yyymmdd2, fromDate, toDate, offgubun
dim datefg, page, parameter, totsellprice, totrealprice, totsuplyprice, totitemno, reload, inc3pl
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	oldlist = requestCheckVar(request("oldlist"),10)
	datefg = requestCheckVar(request("datefg"),32)
	if datefg = "" then datefg = "maechul"
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
	offgubun = requestCheckVar(request("offgubun"),10)
	reload = requestCheckVar(request("reload"),2)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if page = "" then page = 1

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
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

parameter = "shopid="&shopid&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&oldlist="&oldlist&"&datefg="&datefg&"&menupos="&menupos&""

dim ooffsell
set ooffsell = new COffShopSell
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOldData = oldlist
	ooffsell.frectdatefg = datefg
	ooffsell.frectshopid = shopid
	ooffsell.frectoffgubun = offgubun	
	ooffsell.FPageSize = 1000
	ooffsell.FCurrPage = page
	ooffsell.FRectInc3pl = inc3pl	
	ooffsell.GetOffSellByShop_brand

dim i ,totalsum, totalcount ,totalmileage, totalgainmileage ,sellpro, countpro
totalsum = 0
totalcount = 0
totalmileage = 0
totalgainmileage = 0

for i=0 to ooffsell.FResultCount -1
	totalcount = totalcount + ooffsell.FItemList(i).FCount
	totalsum = totalsum + ooffsell.FItemList(i).Fsellsum
	totalmileage = totalmileage + ooffsell.FItemList(i).FSpendMile
	totalgainmileage  = totalgainmileage + ooffsell.FItemList(i).FGainMile
next

totsellprice = 0
totrealprice = 0
totsuplyprice = 0
totitemno = 0
%>

<script language='javascript'>

function frmsubmit(){

	frm.submit();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="reload" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 :
				<% drawmaechul_datefg "datefg" ,datefg ,""%> 
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
						* 매장 : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,11", "", " onchange='frmsubmit();'" %>
					<% end if %>
				<% else %>
					* 매장 : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,11", "", " onchange='frmsubmit();'" %>
				<% end if %>
				<p>
				* 매장구분 : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='frmsubmit();'" %>
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>				
			</td>
		</tr>
		</table> 
    </td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>	
</form>
</table>
<!-- 표 상단바 끝-->

<br>

<!-- 표 중간바 시작-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
    <tr valign="bottom">       
        <td align="left">
        	※ 마일리지 사용금액은 제외 됩니다.      	
	    </td>
	    <td align="right">	       
        </td>        
	</tr>	
</table>
<!-- 표 중간바 끝-->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ooffsell.FTotalCount %></b>
		&nbsp;
		※ 총 1000건 까지 표기 됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">브랜드</td>
	<td align="center">판매액</td>
	<td align="center">매출액</td>
	<td align="center">매입액</td>
	<td align="center">판매수량</td>	
</tr>
<%
if ooffsell.FresultCount>0 then
	
for i=0 to ooffsell.FresultCount-1

totsellprice = totsellprice + ooffsell.FItemList(i).fsellprice
totrealprice = totrealprice + ooffsell.FItemList(i).frealsellprice
totsuplyprice = totsuplyprice + ooffsell.FItemList(i).fsuplyprice
totitemno = totitemno + ooffsell.FItemList(i).fitemno
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
		
	<td align="center">
		<%= ooffsell.FItemList(i).fmakerid %>
	</td>
	<td align="center">
		<%= FormatNumber(ooffsell.FItemList(i).fsellprice,0) %>
	</td>
	<td align="center" bgcolor="#E6B9B8">
		<%= FormatNumber(ooffsell.FItemList(i).frealsellprice,0) %>
	</td>
	<td align="center">
		<%= FormatNumber(ooffsell.FItemList(i).fsuplyprice,0) %>
	</td>
	<td align="center">
		<%= ooffsell.FItemList(i).fitemno %>
	</td>
</tr>   
<% next %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">합계</td>
	<td align="center"><%= FormatNumber(totsellprice,0) %></td>
	<td align="center"><%= FormatNumber(totrealprice,0) %></td>
	<td align="center"><%= FormatNumber(totsuplyprice,0) %></td>
	<td align="center"><%= totitemno %></td>
</tr>	
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->