<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 매출
' History : 2010.06.08 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim shopid ,oldlist ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,yyyymmdd1,yyymmdd2 ,fromDate,toDate
dim datefg , page , parameter , totsellprice , totrealprice , totsuplyprice ,totitemno
dim Term,totsum , makerid
	makerid = requestCheckVar(request("jungsanid"),32)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	Term = requestCheckVar(request("Term"),10)
	oldlist = requestCheckVar(request("oldlist"),2)
	datefg = requestCheckVar(request("datefg"),32)
	if datefg = "" then datefg = "maechul"
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
	if page = "" then page = 1

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

if (C_IS_Maker_Upche) then
	makerid = session("ssBctID")
end if

parameter = "shopid="&shopid&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&oldlist="&oldlist&"&datefg="&datefg&"&menupos="&menupos&""

dim ooffsell
set ooffsell = new COffShopSell
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOldData = oldlist
	ooffsell.frectdatefg = datefg
	ooffsell.frectmakerid = makerid
	ooffsell.frectshopid = shopid
	ooffsell.FPageSize = 1000
	ooffsell.FCurrPage = page
	ooffsell.frectTerm = Term
	ooffsell.GetOffSellByShop_item

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
totsum = 0
%>

<!-- 표 상단바 시작-->
<table width="100%" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="makerid" value="<%= makerid %>">

<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>		
			<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3개월이전내역		
				매출기준 :
				<% drawmaechul_datefg "datefg" ,datefg ,""%> 
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>		
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>	
</form>
</table>
<!-- 표 상단바 끝-->

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

<br>
※ 마일리지 사용금액은 제외 됩니다.
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ooffsell.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ooffsell.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ooffsell.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   		
	<td align="center">주문번호</td>
	<td align="center">상품코드</td>
	<td align="center">옵션명</td>
	<td align="center">상품명</td>
		
	<td align="center">판매액</td>
	<td align="center">결제액</td>
	<% if C_ADMIN_USER then %>
		<td align="center">매입액</td>
	<% end if %>
	<td align="center">판매수량</td>
	<td align="center">합계</td>	
</tr>
<% 
for i=0 to ooffsell.FresultCount-1

totsellprice = totsellprice + ooffsell.FItemList(i).fsellprice
totrealprice = totrealprice + ooffsell.FItemList(i).frealsellprice
totsuplyprice = totsuplyprice + ooffsell.FItemList(i).fsuplyprice
totitemno = totitemno + ooffsell.FItemList(i).fitemno
totsum = totsum + (ooffsell.FItemList(i).frealsellprice * ooffsell.FItemList(i).fitemno)
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#ffffff';>
	<td align="center">
		<%= ooffsell.FItemList(i).forderno %>
	</td>		
	<td align="center">
		<%= ooffsell.FItemList(i).fitemgubun %>-<%= ooffsell.FItemList(i).fitemid %>-<%= ooffsell.FItemList(i).fitemoption %>
	</td>
	<td align="center">
		<%= ooffsell.FItemList(i).fitemoptionname %>
	</td>
			
		<td align="center">
		<%= ooffsell.FItemList(i).fitemname %>
	</td>
	<td align="center">
		<%= FormatNumber(ooffsell.FItemList(i).fsellprice,0) %>
	</td>
			
	<td align="center">
		<%= FormatNumber(ooffsell.FItemList(i).frealsellprice,0) %>
	</td>
	<% if C_ADMIN_USER then %>
		<td align="center">
			<%= FormatNumber(ooffsell.FItemList(i).fsuplyprice,0) %>
		</td>
	<% end if %>
			
	<td align="center">
		<%= ooffsell.FItemList(i).fitemno %>
	</td>
			
	<td align="center">
		<%= FormatNumber(ooffsell.FItemList(i).frealsellprice * ooffsell.FItemList(i).fitemno,0) %>
	</td>	
</tr>   
<% next %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan=4>합계</td>
	<td align="center"><%= FormatNumber(totsellprice,0) %></td>
	<td align="center"><%= FormatNumber(totrealprice,0) %></td>
	<% if C_ADMIN_USER then %>
		<td align="center"><%= FormatNumber(totsuplyprice,0) %></td>
	<% end if %>
	<td align="center"><%= totitemno %></td>
	<td align="center"><%= FormatNumber(totsum,0) %></td>
</tr>	
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ooffsell.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= ooffsell.StartScrollPage-1 %>&<%=parameter%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ooffsell.StartScrollPage to ooffsell.StartScrollPage + ooffsell.FScrollCount - 1 %>
			<% if (i > ooffsell.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ooffsell.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&<%=parameter%>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ooffsell.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&<%=parameter%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->