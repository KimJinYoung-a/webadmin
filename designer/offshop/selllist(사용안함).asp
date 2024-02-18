<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출
' History : 2009.04.07 서동석 생성
'			2010.05.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim page,shopid,jungsanid , yyyy1,mm1,dd1,yyyy2,mm2,dd2 , yyyymmdd1,yyymmdd2
dim fromDate,toDate , i ,datefg , totalitemno , totalsum
	jungsanid = session("ssBctID")
	shopid = request("shopid")
	page = request("page")
	if page="" then page=1
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	datefg = request("datefg")
	if datefg = "" then datefg = "maechul"

totalitemno = 0
totalsum = 0
		
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

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopID = shopid
	ooffsell.FPageSize=20
	ooffsell.FCurrPage=page
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectjungsanId = jungsanid
	ooffsell.frectdatefg = datefg
	
	if shopid<>"" then
		ooffsell.GetDaylySumListByJungsanID
	end if
%>

<script language="javascript">

	function PopItemSellSum(shopid,yyyy1,mm1,dd1,yyyy2,mm2,dd2,oldlist,datefg,Term,jungsanid){		
		var PopItemSellSum = window.open('/common/offshop/dailysellreport_detailitem.asp?shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&oldlist='+oldlist+'&datefg='+datefg+'&Term='+Term+'&jungsanid='+jungsanid,'PopItemSellSum','width=1024,height=768,scrollbars=yes,resizable=yes');
		PopItemSellSum.focus();
	}
	
	function itemsumdetail(menupos,terms,shopid,datefg){
		location.href='todayselldetail.asp?menupos='+menupos+'&terms='+terms+'&shopid='+shopid+'&datefg='+datefg
	}
	
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">        
		SHOP : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		매출기준 :
		<% drawmaechul_datefg "datefg" ,datefg ,""%> <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15">
		* 야간판매를 하는 매장(두타점등등)의 경우 매출기준일로 검색을 하셔야 정확한 매출이 집계 됩니다.
		<br>판매내역은 샾 마감후, 새벽판매매장(새벽5시경), 주간판매매장(오후 10시경) 업데이트 되며,
		<br>정산은 주문일 기준으로 정산 됩니다.
	</td>
</tr>
<tr bgcolor="#EEEEEE" align="center">
	<td>SHOP 구분</td>
	<td>매출발생일</td>
	<td >매출건수(주문건수)</td>
	<td>매출액</td>
	<td>비고</td>	
</tr>
<% if ooffsell.FresultCount > 0 then %>

<% 
for i=0 to ooffsell.FresultCount-1

totalitemno = totalitemno + ooffsell.FItemList(i).FCount
totalSum = totalSum + ooffsell.FItemList(i).FSum
%>
<tr bgcolor="#FFFFFF" height=24 align="center">
	<td><%= ooffsell.FItemList(i).FShopid %></td>
	<td><%= ooffsell.FItemList(i).FTerm %></td>
	<td align="center"><%= ooffsell.FItemList(i).FCount %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSum,0) %></td>
	<td align="center">
		<input type="button" onclick="PopItemSellSum('<%= ooffsell.FItemList(i).Fshopid %>','<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>','','<%= datefg %>','<%= ooffsell.FItemList(i).FTerm %>','<%=jungsanid%>');" value="상품상세" class="button">
		<input type="button" onclick="itemsumdetail('<%= menupos %>','<%= ooffsell.FItemList(i).FTerm %>','<%= ooffsell.FItemList(i).FShopid %>','<%=datefg%>');" value="상품합계상세" class="button">
	</td>
</tr>
<% next %>

<tr bgcolor="#EEEEEE" align="center">
	<td colspan=2>합계</td>	
	<td ><%= FormatNumber(totalitemno,0) %></td>
	<td align="right"><%= FormatNumber(totalSum,0) %></td>
	<td></td>
</tr>

<% else %>

<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center"  >[검색결과가 없습니다.]</td>
</tr>

<% end if %>
</table>

<% if shopid="" then %>
	<script language='javascript'>alert('샾을 선택해 주세요');</script>
<% end if %>

<%
set ooffsell= Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->