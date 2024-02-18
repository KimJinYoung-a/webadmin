<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출
' History : 2010.03.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim page,shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,yyyymmdd1,yyymmdd2,fromDate,toDate
dim datefg , i
	shopid = request("shopid")
	datefg = request("datefg")
	if datefg = "" then datefg = "maechul"
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	page = request("page")
	if page="" then page=1


''계정별 fix
'if ((session("ssBctDiv")>9) and (session("ssBctBigo")<>""))  then shopid=session("ssBctBigo")

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

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopID = shopid
	ooffsell.FPageSize=20
	ooffsell.FCurrPage=page
	ooffsell.FRectNormalOnly = "on"
	ooffsell.frectdatefg = datefg	
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	
	if shopid<>"" then
		ooffsell.GetDaylySumList
	else
		response.write "<script language='javascript'>"
		response.write "alert('지점을 선택하신 후 검색하세요.');"
		response.write "</script>"		
	end if
%>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				샾 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
				매출기준 :
				<% drawmaechuldatefg "datefg" ,datefg ,""%> 
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
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
    <tr valign="bottom">       
        <td align="left">        	
	    </td>
	    <td align="right">	       
        </td>        
	</tr>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#EEEEEE">
	<td width="80">샾구분</td>
	<td width="80">매출일</td>
	<td width="80">매출건수</td>
	<td width="80">매출액</td>
	<td width="80">결제액</td>
	<td width="80">마일리지사용</td>
	<td width="80">마일리지적립</td>
	<td width="60">아이템목록</td>
	<td width="60">주문별목록</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= ooffsell.FItemList(i).FShopid %></td>
	<td><%= ooffsell.FItemList(i).FTerm %></td>
	<td align="center"><%= ooffsell.FItemList(i).FCount %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSum+ooffsell.FItemList(i).FSpendMile,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSum,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSpendMile,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FGainMile,0) %></td>
	<td align="center">
		<a href="todayselldetail.asp?menupos=<%= menupos %>&terms=<%= ooffsell.FItemList(i).FTerm %>&shopid=<%= ooffsell.FItemList(i).FShopid %>&datefg=<%=datefg%>">
		보기</a>
	</td>
	<td align="center">
		<a href="todaysellmaster.asp?menupos=<%= menupos %>&terms=<%= ooffsell.FItemList(i).FTerm %>&shopid=<%= ooffsell.FItemList(i).FShopid %>&datefg=<%=datefg%>">
		보기</a>
	</td>
</tr>
<% next %>
</table>

<%
set ooffsell= Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->