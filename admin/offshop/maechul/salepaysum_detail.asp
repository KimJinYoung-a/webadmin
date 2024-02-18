<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 할인 유형별 통계
' History : 2012.02.10 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2 , fromDate,toDate , shopid ,i ,makerid ,datefg , tmpdate ,discountKind
dim totsellcntsum ,totsellprice ,totrealsellprice ,totsaleprice ,totsuplyprice ,totshopbuyprice , menupos
	makerid = requestCheckVar(request("makerid"),32)
	menupos = requestCheckVar(request("menupos"),10)
	discountKind = requestCheckVar(request("discountKind"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),10)

if datefg = "" then datefg = "maechul"

tmpdate = dateadd("d",-1,date)

if (yyyy1="") then yyyy1 = Cstr(Year(tmpdate))
if (mm1="") then mm1 = Cstr(Month(tmpdate))
if (dd1="") then dd1 = Cstr(day(tmpdate))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

'C_IS_SHOP = TRUE
'C_IS_Maker_Upche = TRUE

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
		makerid = session("ssBctID")	'"7321"

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

dim oreport
set oreport = new COffShopSell
	oreport.FPageSize = 2000
	oreport.FRectFromDate = fromDate
	oreport.FRectToDate = toDate
	oreport.frectdatefg = datefg
	oreport.FRectShopID = shopid
	oreport.frectdiscountKind = discountKind
	oreport.frectmakerid = makerid
	oreport.Getsalepaysum_detail

totsellcntsum = 0
totsellprice = 0
totrealsellprice = 0
totsaleprice = 0
totsuplyprice = 0
totshopbuyprice	 = 0
%>

<script language='javascript'>

function reg(){
	frm.submit();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="showtype" value="showtype">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="1" cellspacing="1" class="a">
		<tr>
			<td>
				브랜드:<% drawSelectBoxDesignerwithName "makerid",makerid %>
				할인유형 : <% DrawdiscountKind "discountKind" ,discountKind , " onchange='reg();'" %>
				<Br>
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					매장 : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
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

<table width="100%" cellspacing="1" cellpadding="3" class="a" bgcolor=#3d3d3d>
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        검색결과 : <b><%= oreport.FTotalcount %></b> ※ 최대 2000건까지 검색가능
    </td>
</tr>
<tr bgcolor="#EEEEEE" align="center">
	<td>매장</td>
	<td>상품번호</td>
	<td>
		상품명<font color="blue">(옵션명)</font>
	</td>
	<td>브랜드</td>
	<td>상품<br>수량</td>
	<td>매출액</td>
	<td>실매출액</td>
	<td>할인액</td>

	<% if not(C_IS_SHOP) then %>
		<td>
			매입액
		</td>
	<% end if %>

	<% if not(C_IS_SHOP) and not(C_IS_Maker_Upche) then %>
		<td>
			매장<br>매입액
		</td>
	<% end if %>
</tr>
<%
if oreport.FResultCount > 0 then

for i=0 to oreport.FResultCount - 1

totsellcntsum = totsellcntsum + oreport.FItemList(i).fsellcntsum
totsellprice = totsellprice + oreport.FItemList(i).fsellprice
totrealsellprice = totrealsellprice + oreport.FItemList(i).frealsellprice
totsaleprice = totsaleprice + oreport.FItemList(i).fsaleprice
totsuplyprice = totsuplyprice + oreport.FItemList(i).fsuplyprice
totshopbuyprice = totshopbuyprice + oreport.FItemList(i).fshopbuyprice
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='#FFFFFF'; align="center">
	<td>
		<%= oreport.FItemList(i).fshopname %>
	</td>
	<td><%=oreport.FItemList(i).fitemgubun%><%=CHKIIF(oreport.FItemList(i).fitemid>=1000000,Format00(8,oreport.FItemList(i).fitemid),Format00(6,oreport.FItemList(i).fitemid))%><%=oreport.FItemList(i).fitemoption%></td>
	<td align="left">
		<%= oreport.FItemList(i).fitemname %>
		<% if oreport.FItemList(i).fitemoptionname <> "" then %>
			<font color="blue">(<%= oreport.FItemList(i).fitemoptionname %>)</font>
		<% end if %>
	</td>
	<td>
		<%= oreport.FItemList(i).fmakerid %>
	</td>
	<td align="right"><%= FormatNumber(oreport.FItemList(i).fsellcntsum,0) %></td>
	<td align="right"><%= FormatNumber(oreport.FItemList(i).fsellprice,0) %></td>
	<td align="right"><%= FormatNumber(oreport.FItemList(i).frealsellprice,0) %></td>
	<td align="right"><%= FormatNumber(oreport.FItemList(i).fsaleprice,0) %></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right"><%= FormatNumber(oreport.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>

	<% if not(C_IS_SHOP) and not(C_IS_Maker_Upche) then %>
		<td align="right"><%= FormatNumber(oreport.FItemList(i).fshopbuyprice,0) %></td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height=24 align="center">
	<td colspan=4>
		총계
	</td>
	<td align="right"><%= FormatNumber(totsellcntsum,0) %></td>
	<td align="right"><%= FormatNumber(totsellprice,0) %></td>
	<td align="right"><%= FormatNumber(totrealsellprice,0) %></td>

	<td align="right"><%= FormatNumber(totsaleprice,0) %></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right"><%= FormatNumber(totsuplyprice,0) %></td>
	<% end if %>

	<% if not(C_IS_SHOP) and not(C_IS_Maker_Upche) then %>
		<td align="right"><%= FormatNumber(totshopbuyprice,0) %></td>
	<% end if %>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" height=24>
	<td align="center" colspan=25>검색 결과가 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->