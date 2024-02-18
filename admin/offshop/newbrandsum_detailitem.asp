<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출 (데이타마트 통계서버에서 가져옴)
' History : 2010.05.10 서동석 생성
'			2012.02.07 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2 , fromDate,toDate , shopid ,i , datelen, datelen2 ,makerid, menupos, page
dim datefg , tmpdate , maechultype ,totrealsellprice ,totitemno ,totprofit ,totsellprice, totsuplyprice
	makerid = requestCheckVar(request("makerid"),32)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),32)
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)

if page="" then page = 1
if datefg = "" then datefg = "maechul"
tmpdate = dateadd("m",-1,date)

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
	oreport.FCurrPage = page
	oreport.frectdatefg = datefg
	oreport.FRectFromDate = fromDate
	oreport.FRectToDate = toDate
	oreport.FRectShopID = shopid
	oreport.frectmakerid = makerid

	'/데이타마트
	oreport.GetNewBrandSell_item_datamart

	'/메인디비 실시간
	'oreport.GetNewBrandSell_item

totrealsellprice = 0
totitemno =0
totprofit = 0
totsellprice = 0
totsuplyprice = 0
%>

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
				* 기간 :
				<% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBoxdynamic yyyy1,"yyyy1",yyyy2,"yyyy2",mm1,"mm1",mm2,"mm2",dd1,"dd1",dd2,"dd2" %>
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
				* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
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

<Br>
<!-- 표 중간바 시작-->
<table width="100%" cellpadding="3" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		※ 하루 전날까지 판매된 매출 통계이며, 하루에 한번 새벽에 업데이트 됩니다.
    </td>
    <td align="right">

    </td>
</tr>
</form>
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oreport.FResultCount %></b> ※ 총 2000건까지 까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>매장명</td>
	<td>날짜</td>
	<td>물류코드</td>
	<td>상품명<Br><font color="blue">[옵션명]</font></td>
	<td>브랜드</td>
	<td>판매가</td>
	<td>매출액</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>매입가</td>
	<% end if %>

	<td>판매<Br>수량</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>매출<Br>수익</td>
	<% end if %>
</tr>
<%
if oreport.FResultCount > 0 then

for i=0 to oreport.FResultCount - 1

totsellprice = totsellprice + oreport.FItemList(i).fsellprice
totrealsellprice = totrealsellprice + oreport.FItemList(i).frealsellprice
totsuplyprice = totsuplyprice + oreport.FItemList(i).fsuplyprice
totitemno = totitemno + oreport.FItemList(i).fitemno
totprofit = totprofit + (oreport.FItemList(i).frealsellprice - oreport.FItemList(i).fsuplyprice)
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF"; align="center">
	<td><%= oreport.FItemList(i).fshopname %><Br>(<%= oreport.FItemList(i).fshopid %>)</td>
	<td><%= oreport.FItemList(i).fIXyyyymmdd %></td>
	<td><%= oreport.FItemList(i).fitemgubun %><%= CHKIIF(oreport.FItemList(i).fitemid>=1000000,Format00(8,oreport.FItemList(i).fitemid),Format00(6,oreport.FItemList(i).fitemid)) %><%= oreport.FItemList(i).fitemoption %></td>
	<td>
		<%= oreport.FItemList(i).fitemname %>

		<% if oreport.FItemList(i).fitemoptionname <> "" then %>
			<BR><font color="blue">[<%= oreport.FItemList(i).fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td><%= oreport.FItemList(i).fmakerid %></td>
	<td align="right"><%= FormatNumber(oreport.FItemList(i).fsellprice,0) %></td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(oreport.FItemList(i).frealsellprice,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(oreport.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>

	<td><%= oreport.FItemList(i).fitemno %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
			<%= FormatNumber(oreport.FItemList(i).frealsellprice - oreport.FItemList(i).fsuplyprice,0) %>
		</td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td colspan=5>합계</td>
	<td align="right"><%= FormatNumber(totsellprice,0) %></td>
	<td align="right"><%= FormatNumber(totrealsellprice,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(totsuplyprice,0) %></td>
	<% end if %>

	<td><%= FormatNumber(totitemno,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(totprofit,0) %></td>
	<% end if %>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" height=24>
	<td align="center" colspan=15>검색 결과가 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->