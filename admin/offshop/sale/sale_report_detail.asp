<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 통계
' History : 2012.10.25 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/sale/salereport_cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->

<%
dim SType , sale_code,i ,page ,shopid, yyyy1, mm1, dd1, yyyy2, mm2, dd2
dim fromDate , toDate, menupos, inc3pl
	SType = requestCheckVar(request("SType"),1)
	sale_code = requestCheckVar(request("sale_code"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if SType = "" then SType = "D"
if page = "" then page = 1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-90)
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

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

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

dim oReport
set oReport = new Csalereport_list
	oReport.FRectsale_code = sale_code
	oReport.FRectshopid = shopid
	oReport.frectevt_startdate = fromDate
	oReport.frectevt_enddate = toDate
	oReport.FPageSize = 1000
	oReport.FCurrPage = page
	oReport.FRectInc3pl = inc3pl
%>

<script language="javascript">

	function regsubmit(){
		frm.submit();
	}

	//상품매출
	function item_detail(SType,shopid,sale_code,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
		location.href='?SType='+SType+'&sale_code='+sale_code+'&shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&menupos=<%=menupos%>';
	}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 :
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					<% if not(C_IS_Maker_Upche) then %>
						* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% else %>
						* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% end if %>

				<p>
				* 할인번호 : <input type="text" name="sale_code" size="10" value="<%= sale_code %>">
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="regsubmit();">
	</td>
</tr>
</table>
<!-- 표 상단바 끝-->
<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	※1000건 까지 검색가능
    </td>
    <td align="right">
		분류:
		<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %> onclick="regsubmit();">날짜별
		<input type="radio" name="SType" value="I" <% If SType = "I" Then response.write "checked" %> onclick="regsubmit();">상품별
		<input type="radio" name="SType" value="B" <% If SType = "B" Then response.write "checked" %> onclick="regsubmit();">브랜드별
    </td>
</tr>
</form>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">

<%
'// 날짜별 할인 통계
if SType = "D" then

	'//통계테이블에서 가져옴
	oReport.getsaledate_sum()
%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>구매일</td>
		<td>매장</td>
		<td>할인<Br>코드</td>
		<td>할인명</td>
		<!--<td>매출액</td>-->
		<td>매출액</td>
		<!--<td>주문건수</td>-->
		<td>판매<br>수량</td>
		<td>비고</td>
	</tr>
	<%
	dim totsellprice ,totrealsellprice ,totselljumuncnt ,totsellCnt
		totsellprice = 0
		totrealsellprice = 0
		totselljumuncnt = 0
		totsellCnt = 0

	if oReport.FResultCount > 0 then

	for i=0 to oReport.FResultCount-1

	totsellprice = totsellprice + oReport.FItemList(i).ftotsellprice
	totrealsellprice = totrealsellprice + oReport.FItemList(i).ftotrealsellprice
	totselljumuncnt = totselljumuncnt + oReport.FItemList(i).ftotselljumuncnt
	totsellCnt = totsellCnt + oReport.FItemList(i).ftotsellCnt
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td width=80><%= oReport.FItemList(i).fyyyymmdd %></td>
		<td><%= oReport.FItemList(i).fshopname %></td>
		<td width=60><%= oReport.FItemList(i).fsale_code %></td>
		<td><%= oReport.FItemList(i).fsale_name %></td>
		<!--<td width=80 align="right"><%'= FormatNumber(oReport.FItemList(i).ftotsellprice,0) %></td>-->
		<td width=80 align="right" bgcolor="#E6B9B8"><%= FormatNumber(oReport.FItemList(i).ftotrealsellprice,0) %></td>
		<!--<td width=50 align="right"><%'= FormatNumber(oReport.FItemList(i).ftotselljumuncnt,0) %></td>-->
		<td width=50 align="right"><%= FormatNumber(oReport.FItemList(i).ftotsellCnt,0) %></td>
		<td width=80>
			<input type="button" class="button" value="상품상세" onclick="item_detail('I','<%= oReport.FItemList(i).fshopid %>','<%= oReport.FItemList(i).fsale_code %>','<%= left(oReport.FItemList(i).fyyyymmdd,4) %>','<%= mid(oReport.FItemList(i).fyyyymmdd,6,2) %>','<%= right(oReport.FItemList(i).fyyyymmdd,2) %>','<%= left(oReport.FItemList(i).fyyyymmdd,4) %>','<%= mid(oReport.FItemList(i).fyyyymmdd,6,2) %>','<%= right(oReport.FItemList(i).fyyyymmdd,2) %>');">
		</td>
	</tr>
	<%
	next
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td colspan=4>총합</td>
		<!--<td align="right"><%'= FormatNumber(totsellprice,0) %></td>-->
		<td align="right"><%= FormatNumber(totrealsellprice,0) %></td>
		<!--<td align="right"><%'= FormatNumber(totselljumuncnt,0) %></td>-->
		<td align="right"><%= FormatNumber(totsellCnt,0) %></td>
		<td></td>
	</tr>

	<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">등록된 내용이 없습니다.</td>
	</tr>
	<%
	end if
	%>
<%
'// 상품별 할인 통계
elseif SType = "I" then

	oReport.getsaleitem_sum()
%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>매장</td>
		<td>할인<Br>코드</td>
		<td>할인명</td>
		<td>상품코드</td>
		<td>브랜드</td>
		<td>상품명<font color='blue'>(옵션명)<font></td>
		<!--<td>매출액</td>-->
		<td>매출액</td>
		<td>판매<br>수량</td>
	</tr>
	<%
	dim totsuplyprice, totbuyprice, totitemno
		totsellprice = 0
		totrealsellprice = 0
		totsuplyprice = 0
		totbuyprice = 0
		totitemno = 0

	if oReport.FResultCount > 0 then

	for i=0 to oReport.FResultCount-1

	totsellprice = totsellprice + oReport.FItemList(i).ftotsellprice
	totrealsellprice = totrealsellprice + oReport.FItemList(i).ftotrealsellprice
	totsuplyprice = totsuplyprice + oReport.FItemList(i).ftotsuplyprice
	totbuyprice = totbuyprice + oReport.FItemList(i).ftotbuyprice
	totitemno = totitemno + oReport.FItemList(i).ftotitemno
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td>
			<%= oReport.FItemList(i).fshopname %>
		</td>
		<td width=60>
			<%= oReport.FItemList(i).fsale_code %>
		</td>
		<td>
			<%= oReport.FItemList(i).fsale_name %>
		</td>
		<td width=90>
			<%= oReport.FItemList(i).fitemgubun %><%= CHKIIF(oReport.FItemList(i).FItemid>=1000000,Format00(8,oReport.FItemList(i).FItemid),Format00(6,oReport.FItemList(i).FItemid)) %><%= oReport.FItemList(i).fitemoption %>
		</td>
		<td>
			<%= oReport.FItemList(i).fmakerid %>
		</td>
		<td>
			<%= oReport.FItemList(i).fitemname %>
			<%
			if oReport.FItemList(i).fitemoption <> "0000" then
				response.write "<font color='blue'>("&oReport.FItemList(i).fitemoptionname&")<font>"
			end if
			%>
		</td>
		<!--<td width=80 align="right">
			<%'= FormatNumber(oReport.FItemList(i).ftotsellprice,0) %>
		</td>-->
		<td width=80 align="right" bgcolor="#E6B9B8">
			<%= FormatNumber(oReport.FItemList(i).ftotrealsellprice,0) %>
		</td>
		<td width=50 align="right">
			<%= oReport.FItemList(i).ftotitemno %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan=6 align="center">총합</td>
		<!--<td align="right">
			<%'= FormatNumber(totsellprice,0) %>
		</td>-->
		<td align="right">
			<%= FormatNumber(totrealsellprice,0) %>
		</td>
		<td align="right">
			<%= FormatNumber(totitemno,0) %>
		</td>
	</tr>
	<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">등록된 내용이 없습니다.</td>
	</tr>
	<% end if %>
<%
'// 브랜드별 할인 통계
elseif SType = "B" then

	oReport.getsalebrand_sum()
%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>매장</td>
		<td>할인<Br>코드</td>
		<td>할인명</td>
		<td>브랜드</td>
		<!--<td>매출액</td>-->
		<td>매출액</td>
		<td>판매<br>수량</td>
	</tr>
	<%
		totsellprice = 0
		totrealsellprice = 0
		totsuplyprice = 0
		totbuyprice = 0
		totitemno = 0

	if oReport.FResultCount > 0 then

	for i=0 to oReport.FResultCount-1

	totsellprice = totsellprice + oReport.FItemList(i).ftotsellprice
	totrealsellprice = totrealsellprice + oReport.FItemList(i).ftotrealsellprice
	totsuplyprice = totsuplyprice + oReport.FItemList(i).ftotsuplyprice
	totbuyprice = totbuyprice + oReport.FItemList(i).ftotbuyprice
	totitemno = totitemno + oReport.FItemList(i).ftotitemno
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td>
			<%= oReport.FItemList(i).fshopname %>
		</td>
		<td width=60>
			<%= oReport.FItemList(i).fsale_code %>
		</td>
		<td>
			<%= oReport.FItemList(i).fsale_name %>
		</td>
		<td>
			<%= oReport.FItemList(i).fmakerid %>
		</td>
		<!--<td width=80 align="right">
			<%'= FormatNumber(oReport.FItemList(i).ftotsellprice,0) %>
		</td>-->
		<td width=80 align="right" bgcolor="#E6B9B8">
			<%= FormatNumber(oReport.FItemList(i).ftotrealsellprice,0) %>
		</td>
		<td width=50 align="right">
			<%= oReport.FItemList(i).ftotitemno %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan=4 align="center">총합</td>
		<!--<td align="right">
			<%'= FormatNumber(totsellprice,0) %>
		</td>-->
		<td align="right">
			<%= FormatNumber(totrealsellprice,0) %>
		</td>
		<td align="right">
			<%= FormatNumber(totitemno,0) %>
		</td>
	</tr>
	<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">등록된 내용이 없습니다.</td>
	</tr>
	<% end if %>	
<%
end if
%>
</table>

<%
set oReport = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->