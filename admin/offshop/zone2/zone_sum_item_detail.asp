<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 삽별구역설정
' Hieditor : 2010.12.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->

<%
Dim ozone,i,page , shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,fromDate ,toDate
dim designer , sellgubun ,cdl ,cdm ,cds ,datefg ,menupos , zoneidx ,itemid ,itemname ,searchtype
dim dategubun ,totsumshopsuplycash , ordertype, inc3pl
dim totsellsumshare ,totprofit , totprofitshare , totitemno ,totrealsellprice ,totshopsuplycash
	designer = RequestCheckVar(request("designer"),32)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	sellgubun = requestCheckVar(request("sellgubun"),10)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
	datefg = requestCheckVar(request("datefg"),10)
	zoneidx = requestCheckVar(request("zoneidx"),10)
	menupos = requestCheckVar(request("menupos"),10)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(request("itemname"),124)
	dategubun = requestCheckVar(request("dategubun"),10)
	ordertype = requestCheckVar(request("ordertype"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if ordertype = "" then ordertype = "category"
if dategubun = "" then dategubun = "G"
if datefg = "" then datefg = "maechul"
if sellgubun = "" then sellgubun = "S"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

if page = "" then page = 1
if cdl<>"" and cdm<>"" then cds=""

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
		designer = session("ssBctID")	'"7321"

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

set ozone = new czone_list
	ozone.FPageSize = 1000
	ozone.FCurrPage = page
	ozone.frectdategubun = dategubun
	ozone.frectshopid = shopid
	ozone.FRectStartDay = fromDate
	ozone.FRectEndDay = toDate
	ozone.FRectmakerid = designer
	ozone.FRectCDL = cdl
	ozone.FRectCDM = cdm
	ozone.FRectCDN = cds
	ozone.frectdatefg = datefg
	ozone.frectzoneidx = zoneidx
	ozone.frectitemid = itemid
	ozone.frectitemname = itemname
	ozone.frectsellgubun = sellgubun
	ozone.frectidx = zoneidx
	ozone.frectordertype = ordertype
	ozone.FRectInc3pl = inc3pl

	if shopid <> "" then
		ozone.Getoffshopzone_detail
	end if

	if shopid = "" then response.write "<script>alert('매장을 선택해주세요');</script>"

	totsellsumshare=0
	totprofit =0
	totprofitshare = 0
	totitemno = 0
	totrealsellprice = 0
	totshopsuplycash = 0
	totsumshopsuplycash = 0
%>

<script language="javascript">

	function gopage(page){
		frm.page.value=page;
		frm.submit();
	}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="dategubun" value="<%= dategubun %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 : <% drawmaechul_datefg "datefg" ,datefg ,""%>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>
			<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
				* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* 매장:<% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			<% if not(C_IS_Maker_Upche) then %>
				* 매장:<% drawSelectBoxOffShop "shopid",shopid %>
			<% else %>
				* 매장:<% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gopage('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 브랜드 : <% drawSelectBoxDesignerwithName "designer",designer %>
		&nbsp;&nbsp;
		<% Call zoneselectbox(shopid,"zoneidx",zoneidx,"") %>
		&nbsp;&nbsp;
		* 상품코드 : <input type="text" name="itemid" value="<%=itemid %>" size=10>
		&nbsp;&nbsp;
		* 상품명 : <input type="text" name="itemname" value="<%=itemname %>" size=20>
		<p>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;&nbsp;
		<input type="radio" name="sellgubun" value="S" <% if sellgubun="S" then response.write " checked" %>>결제내역기준
		<input type="radio" name="sellgubun" value="N" <% if sellgubun="N" then response.write " checked" %>>현재등록내역기준
        <p>
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	</td>
</tr>
</table>
<!-- 검색 끝 -->
<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ozone.FTotalCount %></b>
		※ 1000건 까지 검색됩니다.
		<p align="right">
			정렬:
			<select name="ordertype" onchange="gopage('');">
				<option value="" <% if ordertype="" then response.write " selected" %>>정렬선택</option>
				<option value="ea" <% if ordertype="ea" then response.write " selected" %>>수량순</option>
				<option value="totalprice" <% if ordertype="totalprice" then response.write " selected" %>>매출순</option>
				<option value="gain" <% if ordertype="gain" then response.write " selected" %>>수익순</option>
				<option value="unitCost" <% if ordertype="unitCost" then response.write " selected" %>>객단가순</option>
				<option value="category" <% if ordertype="category" then response.write " selected" %>>카테고리순</option>
			</select>
		</p>
	</td>
</tr>
</form>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품번호</td>
	<td>상품명(옵션명)</td>
	<td>브랜드</td>
	<td>대<Br>카테고리</td>
	<td>중<Br>카테고리</td>
	<td>소<Br>카테고리</td>
	<td>판매수</td>
	<td>매출액</td>
	<%' if NOT(C_IS_SHOP) then %>
		<!--<td>매입가</td>
		<td>매입가</td>
		<td>매출액수익</td>
		<td>마진율</td>-->
	<%' end if %>
	<td>조닝명</td>
	<td>조닝<br>크기</td>
	<td>조닝<br>점유율</td>
</tr>
<% if ozone.FresultCount>0 then %>
<%
for i=0 to ozone.FresultCount-1

totitemno = totitemno + ozone.FItemList(i).fitemcnt
totrealsellprice = totrealsellprice + ozone.FItemList(i).fsellsum
totshopsuplycash = totshopsuplycash + ozone.FItemList(i).fshopsuplycash
totsumshopsuplycash = totsumshopsuplycash + ozone.FItemList(i).fshopsuplycash
totprofit = totprofit + (ozone.FItemList(i).fsellsum-ozone.FItemList(i).fshopsuplycash)

if ozone.FItemList(i).fshopsuplycash <> 0 then
	totprofitshare = totprofitshare + round(100-(ozone.FItemList(i).fshopsuplycash/ozone.FItemList(i).fsellsum*100*100)/100,1)
end if
%>
<% if ozone.FItemList(i).fzonename <> "" then %>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% else %>
	<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>

	<td align="center">
		<%= ozone.FItemList(i).fitemgubun %>-<%= CHKIIF(ozone.FItemList(i).fshopitemid>=1000000,Format00(8,ozone.FItemList(i).fshopitemid),Format00(6,ozone.FItemList(i).fshopitemid)) %>-<%= ozone.FItemList(i).fitemoption %>
	</td>
	<td align="center">
		<%= ozone.FItemList(i).fitemname %>
		<% if ozone.FItemList(i).fitemoptionname <> "" then %>
			(<%= ozone.FItemList(i).fitemoptionname %>)
		<% end if %>
	</td>
	<td align="center">
		<%= ozone.FItemList(i).fmakerid %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).fcdl_nm <> "" then %>
			<%= ozone.FItemList(i).fcdl_nm %>
		<% else %>
			-
		<% end if %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).fcdm_nm <> "" then %>
			<%= ozone.FItemList(i).fcdm_nm %>
		<% else %>
			-
		<% end if %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).fcds_nm <> "" then %>
			<%= ozone.FItemList(i).fcds_nm %>
		<% else %>
			-
		<% end if %>
	</td>
	<td align="center">
		<%= ozone.FItemList(i).fitemcnt %>
	</td>
	<td align="center" bgcolor="#E6B9B8">
		<%= FormatNumber(ozone.FItemList(i).fsellsum,0) %>
	</td>
	<%' if NOT(C_IS_SHOP) then %>
		<!--<td>
			<%'= FormatNumber(ozone.FItemList(i).fshopsuplycash,0) %>
		</td>
		<td>
			<%'= FormatNumber(ozone.FItemList(i).fshopsuplycash,0) %>
		</td>
		<td>
			<%'= FormatNumber(ozone.FItemList(i).fsellsum-ozone.FItemList(i).fshopsuplycash,0) %>
		</td>
		<td>
			<%' if ozone.FItemList(i).fshopsuplycash <> 0 then %>
				<%'= FormatNumber(round(100-(ozone.FItemList(i).fshopsuplycash/ozone.FItemList(i).fsellsum*100*100)/100,1),0) %> %
			<%' else %>
				0 %
			<%' end if %>
		</td>-->
	<%' end if %>
	<td align="center">
		<%= ozone.FItemList(i).fzonename %>
	</td>
	<td>
		<%= ozone.FItemList(i).funit %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).funit<>0 then %>
			<%= Clng( ((ozone.FItemList(i).funit / ozone.FItemList(i).frealpyeong) * 10000)) / 100 %> %
		<% end if %>
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=6>
		합계
	</td>
	<td>
		<%= FormatNumber(totitemno,0) %>
	</td>
	<td>
		<%= FormatNumber(totrealsellprice,0) %>
	</td>
	<%' if NOT(C_IS_SHOP) then %>
		<!--<td>
			<%'= FormatNumber(totshopsuplycash,0) %>
		</td>
		<td>
			<%'= FormatNumber(totsumshopsuplycash,0) %>
		</td>
		<td>
			<%'= FormatNumber(totprofit,0) %>
		</td>
		<td>
			<%' if totprofitshare <> 0 then %>
				<%'= FormatNumber(totprofitshare / ozone.fresultcount,0) %> %
			<%' else %>
				0 %
			<%' end if %>
		</td>-->
	<%' end if %>
	<td colspan=3></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set ozone = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->