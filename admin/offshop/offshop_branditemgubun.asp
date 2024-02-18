<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 브랜드 상품구분별 매출
' History : 2010.05.27 한용민 생성
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
dim yyyy1, mm1, dd1, yyyy2, mm2, dd2 , yyyymmdd1, yyymmdd2,oldlist ,nextdateStr,searchnextdate
dim orderserial,itemid,ojumun ,shopid,page ,ckpointsearch,ckipkumdiv4 ,i,iy,cknodate ,makerid
dim order_desum ,rectdispy, rectselly ,offgubun ,sumprice,tottotalsum , tot10sum ,tot90sum ,tot70sum
dim inc3pl
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	shopid = requestCheckVar(request("shopid"),32)
	orderserial = requestCheckVar(request("orderserial"),16)
	itemid = requestCheckVar(request("itemid"),10)
	ckpointsearch = requestCheckVar(request("ckpointsearch"),10)
	cknodate = requestCheckVar(request("cknodate"),10)
	order_desum = requestCheckVar(request("order_desum"),10)
	rectdispy = requestCheckVar(request("rectdispy"),10)
	rectselly = requestCheckVar(request("rectselly"),10)
	offgubun = requestCheckVar(request("offgubun"),10)
	oldlist = requestCheckVar(request("oldlist"),10)
	makerid = requestCheckVar(request("makerid"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

set ojumun = new COffShopSell

if cknodate="" then
	ojumun.FRectStartDay = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectEndDay = searchnextdate
end if

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

if shopid<>"" then offgubun=""

	ojumun.FRectmakerid = makerid
	ojumun.FRectShopID = shopid
	ojumun.FRectOffgubun = offgubun
	ojumun.FRectOldData = oldlist
	ojumun.FRectInc3pl = inc3pl
	ojumun.fbranditemgubunsum()

tottotalsum = 0
tot10sum = 0
tot90sum = 0
tot70sum = 0
%>

<script language='javascript'>

function ViewOrderDetail(itemid){
	window.open("http://www.10x10.co.kr/street/designershop.asp?itemid=" + itemid,"sample");
}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function ReSearch(ifrm){
	ifrm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3년이전
		&nbsp;&nbsp;
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="ReSearch(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
		&nbsp;&nbsp;
		* 매장구분 : <% Call DrawShopDivCombo("offgubun",offgubun) %>
        &nbsp;&nbsp;
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
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

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= ojumun.ftotalcount %></b>	※총 1000건까지 검색가능
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>매장iD</td>
	<td>브랜드ID</td>
	<td>총매출액</td>
	<td>온라인(10)<br>매출액</td>
	<td>오프전용(90)<br>매출액</td>
	<td>각종부자제(70)<br>매출액</td>
</tr>
<% if ojumun.ftotalcount>0 then %>
<%
for i=0 to ojumun.ftotalcount-1

tottotalsum = tottotalsum + ojumun.FItemList(i).ftotalsum
tot10sum = tot10sum + ojumun.FItemList(i).f10sum
tot90sum = tot90sum + ojumun.FItemList(i).f90sum
tot70sum = tot70sum + ojumun.FItemList(i).f70sum
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#c1c1c1"; onmouseout=this.style.background="#FFFFFF";>
	<td align="center"><%= ojumun.FItemList(i).fshopid %></td>
	<td align="center"><%= ojumun.FItemList(i).fmakerid %></td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ojumun.FItemList(i).ftotalsum,0) %></td>
	<td align="right"><%= FormatNumber(ojumun.FItemList(i).f10sum,0)  %></td>
	<td align="right"><%= FormatNumber(ojumun.FItemList(i).f90sum,0) %></td>
	<td align="right"><%= FormatNumber(ojumun.FItemList(i).f70sum,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan=2 align="center">합계</td>
	<td align="right"><% =FormatNumber(tottotalsum,0) %></td>
	<td align="right"><% =FormatNumber(tot10sum,0) %></td>
	<td align="right"><% =FormatNumber(tot90sum,0) %></td>
	<td align="right"><% =FormatNumber(tot70sum,0) %></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>

<%
	set ojumun = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
