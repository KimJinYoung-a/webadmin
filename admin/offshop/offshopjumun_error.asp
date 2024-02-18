<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출
' History : 2010.05.11 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim totalpage, totalnum, q, i , yyyy1,mm1,dd1,yyyy2,mm2,dd2 , yyyymmdd1,yyymmdd2
dim ojumun ,fromDate,toDate,jnx,tmpStr,siteId ,searchId ,ck_nopoint
dim TTLselltotal,TTLbuytotal,TTLsellcnt ,TTLminustotal,TTLminusbuytotal,TTLminuscount
dim shopid , totprofit , totmagin , page, inc3pl
	ck_nopoint  = requestCheckVar(request("ck_nopoint"),10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	shopid = requestCheckVar(request("shopid"),32)
	page  = requestCheckVar(request("page"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if page = "" then page = 1
if (yyyy1="") then
	dim thedate : thedate = dateadd("d",-4,now())
	yyyy1 = Cstr(Year(thedate))
	mm1 = Cstr(Month(thedate))
	dd1 = "01"

	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

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

set ojumun = new COffShopSell
	ojumun.FPageSize = 100
	ojumun.FCurrPage = page
	ojumun.FRectStartDay = fromDate
	ojumun.FRectEndDay = toDate
	ojumun.FRectShopID = shopid
	ojumun.FRectInc3pl = inc3pl
	ojumun.foffshopjumun_error()
%>

<script language='javascript'>

function Check(){
   document.frm.submit();
}

function jFlagDel(orderno,dtlidx){
	//정산구분만 B000으로 변경.
	if(confirm("정산구분을 미지정으로 변경 하시겠습니까? ")){

		var jFlagDel = window.open('doshopjumun.asp?orderno='+orderno+'&idx='+dtlidx+'&mode=jflagdel','jflagedit','width=300,height=200,scrollbars=yes,resizable=yes');
		jFlagDel.focus();
	}
}

function jFlagB031(orderno,dtlidx,jcommcd){
	if(confirm("정산구분을 출고매입(B031)으로 변경 하시겠습니까? ")){

		var jFlagDel = window.open('doshopjumun.asp?orderno='+orderno+'&idx='+dtlidx+'&jcommcd='+jcommcd+'&mode=jflagedit','jflagedit','width=300,height=200,scrollbars=yes,resizable=yes');
		jFlagDel.focus();
	}
}

function popOffContract(shopid,makerid){
    var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + makerid,"popshopupcheinfo","width=700 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
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
		<input type="button" class="button_s" value="검색" onClick="Check();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
        &nbsp;&nbsp;
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="left">
		총건수:&nbsp;<%= FormatNumber(ojumun.ftotalcount,0) %> ※ 100건 까지만 검색 됩니다
	</td>
</tr>
<tr bgcolor="#EEEEEE" align="center">
	<td>매장</td>
	<td>주문번호</td>
	<td>캐셔ID</td>
	<td>브랜드</td>
	<td>상품코드</td>
	<td>옵션명</td>
	<td>상품명</td>
	<td>판매수</td>
	<td>매입가</td>
	<td>판매금액</td>
	<td>비고</td>
</tr>
<% if ojumun.ftotalcount < 1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan=15 align="center"  >[검색결과가 없습니다.]</td>
</tr>
<% else %>
<%
for i=0 to ojumun.ftotalcount - 1
%>
<tr bgcolor="#FFFFFF" align="center">
	<td height="10">
      	<%= ojumun.FItemList(i).fshopid %>
  	</td>
	<td height="10">
      	<%= ojumun.FItemList(i).forderno %>
  	</td>
	<td height="10">
      	<%= ojumun.FItemList(i).fcasherid %>
  	</td>
	<td height="10">
	    <% if (C_IS_SHOP) then %>
      	<%= ojumun.FItemList(i).fmakerid %>
      	<% else %>
      	<a href="javascript:popOffContract('<%= ojumun.FItemList(i).fshopid %>','<%= ojumun.FItemList(i).fmakerid %>');"><%= ojumun.FItemList(i).fmakerid %></a>
      	<% end if %>
  	</td>
	<td height="10">
      	<%=ojumun.FItemList(i).fitemgubun%>-<%=CHKIIF(ojumun.FItemList(i).fitemid>=1000000,Format00(8,ojumun.FItemList(i).fitemid),Format00(6,ojumun.FItemList(i).fitemid))%>-<%=ojumun.FItemList(i).fitemoption%>
  	</td>
	<td height="10">
      	<%= ojumun.FItemList(i).fitemoptionname %>
  	</td>
	<td height="10">
      	<%= ojumun.FItemList(i).fitemname %>
  	</td>
	<td height="10">
      	<%= ojumun.FItemList(i).fitemno %>
  	</td>
	<td height="10">
      	<%= FormatNumber(ojumun.FItemList(i).fsuplyprice,0) %>
  	</td>
	<td height="10">
      	<%= FormatNumber(ojumun.FItemList(i).fsellprice,0) %>
  	</td>
	<td height="10">
		<%
		'//함부로 지우믄 안댐.
		if (C_ADMIN_AUTH) then
		%>
		<input type="button" value="미지정변경" onclick="jFlagDel('<%= ojumun.FItemList(i).forderno %>','<%= ojumun.FItemList(i).fdetailidx %>');" class="button">
		<br>
		<% if (ojumun.FItemList(i).fshopid="streetshop011" or ojumun.FItemList(i).fshopid="streetshop999") then %>
		<input type="button" value="출고매입변경" onclick="jFlagB031('<%= ojumun.FItemList(i).forderno %>','<%= ojumun.FItemList(i).fdetailidx %>','B031');" class="button">
		<% end if %>
		<% end if %>
  	</td>
</tr>
<% next %>
<% end if %>
</table>

<%
set ojumun = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
