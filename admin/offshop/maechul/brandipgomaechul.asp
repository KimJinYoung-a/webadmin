<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 브랜드입고대비판매율
' History : 2011.06.21 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offshop_reportcls.asp"-->

<%
dim page , yyyymm ,osum ,shopid , yyyy1 , mm1 ,i ,commcd, inc3pl
dim totipgosum ,tottotsellsum , tottotipgocnt ,tottotsellcnt ,totremainSum ,totstsum ,totstno
	yyyy1    = requestCheckVar(request("yyyy1"),4)
	mm1    = requestCheckVar(request("mm1"),2)
	shopid    = requestCheckVar(request("shopid"),32)
	commcd    = requestCheckVar(request("commcd"),10)
	page    = requestCheckVar(request("page"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if page="" then page="1"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
				
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

if shopid = "" then shopid = "streetshop011"

set osum = new COffshopReport
	osum.FPageSize = 1000
	osum.FCurrPage = page
	osum.FRectShopID = shopid
	osum.frectcommcd = commcd
	osum.FRectyyyymm = yyyy1 & "-" & Format00(2,mm1)
	osum.FRectInc3pl = inc3pl	
	osum.getbrandipgomaechul
%>

<script language='javascript'>

	function ReSearch(page){
		frm.page.value=page;
		frm.submit();
	}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 :
		<% DrawYMBox yyyy1,mm1 %>
		&nbsp;&nbsp;	
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>	
			<% if not C_IS_OWN_SHOP and shopid <> "" then %>
				* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* 매장 : <% drawSelectBoxOffShopNot000 "shopid",shopid %>
			<% end if %>
		<% else %>
			* 매장 : <% drawSelectBoxOffShopNot000 "shopid",shopid %>
		<% end if %>	    
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="ReSearch('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 매입구분 : <% drawSelectBoxOFFJungsanCommCD "commcd",commcd %>
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
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">        	
    </td>
    <td align="right">	       
    </td>        
</tr>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= osum.ftotalcount %></b> ※총 1000건 까지 검색 됩니다
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>매장</td>
	<td>브랜드</td>
	<td>총입고액</td>
	<td>총판매액</td>
	<td>월말재고액</td>
	<td>입고수량</td>
	<td>판매수량</td>
	<td>월말재고수량</td>
	<td>판매율</td>
	<td>판매액-입고액</td>
	<td>매입구분</td>
</tr>
<% if osum.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
<% 
for i=0 to osum.FResultCount - 1

totipgosum = totipgosum + osum.FItemList(i).fipgosum
tottotsellsum = tottotsellsum + osum.FItemList(i).ftotsellsum
tottotipgocnt = tottotipgocnt + osum.FItemList(i).ftotipgocnt
tottotsellcnt = tottotsellcnt + osum.FItemList(i).ftotsellcnt
totremainSum = totremainSum + osum.FItemList(i).fremainSum
totstsum = totstsum + osum.FItemList(i).fstsum
totstno = totstno + osum.FItemList(i).fstno
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF"; align="center">
	<td><%= osum.FItemList(i).fshopid %></td>
	<td><%= osum.FItemList(i).fmakerid %></td>
	<td><%= FormatNumber(osum.FItemList(i).fipgosum,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).ftotsellsum,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).fstsum,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).ftotipgocnt,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).ftotsellcnt,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).fstno,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).fpro,0) %>%</td>
	<td><%= FormatNumber(osum.FItemList(i).fremainSum,0) %></td>
	<td><%= osum.FItemList(i).fcomm_name %> (<%= osum.FItemList(i).fcomm_cd %>)</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=2></td>
	<td><%= FormatNumber(totipgosum,0) %></td>
	<td><%= FormatNumber(tottotsellsum,0) %></td>
	<td><%= FormatNumber(totstsum,0) %></td>
	<td><%= FormatNumber(tottotipgocnt,0) %></td>
	<td><%= FormatNumber(tottotsellcnt,0) %></td>
	<td><%= FormatNumber(totstno,0) %></td>
	<td></td>
	<td><%= FormatNumber(totremainSum,0) %></td>
	<td></td>	
</tr>
<% end if %>
</table>

<%
	set osum = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
