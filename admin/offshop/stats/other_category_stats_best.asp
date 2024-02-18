<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 수작업 카테고리 통계 
'				이페이지는 현 카테고리와 무관하게 수작업으로 작성되는 통계입니다
' Hieditor : 2011.11.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/stats/other_category_stats_cls.asp"-->

<%
Dim othercate,i,page , parameter , shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,fromDate ,toDate ,othercheck
dim designer ,datefg , othercdl ,totsellcnt , totsellsum , totsuplysum , menupos ,catecdm, inc3pl
	designer = RequestCheckVar(request("designer"),32)
	othercdl = RequestCheckVar(request("othercdl"),10)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),10)
	menupos = requestCheckVar(request("menupos"),10)
	catecdm = requestCheckVar(request("catecdm"),3)
	othercheck = requestCheckVar(request("othercheck"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if datefg = "" then datefg = "maechul"			
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)
		
if page = "" then page = 1
if othercdl = "" then othercdl = "070"
		
'직영/가맹점
if (C_IS_SHOP) then
	
	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if	
end if

set othercate = new cothercate_list
	othercate.FPageSize = 100
	othercate.FCurrPage = page
	othercate.frectshopid = shopid
	othercate.FRectStartDay = fromDate
	othercate.FRectEndDay = toDate
	othercate.FRectmakerid = designer	
	othercate.frectdatefg = datefg
	othercate.frectothercdl = othercdl
	othercate.frectcatecdm = catecdm
	othercate.frectothercheck = othercheck
	othercate.FRectInc3pl = inc3pl
	
	if shopid <> "" then
		othercate.getother_category_best
	end if
	
	if shopid = "" then response.write "<script>alert('매장을 선택해주세요');</script>"

totsellcnt = 0
totsellsum = 0
totsuplysum = 0
%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="othercdl" value="<%= othercdl %>">
<input type="hidden" name="datefg" value="<%= datefg %>">
<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
<input type="hidden" name="mm1" value="<%= mm1 %>">
<input type="hidden" name="dd1" value="<%= dd1 %>">
<input type="hidden" name="yyyy2" value="<%= yyyy2 %>">
<input type="hidden" name="mm2" value="<%= mm2 %>">
<input type="hidden" name="dd2" value="<%= dd2 %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 브랜드 : <% drawSelectBoxDesignerwithName "designer",designer %>
        &nbsp;&nbsp;
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>		
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
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
		검색결과 : <b><%= othercate.FTotalCount %></b>
		※ 100건 까지 검색됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>순위</td>
	<td>상품명(옵션명)</td>
	<td>상품명</td>
	<td>브랜드</td>
	<td>판매가</td>
	<td>판매수</td>
	<td>매출액</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>매입액</td>
	<% end if %>
	
	<td>비고</td>
</tr>
<% if othercate.FTotalCount>0 then %>
<% 
for i=0 to othercate.FTotalCount-1 

totsellcnt = totsellcnt + othercate.FItemList(i).fsellcnt
totsellsum = totsellsum + othercate.FItemList(i).fsellsum
totsuplysum = totsuplysum + othercate.FItemList(i).fsuplysum
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>	
	<td>
		<%= i+1 %>
	</td>
	<td>
		<%= othercate.FItemList(i).fitemid %>	
	</td>
	<td>
		<%= othercate.FItemList(i).fitemname %>
		<% if othercate.FItemList(i).fitemoptionname <> "" and not isnull(othercate.FItemList(i).fitemoptionname) then %>
			(<%= othercate.FItemList(i).fitemoptionname %>)
		<% end if %>			
	</td>
	<td>
		<%= othercate.FItemList(i).fmakerid %>
	</td>
	<td>
		<%= FormatNumber(othercate.FItemList(i).fsellprice,0) %>
	</td>
	<td>
		<%= FormatNumber(othercate.FItemList(i).fsellcnt,0) %>
	</td>
	<td bgcolor="#E6B9B8">
		<%= FormatNumber(othercate.FItemList(i).fsellsum,0) %>
	</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>
			<%= FormatNumber(othercate.FItemList(i).fsuplysum,0) %>
		</td>
	<% end if %>
	
	<td width=100>
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">	
	<td colspan=5>합계</td>
	<td>
		<%= FormatNumber(totsellcnt,0) %>
	</td>
	<td>
		<%= FormatNumber(totsellsum,0) %>
	</td>
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>
			<%= FormatNumber(totsuplysum,0) %>
		</td>	
	<% end if %>	
	<td></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set othercate = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->