<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 카테고리별 통계
' History : 2009.04.07 서동석 생성
'			2010.05.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopmileagecls.asp"-->
<%
dim page,shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,yyyymmdd1,yyymmdd2 ,fromDate,toDate
dim ooffmilde ,i
	shopid = request("shopid")
	page = request("page")
	if page="" then page=1
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-3)
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

set ooffmilde = new COffShopMileage
	ooffmilde.FPageSize=100
	ooffmilde.FCurrpage=page
	ooffmilde.FRectStartDay = fromDate
	ooffmilde.FRectEndDay = toDate
	ooffmilde.FRectShopid=shopid
	ooffmilde.COffShopMileageList
%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>	
			<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
				매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				매장 : <% drawSelectBoxOffShopNot000 "shopid",shopid %>	
			<% end if %>
		<% else %>
			<% if not(C_IS_Maker_Upche) then %> 
				매장 : <% drawSelectBoxOffShopNot000 "shopid",shopid %>	
			<% else %>
				매장 : <% drawSelectBoxOffShopNot000 "shopid",shopid %>	
			<% end if %>
		<% end if %>	
		
		기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
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
<Br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		총건수:<%= ooffmilde.FTotalCount%>, 페이지: <%= page %>/<%= ooffmilde.FTotalPage%>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">회원번호</td>
	<td width="100">회원명</td>
	<td width="100">샾구분</td>
	<td width="80">마일리지</td>
	<td width="100">적요</td>
	<td width="80">저장일</td>
</tr>
<% if ooffmilde.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
<% for i=0 to ooffmilde.FresultCount-1 %>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff'; align="center">
	<td><%= ooffmilde.FItemList(i).Fpointuserno %></td>
	<td ><%= ooffmilde.FItemList(i).Fpointusername %></td>
	<td><%= ooffmilde.FItemList(i).Fshopid %></td>
	<td align="right"><%= ooffmilde.FItemList(i).Fpoint %></td>
	<td align="let"><%= ooffmilde.FItemList(i).Fjukyo %></td>
	<td><%= ooffmilde.FItemList(i).Fregdate %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan=6 align=center>
	<% if ooffmilde.HasPreScroll then %>
		<a href="?page=<%= ooffmilde.StarScrollPage-1 %>&menupos=<%= menupos %>&shopid=<%= shopid %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ooffmilde.StarScrollPage to ooffmilde.FScrollCount + ooffmilde.StarScrollPage - 1 %>
		<% if i>ooffmilde.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>&shopid=<%= shopid %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ooffmilde.HasNextScroll then %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>&shopid=<%= shopid %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% end if %>
</table>

<%
set ooffmilde = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->