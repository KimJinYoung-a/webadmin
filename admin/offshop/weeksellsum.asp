<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 주별매출통계
' History : 2010.06.10 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->

<%
dim research , shopid ,opt_rect ,i,p1,p2,p3,p4 ,maybe_monthcount ,maybe_monthsum ,dayno, currno
dim nowdate, nowyyyymm, page, inc3pl
	opt_rect = requestCheckVar(request("opt_rect"),16)
	research = requestCheckVar(request("research"),2)
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if page="" then page=1
if research<>"on" then	
	if opt_rect="" then opt_rect="24"
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

dim oreport
set oreport = new COffShopSell
	oreport.FRectSearchType = opt_rect
	oreport.FRectShopID = shopid
	oreport.FCurrPage = page
	oreport.FRectInc3pl = inc3pl	
	oreport.Fpagesize=500
	oreport.fweeksellsum

if opt_rect="all" then
	nowdate = CStr(date)
	nowyyyymm = left(nowdate,7)
	currno = CInt(right(nowdate,2))
	nowdate = dateserial(Left(nowdate,4),Mid(nowdate,6,2)+1,0)
	dayno = CInt(right(nowdate,2))
end if
%>

<!-- 표 상단바 시작-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
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
	            &nbsp;&nbsp;
				<input type="radio" name="opt_rect" value="24" <% if opt_rect="24" then response.write "checked" %> >24주
				<input type="radio" name="opt_rect" value="48" <% if opt_rect="48" then response.write "checked" %> >48주
				<input type="radio" name="opt_rect" value="all" <% if opt_rect="all" then response.write "checked" %> >전체
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>					
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

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oreport.FResultCount %></b> ※ 최대 500건 까지 검색됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
  <td width=100>기간</td>
  <td></td>
  <td width=100>매출액<br>(마일리지포함)</td>
  <td width=50>건수</td>
</tr>
<%
if oreport.FResultCount > 0 then 
for i=0 to oreport.FResultCount - 1

if Left(oreport.FItemList(i).Fsitename,7)=nowyyyymm then
	
	if oreport.FItemList(i).Fselltotal <> 0 and oreport.FItemList(i).Fselltotal <> "" then
		maybe_monthsum = CLng( (oreport.FItemList(i).Fselltotal+oreport.FItemList(i).fspendmile) * dayno / currno)
	else
		maybe_monthsum = 0
	end if
	
	if oreport.FItemList(i).Fsellcnt <> 0 and oreport.FItemList(i).Fsellcnt <> "" then
		maybe_monthcount = CLng(oreport.FItemList(i).Fsellcnt * dayno / currno)
	else
		maybe_monthcount = 0
	end if

	if maybe_monthcount>oreport.maxc then
		oreport.maxc = maybe_monthcount
	end if

	if maybe_monthsum>oreport.maxt then
		oreport.maxt = maybe_monthsum
	end if
	
	if maybe_monthsum <> 0 and maybe_monthsum <> 0 and oreport.maxt <> 0 and oreport.maxt <> "" then
		p3 = Clng(maybe_monthsum/oreport.maxt*100)
	else
		p3 = 0
	end if
	
	if maybe_monthcount <> 0 and maybe_monthcount <> "" and oreport.maxc <> 0 and oreport.maxc <> "" then
		p4 = Clng(maybe_monthcount/oreport.maxc*100)
	else
		p4 = 0
	end if
end if

if oreport.FItemList(i).Fselltotal <> 0 and oreport.FItemList(i).Fselltotal <> "" and oreport.maxt <> 0 and oreport.maxt <> "" then
	p1 = Clng( (oreport.FItemList(i).Fselltotal+oreport.FItemList(i).fspendmile) /oreport.maxt*100)
else
	p1 = 0	
end if

if oreport.FItemList(i).Fsellcnt <> 0 and oreport.FItemList(i).Fsellcnt <> "" and oreport.maxc <> 0 and oreport.maxc <> "" then
	p2 = Clng(oreport.FItemList(i).Fsellcnt/oreport.maxc*100)
else
	p2 = 0	
end if
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover="this.style.background='c1c1c1'" onmouseout="this.style.background='#FFFFFF'">
	<td>
      	<%= oreport.FItemList(i).Fsitename %>주
	</td>
	<td>
      	<% if Left(oreport.FItemList(i).Fsitename,7)=nowyyyymm then %>
      		<div align="left"> <img src="/images/dot4.gif" height="3" width="<%= p3 %>%"></div><br>
      		<div align="left"> <img src="/images/dot4.gif" height="3" width="<%= p4 %>%"></div><br>
      	<% end if %>
      	
		<div align="left"> <img src="/images/dot1.gif" height="3" width="<%= p1 %>%"></div><br>
      		<div align="left"> <img src="/images/dot2.gif" height="3" width="<%= p2 %>%"></div>
      	</td>
	<td bgcolor="#E6B9B8">
	  	<% if Left(oreport.FItemList(i).Fsitename,7)=nowyyyymm then %>
	  		<font color="#AAAAAA"><%= FormatNumber(maybe_monthsum,0) %></font><br>
	  	<% end if %>
	    
	    <%= FormatNumber(oreport.FItemList(i).Fselltotal+oreport.FItemList(i).fspendmile,0) %> <br>
	</td>
	<td>
	  	<% if Left(oreport.FItemList(i).Fsitename,7)=nowyyyymm then %>
	  	<font color="#AAAAAA"><%= FormatNumber(maybe_monthcount,0) %></font><br>
	  	<% end if %>
		
		<%= FormatNumber(oreport.FItemList(i).Fsellcnt,0) %>
	</td>
</tr>
<% 
next

else
%>
<tr bgcolor="#FFFFFF" height=24>
	<td align="center" colspan=15>검색 결과가 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->