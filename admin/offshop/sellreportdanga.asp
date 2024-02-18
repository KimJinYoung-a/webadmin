<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 기간별객단가
' History : 2009.04.07 서동석 생성
'			2010.05.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim page,shopid ,oldlist ,fromDate,toDate ,yyyymmdd1,yyymmdd2 ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,i
dim inc3pl
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	oldlist = requestCheckVar(request("oldlist"),2)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if page="" then page=1
	
if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-14)
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

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if	

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopID = shopid
	ooffsell.FCurrPage=page
	ooffsell.FRectOldData = oldlist
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectInc3pl = inc3pl
	
	if shopid<>"" then
		ooffsell.GetReportByDanga
	else
		response.write "<script language='javascript'>"
		response.write "	alert('매장을 선택해 주세요');"
		response.write "</script>"
	end if
%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<!--
			<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3개월이전내역
			&nbsp;
		-->
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
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<Br>
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
		검색결과 : <b><%= ooffsell.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>구분</td>
	<td>매출건수</td>
	<td>총건수대비%</td>
	<td>매출액<Br>(마일리지포함)</td>
	<td>총매출대비%</td>
</tr>
<% if ooffsell.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
<% 
for i=0 to ooffsell.FresultCount-1
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff'; align="center">
	<td><%= ooffsell.FItemList(i).FTerm %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FCount,0) %></td>
	<td align="right">
		<% if ooffsell.maxc<>0 then %>
			<%= CLng(ooffsell.FItemList(i).FCount/ooffsell.maxc*100) %> %
		<% end if %>
	</td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell.FItemList(i).FSum+ooffsell.FItemList(i).fspendmile,0) %></td>
	<td align="right">
		<% if ooffsell.maxt<>0 then %>
			<%= CLng( (ooffsell.FItemList(i).FSum+ooffsell.FItemList(i).fspendmile) /ooffsell.maxt*100) %> %
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="right">
	<td align="center">총계</td>
	<td align="right"><%= FormatNumber(ooffsell.maxc,0) %></td>
	<td></td>
	<td align="right"><%= FormatNumber(ooffsell.maxt,0) %></td>
	<td></td>
</tr>
<% end if %>
</table>

<%
set ooffsell= Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->