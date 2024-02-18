<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 할인 유형별 통계
' History : 2012.02.10 한용민 수정
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
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2 , fromDate,toDate , shopid ,i ,makerid ,datefg , tmpdate , totsellcnt
dim inc3pl
	makerid = requestCheckVar(request("makerid"),32)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"

tmpdate = dateadd("d",-1,date)

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
	oreport.FPageSize = 500
	oreport.FRectFromDate = fromDate
	oreport.FRectToDate = toDate
	oreport.frectdatefg = datefg
	oreport.FRectShopID = shopid
	oreport.frectmakerid = makerid
	oreport.FRectInc3pl = inc3pl	
	oreport.Getsalepaysum

%>

<script language='javascript'>

function detailitem(discountKind,shopid,yyyy1,mm1,dd1,yyyy2,mm2,dd2,datefg,makerid){
	var detailitem = window.open('/admin/offshop/maechul/salepaysum_detail.asp?menupos=<%=menupos%>&makerid='+makerid+'&shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&datefg='+datefg+'&discountKind='+discountKind,'detailitem','width=1024,height=768,scrollbars=yes,resizable=yes');
	detailitem.focus();
}

function reg(){
	frm.submit();
}

</script>

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
				* 기간 : <% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
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
</form>
</table>
<!-- 표 중간바 끝-->

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        검색결과 : <b><%= oreport.FTotalcount %></b>
    </td>
</tr>
<tr bgcolor="#EEEEEE" align="center">
	<td>매장</td>
	<td>직원<br>할인</td>
	<td>샘플<br>할인</td>
	<td>패키지<br>할인</td>
	<td>고객<br>5%할인</td>
	<td>고객<br>할인</td>
	<td>tenday<br>10%할인</td>
	<td>두타쿠폰북<br>할인</td>
</tr>
<%
if oreport.FResultCount > 0 then
	
for i=0 to oreport.FResultCount - 1
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='#FFFFFF'; align="center">
	<td>
		<%= oreport.FItemList(i).fshopname %>
	</td>
	<td align="right">
		<% if oreport.FItemList(i).f1sellcnt <> 0 then %>
			<a href="javascript:detailitem('1','<%= oreport.FItemList(i).fshopid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%=datefg%>','<%= makerid %>');">
			<%= FormatNumber(oreport.FItemList(i).f1sellprice-oreport.FItemList(i).f1realsellprice,0) %>
			<Br><%= FormatNumber(oreport.FItemList(i).f1sellcnt,0) %> 건
			</a>
		<% else %>
		
		<% end if %>
	</td>	
	<td align="right">
		<% if oreport.FItemList(i).f2Sellcnt <> 0 then %>
			<a href="javascript:detailitem('2','<%= oreport.FItemList(i).fshopid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%=datefg%>','<%= makerid %>');">
			<%= FormatNumber(oreport.FItemList(i).f2sellprice-oreport.FItemList(i).f2realsellprice,0) %>
			<Br><%= FormatNumber(oreport.FItemList(i).f2sellcnt,0) %> 건
			</a>
		<% else %>
		
		<% end if %>			
	</td>
	<td align="right">
		<% if oreport.FItemList(i).f4sellcnt <> 0 then %>
			<a href="javascript:detailitem('4','<%= oreport.FItemList(i).fshopid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%=datefg%>','<%= makerid %>');">
			<%= FormatNumber(oreport.FItemList(i).f4sellprice-oreport.FItemList(i).f4realsellprice,0) %>
			<Br><%= FormatNumber(oreport.FItemList(i).f4sellcnt,0) %> 건
			</a>
		<% else %>
		
		<% end if %>		
	</td>
	<td align="right">
		<% if oreport.FItemList(i).f5sellcnt <> 0 then %>
			<a href="javascript:detailitem('5','<%= oreport.FItemList(i).fshopid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%=datefg%>','<%= makerid %>');">
			<%= FormatNumber(oreport.FItemList(i).f5sellprice-oreport.FItemList(i).f5realsellprice,0) %>
			<Br><%= FormatNumber(oreport.FItemList(i).f5sellcnt,0) %> 건
			</a>
		<% else %>
		
		<% end if %>			
	</td>
	<td align="right">
		<% if oreport.FItemList(i).f6sellcnt <> 0 then %>
			<a href="javascript:detailitem('6','<%= oreport.FItemList(i).fshopid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%=datefg%>','<%= makerid %>');">
			<%= FormatNumber(oreport.FItemList(i).f6sellprice-oreport.FItemList(i).f6realsellprice,0) %>
			<Br><%= FormatNumber(oreport.FItemList(i).f6sellcnt,0) %> 건
			</a>
		<% else %>
		
		<% end if %>			
	</td>
	<td align="right">
		<% if oreport.FItemList(i).f7sellcnt <> 0 then %>
			<a href="javascript:detailitem('7','<%= oreport.FItemList(i).fshopid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%=datefg%>','<%= makerid %>');">
			<%= FormatNumber(oreport.FItemList(i).f7sellprice-oreport.FItemList(i).f7realsellprice,0) %>
			<Br><%= FormatNumber(oreport.FItemList(i).f7sellcnt,0) %> 건
			</a>
		<% else %>
		
		<% end if %>			
	</td>
	<td align="right">
		<% if oreport.FItemList(i).f9sellcnt <> 0 then %>
			<a href="javascript:detailitem('9','<%= oreport.FItemList(i).fshopid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%=datefg%>','<%= makerid %>');">
			<%= FormatNumber(oreport.FItemList(i).f9sellprice-oreport.FItemList(i).f9realsellprice,0) %>
			<Br><%= FormatNumber(oreport.FItemList(i).f9sellcnt,0) %> 건
			</a>
		<% else %>
		
		<% end if %>			
	</td>
</tr>
<% next %>

<% else %>
<tr bgcolor="#FFFFFF" height=24>
	<td align="center" colspan=25>검색 결과가 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->