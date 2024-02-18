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
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->
<%
Dim ozone,i,page , parameter , shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,fromDate ,toDate
dim designer , sellgubun ,cdl ,cdm ,cds ,datefg
	designer = RequestCheckVar(request("designer"),32)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	sellgubun = requestCheckVar(request("sellgubun"),1)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
	datefg = requestCheckVar(request("datefg"),10)

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

'직영/가맹점
if (C_IS_SHOP) then
	
	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if	
end if

set ozone = new czone_list
	ozone.FPageSize = 500
	ozone.FCurrPage = page
	ozone.frectshopid = shopid
	ozone.FRectStartDay = fromDate
	ozone.FRectEndDay = toDate
	ozone.FRectmakerid = designer	
	ozone.FRectCDL = cdl
	ozone.FRectCDM = cdm
	ozone.FRectCDN = cds	
	ozone.frectdatefg = datefg
	ozone.frectsellgubun = sellgubun

	if shopid <> "" then
		ozone.Getoffshopzonesum
	end if
	
	if shopid = "" then response.write "<script>alert('매장을 선택해주세요');</script>"
		
parameter = "shopid="&shopid&"&sellgubun="&sellgubun&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&menupos="&menupos&"&designer="&designer&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds
%>

<script language="javascript">

//그룹설정
function zone_groupreg(){
	var zone_groupreg = window.open('/admin/offshop/zone/zone_common.asp?menupos=<%=menupos%>','zone_groupreg','width=1024,height=768,scrollbars=yes,resizable=yes');
	zone_groupreg.focus();
}

//매장구역설정
function zone_reg(){
	var zone_reg = window.open('/admin/offshop/zone/zone.asp?menupos=<%=menupos%>','zone_reg','width=1024,height=768,scrollbars=yes,resizable=yes');
	zone_reg.focus();
}

//구역별상품등록
function zone_item(){

	if (frm.shopid.value==''){
		alert('매장을 선택해 주세요');
		return;
	}
		
	var zone_item = window.open('zone_item.asp?menupos=<%=menupos%>&shopid=<%=shopid%>&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>&datefg=<%=datefg%>','zone_item','width=1024,height=768,scrollbars=yes,resizable=yes');
	zone_item.focus();
}

//상세매출
function item_detail(idx,searchtype){

	var item_detail = window.open('zone_sum_detail.asp?idx='+idx+'&searchtype='+searchtype+'&<%=parameter%>','item_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
	item_detail.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		매장:<% drawSelectBoxOffShop "shopid",shopid %>
		매출기준 :
		<% drawmaechul_datefg "datefg" ,datefg ,""%> 
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<Br><!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		브랜드 : <% drawSelectBoxDesignerwithName "designer",designer %>
		<input type="radio" name="sellgubun" value="S" <% if sellgubun="S" then response.write " checked" %>>결제내역기준
		<input type="radio" name="sellgubun" value="N" <% if sellgubun="N" then response.write " checked" %>>현재등록내역기준
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ [사용법] "그룹설정" 에서 매장에 존재하는 그룹을 먼저 설정후 , "샵구역설정" 에서 각매장별로 구역을 설정 하신후
		<br>"샵상품구역설정" 에서 샵구역별로 상품을 넣으셔야 매출이 집계됩니다.
		<%
		'/결제내역기준
		if sellgubun="S" then
		%>
			<br>[참고] 포스에서 결제가 되면, 해당 상품이 등록된 구역 내역이 저장 되며, 이를 기준으로 통계가 산출 됩니다
			<br>그러므로 구역에 상품을 등록하지 않으면, 통계가 남지 않습니다.
		<%
		'/현재등록내역기준
		else
		%>
			<br>[참고] 현재 구역별 상품 등록에 등록되어져 있는 내역 입니다. 상품별 예전 내역은 노출 되지 않습니다.
		<% end if %>
	</td>
	<td align="right">
		<input type="button" class="button" value="그룹설정" onclick="zone_groupreg();">
		<input type="button" class="button" value="매장구역설정" onclick="zone_reg();">
		<input type="button" class="button" value="구역별상품등록" onclick="zone_item();">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ozone.FTotalCount %></b>
		※ 500건 까지 검색됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">그룹</td>
	<td align="center">매대타입</td>
	<td align="center">상세구역명</td>
	<td>총<br>판매수량</td>
	<td>총<br>실매출액</td>
	<td>점유율</td>
	<td>UNIT당 매출<br>UNIT</td>
	<td>비고</td>
</tr>
<% if ozone.FTotalCount>0 then %>
<% for i=0 to ozone.FTotalCount-1 %>
<% if ozone.FItemList(i).fzonename <> "" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>
	<td>
		<% if ozone.FItemList(i).fzonename = "" then %>
			-
		<% else %>
			<%= getOffShopzonegroup(ozone.FItemList(i).fzonegroup) %>
		<% end if %>
	</td>
	<td>
		<% if ozone.FItemList(i).fzonename = "" then %>
			-
		<% else %>
			<%= getOffShopracktype(ozone.FItemList(i).fracktype) %>
		<% end if %>
	</td>
	<td>
		<% if ozone.FItemList(i).fzonename = "" then %>
			-
		<% else %>
			<%= ozone.FItemList(i).fzonename %>
		<% end if %>
	</td>	
	<td>
		<%= ozone.FItemList(i).fitemcnt %>
	</td>
	<td>
		<%= FormatNumber(ozone.FItemList(i).fsellsum,0) %>
	</td>
	<td>
		<% if ozone.FSumTotal<>0 then %>
			<%= Clng( ((ozone.FItemList(i).fsellsum / ozone.FSumTotal) * 10000)) / 100 %> %
		<% end if %>
	</td>
	<td>
		<% if ozone.FItemList(i).funitvalue <> "" then %>
			<%= FormatNumber(ozone.FItemList(i).funitvalue,0) %>
		<% else %>
			-
		<% end if %>
		<% if ozone.FItemList(i).funit <> "" then %>
			<br><%= FormatNumber(ozone.FItemList(i).funit,0) %>
		<% else %>
			<br>-
		<% end if %>
	</td>
	<td width=200>
		<input type="button" class="button" value="상품상세" onclick="item_detail('<%= ozone.FItemList(i).fidx %>','I');">
		<input type="button" class="button" value="카테고리상세" onclick="item_detail('<%= ozone.FItemList(i).fidx %>','C');">
	</td>	
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan=3>
		합계
	</td>		
	<td align="center">
		<%= FormatNumber(ozone.FCountTotal,0) %>
	</td>
		
	<td align="center">
		<%= FormatNumber(ozone.FSumTotal,0) %>
	</td>
	<td align="center" colspan=4></td>	
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

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->