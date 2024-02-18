<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 상품 베스트셀러
		'[OFF]오프_통계관리>>요일별매출분석 /admin/offshop/weeklysellreport.asp 에서도 사용
' History : 2010.05.18 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offshop_reportcls.asp"-->
<%
dim yyyy1, mm1, dd1, yyyy2, mm2, dd2 , oldlist , yyyymmdd1, yyymmdd2 ,nowdate, searchnextdate
dim orderserial, itemid, oreport ,topn,cdl,cdm,page ,ckpointsearch, ckipkumdiv4 ,i, iy, cknodate
dim order_desum , ordertype , shopid , offgubun ,totalsumprice, totalbuyprice, totalitemno ,totsellsum
dim datefg,cds ,searchgubun ,makerid ,weekdate, buyergubun, inc3pl
	buyergubun = requestCheckVar(request("buyergubun"),10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	cdl = requestCheckVar(request("cdl"),3)
	cdm = requestCheckVar(request("cdm"),3)
	cds = requestCheckVar(request("cds"),3)
	orderserial = requestCheckVar(request("orderserial"),16)
	itemid = requestCheckVar(request("itemid"),10)
	topn = requestCheckVar(request("topn"),10)
	ckpointsearch = requestCheckVar(request("ckpointsearch"),10)
	cknodate = requestCheckVar(request("cknodate"),10)
	order_desum = requestCheckVar(request("order_desum"),10)
	ordertype = requestCheckVar(request("ordertype"),10)
	if ordertype="" then ordertype="ea"
	oldlist = requestCheckVar(request("oldlist"),10)
	shopid = requestCheckVar(request("shopid"),32)
	offgubun = requestCheckVar(request("offgubun"),10)
	datefg = requestCheckVar(request("datefg"),32)
	searchgubun = requestCheckVar(request("searchgubun"),10)
	makerid = RequestCheckVar(request("makerid"),32)
	weekdate = RequestCheckVar(request("weekdate"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if searchgubun = "" then searchgubun = "I"
if datefg = "" then datefg = "maechul"	
if (topn="") then topn=500
if (yyyy1="") then
	nowdate = Left(CStr(now()),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

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
		makerid = session("ssBctID")	'"GREENBEE_1"
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

set oreport = new COffshopReport
	oreport.FRectmakerid = makerid	
	oreport.frectdatefg = datefg
	oreport.FRectCDL = cdl
	oreport.FRectCDM = cdm
	oreport.FRectCDN = cds
	oreport.FPageSize = topn
	oreport.FCurrPage = page	
	oreport.FRectOrdertype = ordertype
	oreport.FRectOldJumun = oldlist
	oreport.FRectOffgubun = offgubun
	oreport.FRectShopID = shopid		
	oreport.frectsearchgubun = searchgubun
	oreport.frectweekdate = weekdate
	oreport.frectbuyergubun = buyergubun
	oreport.FRectInc3pl = inc3pl
	
	if cknodate="" then
		oreport.FRectFromDate = yyyy1 + "-" + mm1 + "-" + dd1
		oreport.FRectToDate = searchnextdate
	end if
	
	oreport.SearchCategoryBestseller

totalsumprice = 0
totalbuyprice = 0
totalitemno = 0
totsellsum = 0
%>

<script language='javascript'>

function ViewOrderDetail(itemid){
    var popwin = window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"category_prd");
    popwin.focus();
}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function ReSearch(page){
	var v = frm.topn.value;
	if (!IsDigit(v)){
		alert('숫자만 가능합니다.');
		frm.topn.focus();
		return;
	}

	if (v>3000){
		alert('3천건 이하만 검색가능합니다.');
		frm.topn.focus();
		return;
	}

	document.frm.page.value= page;
	frm.submit();
}

function chsearchgubun(makerid){
	frm.makerid.value=makerid;
	frm.searchgubun(0).checked = true;
	frm.submit();
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
		* 기간 : <% drawmaechul_datefg "datefg" ,datefg ,""%> 
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> onClick="ReSearch('');">3년이전	
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
			<% if not(C_IS_Maker_Upche) then %> 
				* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
			<% else %>
				* 매장 : <% drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," onchange='ReSearch("""");'","" %>
			<% end if %>
		<% end if %>
		<p>
		* 브랜드 : 
		<% if (C_IS_Maker_Upche) then %>
			<%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
		<% else %>
			<% drawSelectBoxDesignerwithName "makerid",makerid %>
		<% end if %>
		<% if not(C_IS_Maker_Upche) then %>
			&nbsp;&nbsp;
			* 매장 구분 : <% Call DrawShopDivCombo("offgubun",offgubun) %>
		<% end if %>
		&nbsp;&nbsp;
		* 요일:<% drawweekday_select "weekdate" , weekdate ," onchange='ReSearch("""");'" %>		
		&nbsp;&nbsp;
		* 국적구분: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='ReSearch("""");'" %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="ReSearch('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
        &nbsp;&nbsp;        
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
		※ 야간판매를 하는 매장(두타점등등)의 경우 매출기준일로 검색을 하셔야 정확한 매출이 집계 됩니다.
		<br>&nbsp;&nbsp;판매내역은 샾 마감후, 새벽판매매장(새벽5시경), 주간판매매장(오후 10시경) 업데이트 되며,
		<br>&nbsp;&nbsp;정산은 주문일 기준으로 정산 됩니다.	        	      	
    </td>
    <td align="right">
		<input type="radio" name="searchgubun" value="I" <% if searchgubun="I" then response.write " checked" %> onClick="ReSearch('');">상품기준
		<input type="radio" name="searchgubun" value="B" <% if searchgubun="B" then response.write " checked" %> onClick="ReSearch('');">브랜드기준    	
		/ 검색갯수:
		<input type="text" name="topn" value="<%= topn %>" size="7" maxlength="6" >
		정렬:
		<% drawordertype "ordertype" ,ordertype ," onchange='ReSearch("""");'" ,searchgubun  %> 
    </td>        
</tr>
</form>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="left">
		검색결과 : <b><%= oreport.ftotalcount %></b>
	</td>
</tr>

<%
'/상품기준
if searchgubun = "I" then
%>
	<tr bgcolor="#EEEEEE" align="center">
		<td>순위</td>
		<td>상품번호</td>
		<td>상품</td>
		<td>단가</td>
		<td>브랜드ID</td>
		<td>옵션</td>
		<td>바코드</td>
		<td>범용바코드</td>		
		<td>판매수량</td>
		<td>판매액</td>
		<td>매출액</td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			<td>매입액</td>
		<% end if %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td>수익</td>
			<td>마진율</td>
		<% end if %>
	</tr>
	<% if oreport.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center"  >[검색결과가 없습니다.]</td>
	</tr>
	<% else %>
	<% 
	for i=0 to oreport.FResultCount -1
	
	totalitemno   =  totalitemno + oreport.FItemList(i).FItemNo
	totalsumprice =  totalsumprice + oreport.FItemList(i).Fselltotal
	totalbuyprice =  totalbuyprice + oreport.FItemList(i).Fbuytotal
	totsellsum = totsellsum + oreport.FItemList(i).ftotsellsum
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%=i+1%></td>
		<td height="25">
			<% if oreport.FItemList(i).Fitemgubun = "10" then %>
				<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oreport.FItemList(i).FItemID %>" class="zzz" target="_blank"><%= oreport.FItemList(i).FItemID  %></a>
			<% else %>
				<%= oreport.FItemList(i).FItemID  %>
			<% end if %>
		</td>
		<td><%= oreport.FItemList(i).FItemName %></td>
		<td><%= FormatNumber(oreport.FItemList(i).fsellprice,0) %></td>
		<td><%= oreport.FItemList(i).FMakerid %></td>
		<% if (oreport.FItemList(i).fitemoptionname="") then %>
			<td>&nbsp;</td>
		<% else %>
			<td><%= oreport.FItemList(i).fitemoptionname %></td>
		<% end if %>
		<td>
			<%= oreport.FItemList(i).GetBarCode %>
		</td>
		<td>
			<%= oreport.FItemList(i).fextbarcode %>
		</td>		
		<td><%= FormatNumber(oreport.FItemList(i).FItemNo,0) %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(i).ftotsellsum,0) %></td>
		<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(oreport.FItemList(i).Fselltotal,0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			<td align="right"><%= FormatNumber(oreport.FItemList(i).Fbuytotal,0) %></td>
		<% end if %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><%= FormatNumber(oreport.FItemList(i).Fselltotal-oreport.FItemList(i).Fbuytotal,0) %></td>
		    <td align="center">
		        <% if oreport.FItemList(i).Fselltotal<>0 then %>
		        <%= 100-CLng(oreport.FItemList(i).Fbuytotal/oreport.FItemList(i).Fselltotal*100*100)/100 %> %
		        <% end if %>
		    </td>
		<% end if %>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" align="center">
	    <td colspan="8"></td>
	    <td><%= FormatNumber(totalitemno,0) %></td>
	    <td align="right"><%= FormatNumber(totsellsum,0) %></td>
	    <td align="right"><%= FormatNumber(totalsumprice,0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			<td align="right"><%= FormatNumber(totalbuyprice,0) %></td>
		<% end if %>

	    <% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		    <td align="right"><%= FormatNumber(totalsumprice-totalbuyprice,0) %></td>
		    <td>
		        <% if totalsumprice<>0 then %>
		        <%= 100-CLng(totalbuyprice/totalsumprice*100*100)/100 %> %
		        <% end if %>
		    </td>
		<% end if %>
	</tr>
	<% end if %>
<%
'/브랜드기준
elseif searchgubun = "B" then
%>
	<tr bgcolor="#EEEEEE" align="center">
		<td>순위</td>
		<td>브랜드ID</td>
		<td>판매수량</td>
		<td>판매액</td>
		<td>매출액</td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			<td>매입액</td>
		<% end if %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td>수익</td>
			<td>마진율</td>
		<% end if %>
	</tr>
	<% if oreport.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center"  >[검색결과가 없습니다.]</td>
	</tr>
	<% else %>
	<% 
	for i=0 to oreport.FResultCount -1
	
	totalitemno   =  totalitemno + oreport.FItemList(i).FItemNo
	totalsumprice =  totalsumprice + oreport.FItemList(i).Fselltotal
	totalbuyprice =  totalbuyprice + oreport.FItemList(i).Fbuytotal
	totsellsum = totsellsum + oreport.FItemList(i).ftotsellsum
	%>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%=i+1%></td>
		<td align="center">
			<a href="javascript:chsearchgubun('<%= oreport.FItemList(i).FMakerid %>');" onfocus="this.blur();"><%= oreport.FItemList(i).FMakerid %></a>
		</td>
		<td align="center"><%= FormatNumber(oreport.FItemList(i).FItemNo,0) %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(i).ftotsellsum,0) %></td>
		<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(oreport.FItemList(i).Fselltotal,0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			<td align="right"><%= FormatNumber(oreport.FItemList(i).Fbuytotal,0) %></td>
		<% end if %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><%= FormatNumber(oreport.FItemList(i).Fselltotal-oreport.FItemList(i).Fbuytotal,0) %></td>
		    <td align="center">
		        <% if oreport.FItemList(i).Fselltotal<>0 then %>
		        <%= 100-CLng(oreport.FItemList(i).Fbuytotal/oreport.FItemList(i).Fselltotal*100*100)/100 %> %
		        <% end if %>
		    </td>
		<% end if %>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" align="center">
	    <td colspan="2"></td>
	    <td align="center"><%= FormatNumber(totalitemno,0) %></td>
	    <td align="right"><%= FormatNumber(totsellsum,0) %></td>
	    <td align="right"><%= FormatNumber(totalsumprice,0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			<td align="right"><%= FormatNumber(totalbuyprice,0) %></td>
		<% end if %>

	    <% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		    <td align="right"><%= FormatNumber(totalsumprice-totalbuyprice,0) %></td>
		    <td align="center">
		        <% if totalsumprice<>0 then %>
		        <%= 100-CLng(totalbuyprice/totalsumprice*100*100)/100 %> %
		        <% end if %>
		    </td>
		<% end if %>
	</tr>
	<% end if %>
<% end if %>
</table>

<%
	set oreport = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->