<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopcostpermeachulcls.asp"-->
<%
Const isShowIpgoPrice = FALSE

dim i, t

dim yyyy1,mm1,isusing,sysorreal, research, shopid, designer
dim mwgubun
dim vatyn
dim incminusstockyn
dim prevyyyy1, prevmm1
dim yyyy2, mm2, dd2

yyyy1     = RequestCheckVar(request("yyyy1"),10)
mm1       = RequestCheckVar(request("mm1"),10)
isusing   = RequestCheckVar(request("isusing"),10)
sysorreal = RequestCheckVar(request("sysorreal"),10)
research  = RequestCheckVar(request("research"),10)
shopid    = RequestCheckVar(request("shopid"),32)
designer  = RequestCheckVar(request("designer"),32)
mwgubun   = RequestCheckVar(request("mwgubun"),10)
vatyn     = RequestCheckVar(request("vatyn"),10)
incminusstockyn     = RequestCheckVar(request("incminusstockyn"),10)

if (sysorreal="") then sysorreal="sys" ''real
''sysorreal="sys"

if (vatyn="") then vatyn="Y"
if (incminusstockyn="") then incminusstockyn="Y"

dim nowdate
if yyyy1="" then
	'// 전달
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

'// 전달 말일
nowdate = DateAdd("m", 1, (yyyy1 + "-" + mm1 + "-01"))
nowdate = DateAdd("d", -1, nowdate)
yyyy2 = Left(CStr(nowdate),4)
mm2 = Mid(CStr(nowdate),6,2)
dd2 = Right(CStr(nowdate),2)

'// 전전달
nowdate = DateAdd("m", -1, (yyyy1 + "-" + mm1 + "-01"))
prevyyyy1 = Left(CStr(nowdate),4)
prevmm1 = Mid(CStr(nowdate),6,2)



'// ===========================================================================
dim oshopcostpermeachul
set oshopcostpermeachul = new COffShopCostPerMeachul

oshopcostpermeachul.FRectShopID   = shopid
''oshopcostpermeachul.FRectMakerID   = designer
oshopcostpermeachul.FRectYYYYMM   = yyyy1 + "-" + mm1
oshopcostpermeachul.FRectGubun = sysorreal
oshopcostpermeachul.FRectMWDiv    = mwgubun

if (shopid <> "") then
	oshopcostpermeachul.GetOffShopCostPerMeachulByBrandList
end if

dim itemcost, itemcostpermeachul, itemcostpermeachulunit, pointprice
dim totshopbuysumprevmonth, totshopbuysumthismonth, totshopmeachul, totshopmeaip, totshoperrorthismonth
dim totitemcost, totitemcostpermeachul, totshopminusprevmonth, totshopminusthismonth
dim itemrotationrate, itemgainlossrate
dim shopbuysumdiff
dim avgshopbuy
%>
<script language='javascript'>

function PopPrevMonthStock(makerid, mwgubun) {
	var shopid = "<%= shopid %>";
	var yyyy1 = "<%= prevyyyy1 %>";
	var mm1 = "<%= prevmm1 %>";
	var showminus = "<% if (incminusstockyn = "Y") then %>on<% end if %>";

	var popwin = window.open('/admin/newreport/monthlystockShop_detail.asp?research=on&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&shopid=' + shopid + '&makerid=' + makerid + '&sysorreal=<%=sysorreal%>&mwgubun=' + mwgubun + '&showminus=' + showminus,'PopPrevMonthStock','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopThisMonthStock(makerid, mwgubun) {
	var shopid = "<%= shopid %>";
	var yyyy1 = "<%= yyyy1 %>";
	var mm1 = "<%= mm1 %>";
	var showminus = "<% if (incminusstockyn = "Y") then %>on<% end if %>";

	var popwin = window.open('/admin/newreport/monthlystockShop_detail.asp?research=on&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&shopid=' + shopid + '&makerid=' + makerid + '&sysorreal=<%=sysorreal%>&mwgubun=' + mwgubun + '&showminus=' + showminus,'PopPrevMonthStock','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopThisMonthError(makerid) {
	var shopid = "<%= shopid %>";
	var yyyy1 = "<%= yyyy1 %>";
	var mm1 = "<%= mm1 %>";
	var dd1 = "01";
	var yyyy2 = "<%= yyyy2 %>";
	var mm2 = "<%= mm2 %>";
	var dd2 = "<%= dd2 %>";

	var popwin = window.open('/admin/stock/off_baditem_list.asp?shopid=' + shopid + '&makerid=' + makerid + '&itembarcode=&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&dd1=' + dd1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2 + '&dd2=' + dd2 + '&errType=D','PopThisMonthError','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopThisMonthMeachul(makerid) {
	var shopid = "<%= shopid %>";
	var yyyy1 = "<%= yyyy1 %>";
	var mm1 = "<%= mm1 %>";
	var dd1 = "01";
	var yyyy2 = "<%= yyyy2 %>";
	var mm2 = "<%= mm2 %>";
	var dd2 = "<%= dd2 %>";

	var popwin = window.open('/admin/offshop/brandselldetail.asp?yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&dd1=' + dd1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2 + '&dd2=' + dd2 + '&designer=' + makerid + '&datefg=maechul&shopid=' + shopid + '&offgubun=&oldlist=','PopThisMonthMeachul','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopShopMakerMonthlyMaeip(makerid, jungsangubun) {
	var shopid = "<%= shopid %>";
	var yyyymm = "<%= yyyy1 %>-<%= mm1 %>";

	var popwin = window.open('/admin/analysis/popShopMaeip.asp?shopid=' + shopid + '&yyyymm=' + yyyymm + '&makerid=' + makerid + '&jungsangubun=' + jungsangubun,'PopShopMakerMonthlyMaeip','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %> 월 원가율
			&nbsp;&nbsp;
			매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp;
			브랜드 :
			<% drawSelectBoxDesignerwithName "designer", designer %>
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">재고자산 구분:</font>
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >시스템재고
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >실사재고
        	&nbsp;&nbsp;
        	<font color="#CC3333">당월 매입구분:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> > 전체
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> > 매입(매장매입+출고매입)
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> > 위탁(업체위탁+위탁판매)
        	<input type="radio" name="mwgubun" value="B031" <% if mwgubun="B031" then response.write "checked" %> > 출고매입
        	<input type="radio" name="mwgubun" value="B022" <% if mwgubun="B022" then response.write "checked" %> > 매장매입
        	<input type="radio" name="mwgubun" value="B012" <% if mwgubun="B012" then response.write "checked" %> > 업체위탁
        	<input type="radio" name="mwgubun" value="B011" <% if mwgubun="B011" then response.write "checked" %> > 위탁판매
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> > 미지정
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">부가세:</font>
        	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >포함
        	<!--
        	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >제외
        	-->
        	&nbsp;&nbsp;
			<font color="#CC3333">마이너스재고:</font>
        	<input type="radio" name="incminusstockyn" value="Y" <% if incminusstockyn="Y" then response.write "checked" %> >포함
        	<!--
        	<input type="radio" name="incminusstockyn" value="N" <% if incminusstockyn="N" then response.write "checked" %> >제외
        	-->
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>


<br><br><font size=5>수정중입니다.</font><br><br>

<p>

<% if (shopid = "") then %>
	<script>alert("먼저 매장을 선택하세요.");</script>
	<% response.end %>
<% end if %>

* 기초재고, 기말재고 및 당월매입액은 <font color=red>현재 본사매입가, 현재 브랜드, 현재 정산구분</font>을 기준으로 합니다.<br>
* 기초재고, 기말재고는 <font color=red>시스템재고</font>를 기준으로 합니다.<br>
* 당월원가액은 매입상품은 I = A - D(마이너스재고 포함), E = I + B, 위탁상품은 E = B 으로 계산됩니다.<br>
<!--<br>
* 기초재고 및 기말재고는 <font color=red>마이너스재고</font>를 포함한 금액입니다.
-->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>브랜드</td>
    	<td width="80">정산구분</td>
    	<td width="70">소비자매출<br>(C)</td>
    	<td width="70"><%=sysorreal%>기초<br>(A)</td>
    	<td width="70">당월매입<br>(B)</td>

    	<td width="70"><%=sysorreal%>기말<br>(D)</td>

    	<td width="70">당월원가액<br>(E)</td>
    	<td width="70">원가율<br>(F=E/C)</td>
    	<td width="70">손익<br>(1 - F)</td>

    	<td width="70">평균재고<br>(G=(D+A)/2)</td>
    	<td width="70">당월<br>재고회전율<br>(E / G)</td>

    	<td width="70">당월오차(H)</td>
    	<td width="70">재고변동<br>(I)</td>
    	<td></td>
    </tr>
    <%
    totshopbuysumprevmonth = 0
    totshopbuysumthismonth = 0
    totshopmeachul = 0
    totshopmeaip = 0
    totshoperrorthismonth = 0
    totshopminusprevmonth = 0
    totshopminusthismonth = 0
    %>
    <% for i=0 to oshopcostpermeachul.FResultCount-1 %>

    	<% if (designer = "") or (designer = oshopcostpermeachul.FItemList(i).Fmakerid) then %>
    	<% if (mwgubun = "") or (mwgubun = "M") or (mwgubun = "W") or (mwgubun = oshopcostpermeachul.FItemList(i).Fmwdiv) then %>

    	<%

		if (vatyn <> "Y") and (oshopcostpermeachul.FItemList(i).Fvatyn <> "N") then

			'// 부가세 제외
			oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth = oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth * 10 / 11
			oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth = oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth * 10 / 11

			oshopcostpermeachul.FItemList(i).Fshopmeachul = oshopcostpermeachul.FItemList(i).Fshopmeachul * 10 / 11
			oshopcostpermeachul.FItemList(i).Fshopmeaip = oshopcostpermeachul.FItemList(i).Fshopmeaip * 10 / 11

			oshopcostpermeachul.FItemList(i).Fshoperrorthismonth = oshopcostpermeachul.FItemList(i).Fshoperrorthismonth * 10 / 11
			oshopcostpermeachul.FItemList(i).Fshopminusprevmonth = oshopcostpermeachul.FItemList(i).Fshopminusprevmonth * 10 / 11
			oshopcostpermeachul.FItemList(i).Fshopminusthismonth = oshopcostpermeachul.FItemList(i).Fshopminusthismonth * 10 / 11

		end if

		totshopbuysumprevmonth 	= totshopbuysumprevmonth + oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth
		totshopbuysumthismonth 	= totshopbuysumthismonth + oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth
		totshopmeachul 			= totshopmeachul + oshopcostpermeachul.FItemList(i).Fshopmeachul
		totshopmeaip 			= totshopmeaip + oshopcostpermeachul.FItemList(i).Fshopmeaip
		totshoperrorthismonth 	= totshoperrorthismonth + oshopcostpermeachul.FItemList(i).Fshoperrorthismonth
		totshopminusprevmonth 	= totshopminusprevmonth + oshopcostpermeachul.FItemList(i).Fshopminusprevmonth
		totshopminusthismonth 	= totshopminusthismonth + oshopcostpermeachul.FItemList(i).Fshopminusthismonth


		'당월원가액(E = A + B - D or B)
		if (oshopcostpermeachul.FItemList(i).Fmwdiv = "B011") or (oshopcostpermeachul.FItemList(i).Fmwdiv = "B012") or (oshopcostpermeachul.FItemList(i).Fmwdiv = "B013") then
			'위탁
			shopbuysumdiff = 0
			itemcost = oshopcostpermeachul.FItemList(i).Fshopmeaip
		else
			'매입
			shopbuysumdiff = oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth - oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth
			itemcost = shopbuysumdiff + oshopcostpermeachul.FItemList(i).Fshopmeaip
			''((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth + oshopcostpermeachul.FItemList(i).Fshopmeaip) - oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth)
		end if

		itemcostpermeachul = 0
		itemgainlossrate = 0

		'원가율, 손익
		if (oshopcostpermeachul.FItemList(i).Fshopmeachul <= 0) then
			'// 매출없음(원가율 계산불가)
			itemcostpermeachul = "--"
			itemgainlossrate = "--"
		else
			t = (itemcost / oshopcostpermeachul.FItemList(i).Fshopmeachul) * 100.0

			itemcostpermeachul = FormatNumber(t, 1)
			itemgainlossrate = FormatNumber((100.0 - t), 1)
		end if

		' 평균재고자산(마이너스 재고 제외)
		''avgshopbuy = (((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth - oshopcostpermeachul.FItemList(i).Fshopminusprevmonth) + (oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth - oshopcostpermeachul.FItemList(i).Fshopminusthismonth)) / 2)
        avgshopbuy = (((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth ) + (oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth )) / 2)

		'재고회전율
		if (avgshopbuy = 0) or isNULL(avgshopbuy) then
			'// 재고없음(계산불가)
			itemrotationrate = "--"
		else
			''t = (oshopcostpermeachul.FItemList(i).Fshopmeachul / avgshopbuy) * 100.0
			t = (itemcost / avgshopbuy) * 100.0
			itemrotationrate = FormatNumber(t, 1)
		end if

    	%>
    <tr align="center" bgcolor="#FFFFFF" hright=30>
        <td><%= oshopcostpermeachul.FItemList(i).Fmakerid %></td>
        <td><font color="<%= oshopcostpermeachul.FItemList(i).GetDivcdColor %>"><%= oshopcostpermeachul.FItemList(i).Fmwname %></font>
        <%= oshopcostpermeachul.FItemList(i).Fdefaultmargin %>
        </td>
        <td align=right>
        	<acronym title="소비자매출">
        	<a href="javascript:PopThisMonthMeachul('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>')">
        		<%= FormatNumber(oshopcostpermeachul.FItemList(i).Fshopmeachul, 0) %>
        	</a>
        	</acronym>
        </td>

        <td align=right>
        	<acronym title="기초재고">
        	<a href="javascript:PopPrevMonthStock('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>', '<%= oshopcostpermeachul.FItemList(i).Fmwdiv %>')">
				<% if (Not IsNull(oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth)) then %>
					<% if (incminusstockyn = "Y") then %>
						<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth), 0) %>
					<% else %>
						<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth - oshopcostpermeachul.FItemList(i).Fshopminusprevmonth), 0) %>
					<% end if %>
				<% else %>
					<font color="red">ERR</font>
				<% end if %>
       			</a>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="당월매입">
        		<a href="javascript:PopShopMakerMonthlyMaeip('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>', '<%= oshopcostpermeachul.FItemList(i).Fmwdiv %>')">
        			<%= FormatNumber(oshopcostpermeachul.FItemList(i).Fshopmeaip, 0) %>
        		</a>
        	</acronym>
        </td>

        <td align=right>
        	<acronym title="기말재고">
        		<a href="javascript:PopThisMonthStock('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>', '<%= oshopcostpermeachul.FItemList(i).Fmwdiv %>')">
					<% if (incminusstockyn = "Y") then %>
						<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth), 0) %>
					<% else %>
						<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth - oshopcostpermeachul.FItemList(i).Fshopminusthismonth), 0) %>
					<% end if %>
        		</a>
        	</acronym>
        </td>

        <td align=right>
			<% if (Not IsNull(itemcost)) then %>
				<acronym title="당월원가액">
				<%= FormatNumber(itemcost, 0) %>
				</acronym>
			<% end if %>
        </td>
        <!-- 손익금액
        <td align=right>
			<% if (Not IsNull(itemcost)) then %>
				<acronym title="손익">
				<% if (oshopcostpermeachul.FItemList(i).Fshopmeachul - itemcost) < 0 then %><font color=red><% end if %>
				<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopmeachul - itemcost), 0) %>
				</acronym>
			<% end if %>
        </td>
        -->
        <td align=right>
        	<acronym title="원가율">
        		<% if (oshopcostpermeachul.FItemList(i).Fshopmeachul <= 0) then %>
        			--
        		<% else %>
	        		<% if (Abs(itemcostpermeachul) > 500.0) then %>
	        			<font color=red><b>ERR</b></font>
	        		<% else %>
	        			<%= itemcostpermeachul %> %
	        		<% end if %>
	        	<% end if %>
    		</acronym>
        </td>
        <td align=right>
        	<acronym title="손익">
        		<% if (oshopcostpermeachul.FItemList(i).Fshopmeachul <= 0) then %>
        			--
        		<% else %>
	        		<% if (Abs(itemcostpermeachul) > 500.0) then %>
	        			<font color=red><b>ERR</b></font>
	        		<% else %>
		        		<% if ((itemgainlossrate - oshopcostpermeachul.FItemList(i).Fdefaultmargin) > 5.0) or ((itemgainlossrate - oshopcostpermeachul.FItemList(i).Fdefaultmargin) < -10.0) then %><font color=red><b><% end if %>
		        		<%= itemgainlossrate %> %
	        		<% end if %>
	        	<% end if %>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="평균재고자산">
        		<% if (FALSE) and (oshopcostpermeachul.FItemList(i).Fmwdiv <> "B031") and (oshopcostpermeachul.FItemList(i).Fmwdiv <> "B022") then %>
        			--
        		<% elseif (Not IsNull(avgshopbuy)) then %>
        			<%= FormatNumber(avgshopbuy, 0) %>
        		<% end if %>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="당월재고회전율">
        		<% if (avgshopbuy = 0) then %>
        			--
        		<% else %>
	        		<% if (FALSE) and (oshopcostpermeachul.FItemList(i).Fmwdiv <> "B031") and (oshopcostpermeachul.FItemList(i).Fmwdiv <> "B022") then %>
	        			--
	        		<% elseif (Not IsNull(itemrotationrate)) then %>
		        		<% if (Abs(itemrotationrate) >= 1000.0) then %>
		        			<font color=red><b>ERR</b></font> <%'=itemcostpermeachul%>
		        		<% else %>
		        			<% if (itemrotationrate < 30.0) then %><font color=red><b><% end if %>
		        			<% if (itemrotationrate > 60.0) then %><font color=green><% end if %>
		        			<%= itemrotationrate %> %
		        		<% end if %>
	        		<% end if %>
        		<% end if %>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="당월오차">
        	<a href="javascript:PopThisMonthError('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>')">
        		<%= FormatNumber(oshopcostpermeachul.FItemList(i).Fshoperrorthismonth, 0) %>
        	</a>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="재고변동(기초-기말, 마이너스 재고 포함)">
        		<% if (shopbuysumdiff = 0) then %>
        			--
        		<% else %>
        			<%= FormatNumber(shopbuysumdiff, 0) %>
        		<% end if %>
        	</acronym>
        </td>

    	<td></td>
    </tr>
    	<% end if %>
    	<% end if %>
    <% next %>
    <%
	totitemcost = totshopbuysumprevmonth + totshopmeaip - totshopbuysumthismonth
    %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td >총계</td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totshopmeachul, 0) %></td>
    	<td align="right" ><%= FormatNumber((totshopbuysumprevmonth - totshopminusprevmonth), 0) %></td>
    	<td align="right" ><%= FormatNumber(totshopmeaip, 0) %></td>

    	<td align="right" ><%= FormatNumber((totshopbuysumthismonth ), 0) %></td>
    	<td align="right" ></td>
    	<!--
    	<td align="right" ><%= FormatNumber((totshopmeachul - totitemcost), 0) %></td>
    	-->
    	<td align="right" >
    		<!--
    		<% if (totshopmeachul <> 0) then %>
    			<%= FormatNumber((totitemcost / totshopmeachul) * 100, 1) %>%
    		<% else %>
    			--
    		<% end if %>
    		-->
    	</td>
    	<td align="right" ></td>
    	<td align="right" ></td>
    	<td align="right" ></td>
    	<td align="right" ><%= FormatNumber(totshoperrorthismonth, 0) %></td>
    	<td align="right" ><%= FormatNumber((totshopbuysumprevmonth - totshopbuysumthismonth), 0) %></td>
    	<td align="right" ></td>
    </tr>
</table>
<%
set oshopcostpermeachul = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
