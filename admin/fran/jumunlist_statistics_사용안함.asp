<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 주문서관리 통계
' History : 2010.06.11 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim page, shopid ,designer, statecd, baljucode ,notipgo, minusjumun, shopdiv ,totaljumunsuply, totalfixsuply
dim yyyy1,mm1 ,dd1,yyyy2,mm2,dd2, totaljumunsellcash ,i ,fromDate ,toDate , datefg ,parameter , arridx
dim tot10totalsellcash ,tot90totalsellcash , tot70totalsellcash , tot10totalsuplycash , tot90totalsuplycash ,tot70totalsuplycash
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	designer = request("designer")
	statecd  = request("statecd")
	baljucode= request("baljucode")
	notipgo = request("notipgo")
	minusjumun = request("minusjumun")
	shopdiv = request("shopdiv")
	shopid = request("shopid")
	page = request("page")
	menupos = request("menupos")
	datefg = request("datefg")	

	if page="" then page=1
				
if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))-1
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
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
		designer = session("ssBctID")	'"7321"

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if	

dim osheet
set osheet = new COrderSheet
	osheet.FRectFromDate = fromDate
	osheet.FRectToDate = toDate
	osheet.frectdatefg = datefg
	osheet.FCurrPage = page
	osheet.Fpagesize=20
	
	if baljucode<>"" then
		osheet.FRectBaljuCode = baljucode
	else
		osheet.FRectBaljuid = shopid
		osheet.FRectStatecd = statecd
		osheet.FRectMakerid = designer
		osheet.FRectDivCodeArr = "('501','502','503')"
		osheet.FRectNotIpgoOnly = notipgo
		osheet.FRectMinusOnly = minusjumun
		osheet.FRectshopdiv = shopdiv
	end if
	
	osheet.GetOrderSheetList_statistics
	
parameter = "shopid="&shopid&"&notipgo="& notipgo &"&minusjumun="& minusjumun &"&shopdiv="& shopdiv &"&statecd="& statecd &"&designer="& designer &"&menupos="& menupos
parameter = parameter & "&yyyy1="&yyyy1&"&mm1="& mm1&"&dd1="& dd1&"&yyyy2="& yyyy2&"&mm2="& mm2&"&dd2="& dd2&"&baljucode="& baljucode&"&datefg="& datefg

tot10totalsellcash = 0
tot90totalsellcash = 0
tot70totalsellcash = 0
tot10totalsuplycash = 0
tot90totalsuplycash = 0
tot70totalsuplycash = 0
%>

<script language='javascript'>

function pop_detail(idx){
	var pop_detail = window.open('jumunlist_statistics_detail.asp?idx='+idx,'pop_detail','width=1024,height=768,resizable=yes,scrollbars=yes');
	pop_detail.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드포함 : <% drawSelectBoxDesignerwithName "designer", designer %>
		<br>
		주문코드 : <input type="text" class="text" name="baljucode" value="<%= baljucode %>" size="10" maxlength="8">
		주문상태 :
		<select name="statecd" class="select">
			<option value="">전체
			<option value=" " <% if statecd=" " then response.write "selected" %> >작성중
			<option value="0" <% if statecd="0" then response.write "selected" %> >주문접수
			<option value="1" <% if statecd="1" then response.write "selected" %> >주문확인
			<option value="2" <% if statecd="2" then response.write "selected" %> >입금대기
			<option value="5" <% if statecd="5" then response.write "selected" %> >배송준비
			<option value="6" <% if statecd="6" then response.write "selected" %> >출고대기
			<option value="7" <% if statecd="7" then response.write "selected" %> >출고완료
			<option value="8" <% if statecd="8" then response.write "selected" %> >입고대기
			<option value="9" <% if statecd="9" then response.write "selected" %> >입고완료
		</select>
		<input type="checkbox" name="notipgo" <% if notipgo="on" then response.write "checked" %> >출고미처리만
     	&nbsp;
     	<input type="checkbox" name="minusjumun" <% if minusjumun="on" then response.write "checked" %> >마이너스주문만
     	&nbsp;     	
     	SHOP구분 : 
     	<input type="radio" name="shopdiv" value="" <% if shopdiv="" then response.write "checked" %> >전체
		<input type="radio" name="shopdiv" value="j" <% if shopdiv="j" then response.write "checked" %> >직영
		<input type="radio" name="shopdiv" value="f" <% if shopdiv="f" then response.write "checked" %> >가맹점			
	</td>		
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>	
			<% if not C_IS_OWN_SHOP and shopid <> "" then %>
				매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				매장 : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			매장 : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>					
		날짜기준 : 
		<% drawipgo_datefg "datefg" ,datefg ,""%> 
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
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
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type=hidden name="idxarr">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= osheet.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= osheet.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>주문코드</td>
	<td>공급자</td>
	<td>공급받는자<br>주문자(SHOP)</td>
	<td>주문상태</td>
	<td>주문일/<br>입고(요청)일</td>
	<td>총주문액<br>(소비자가)</td>
	<td>총주문액<br>(공급가)</td>
	<td>확정금액<br>(공급가)</td>
	<td>출고일</td>
	<td>상품구분별<br>확정금액(판매가)</td>
	<td>상품구분별<br>확정금액(공급가)</td>	
	<td>비고</td>
</tr>
<% if osheet.FResultCount >0 then %>
<% 
for i=0 to osheet.FResultcount-1

arridx = arridx & osheet.FItemList(i).fidx & ","

tot10totalsellcash = tot10totalsellcash + osheet.FItemList(i).f10totalsellcash
tot90totalsellcash = tot90totalsellcash + osheet.FItemList(i).F90totalsellcash
tot70totalsellcash = tot70totalsellcash + osheet.FItemList(i).F70totalsellcash
tot10totalsuplycash = tot10totalsuplycash + osheet.FItemList(i).F10totalsuplycash
tot90totalsuplycash = tot90totalsuplycash + osheet.FItemList(i).F90totalsuplycash
tot70totalsuplycash = tot70totalsuplycash + osheet.FItemList(i).F70totalsuplycash
totaljumunsellcash = totaljumunsellcash + osheet.FItemList(i).Fjumunsellcash

if osheet.FItemList(i).Ftargetid="10x10" then
	totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunsuplycash
	totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalsuplycash
else
	totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunbuycash
	totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalbuycash
end if
%>
<tr bgcolor="#FFFFFF" align="center">
	<td rowspan=2 >
		<%= osheet.FItemList(i).Fbaljucode %>
	</td>
	<% if osheet.FItemList(i).Ftargetid<>"10x10" then %>
	<td rowspan=2 ><b><%= osheet.FItemList(i).Ftargetid %></b><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
	<% else %>
	<td rowspan=2 ><%= osheet.FItemList(i).Ftargetid %><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
	<% end if %>
	<td rowspan=2 ><%= osheet.FItemList(i).Fbaljuid %><br>(<%= osheet.FItemList(i).Fbaljuname %>)</td>
	<td rowspan=2 >
		<font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font>
		<br><%= osheet.FItemList(i).FAlinkCode %>
	</td>
	<td ><font color="#777777"><%= Left(osheet.FItemList(i).FRegdate,10) %></font></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
	<% if osheet.FItemList(i).Ftargetid="10x10" then %>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsuplycash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalsuplycash,0) %></td>
	<% else %>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunbuycash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalbuycash,0) %></td>
	<% end if %>
	<td ><%= Left(osheet.FItemList(i).Fbeasongdate,10) %></td>
	<td rowspan=2>
		<%= FormatNumber(osheet.FItemList(i).f10totalsellcash,0) %> (10)
		<br><%= FormatNumber(osheet.FItemList(i).f90totalsellcash,0) %> (90)
		<br><%= FormatNumber(osheet.FItemList(i).f70totalsellcash,0) %> (70)
	</td>
	<td rowspan=2>
		<%= FormatNumber(osheet.FItemList(i).f10totalsuplycash,0) %> (10)
		<br><%= FormatNumber(osheet.FItemList(i).f90totalsuplycash,0) %> (90)
		<br><%= FormatNumber(osheet.FItemList(i).f70totalsuplycash,0) %> (70)
	</td>	
	<td rowspan=2>
		<input type="button" onclick="pop_detail('<%= osheet.FItemList(i).fidx %>');" class="button" value="상세">
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>
	    <% if IsNULL(osheet.FItemList(i).FIpgodate) then %>
	    	<%= Left(osheet.FItemList(i).Fscheduledate,10) %>
	    <% else %>
	    	<%= Left(osheet.FItemList(i).FIpgodate,10) %>
	    <% end if %>
	</td>
    <td>    	

    </td>
	<td colspan=3><font color="#777777"><%= DdotFormat(osheet.FItemList(i).Fbrandlist,30) %></font></td>
	
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=6>총계</td>
	<td align="right"><%= formatNumber(totaljumunsellcash,0) %></td>
	<td align="right"><%= formatNumber(totaljumunsuply,0) %></td>
	<td align="right"><%= formatNumber(totalfixsuply,0) %></td>
	<td align="center">
		<%= FormatNumber(tot10totalsellcash,0) %> (10)
		<br><%= FormatNumber(tot90totalsellcash,0) %> (90)
		<br><%= FormatNumber(tot70totalsellcash,0) %> (70)	
	</td>
	<td align="center">
		<%= FormatNumber(tot10totalsuplycash,0) %> (10)
		<br><%= FormatNumber(tot90totalsuplycash,0) %> (90)
		<br><%= FormatNumber(tot70totalsuplycash,0) %> (70)	
	</td>
	<td align="center">
		<input type="button" onclick="pop_detail('<%= arridx %>');" class="button" value="종합상세">
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=15 align=center>[ 검색결과가 없습니다. ]</td>
</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if osheet.HasPreScroll then %>
			<a href="?page=<%= osheet.StartScrollPage-1 %>&<%=parameter%>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + osheet.StartScrollPage to osheet.FScrollCount + osheet.StartScrollPage - 1 %>
			<% if i>osheet.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&<%=parameter%>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if osheet.HasNextScroll then %>
			<a href="?page=<%= i %>&<%=parameter%>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set osheet = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
