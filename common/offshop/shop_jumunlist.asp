<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 주문관리(물류)
' Hieditor : 2009.04.07 서동석 생성
'			 2011.03.28 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<%
dim page, shopid, statecd, baljucode ,baljunum ,nowdate ,refer
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2,fromDate,toDate ,beforeidxclick
dim i ,totaljumunsuply, totalfixsuply, totaljumunsellcash ,makerid
	yyyy1 = RequestCheckVar(request("yyyy1"),4)
	yyyy2 = RequestCheckVar(request("yyyy2"),4)
	mm1	  = RequestCheckVar(request("mm1"),2)
	mm2	  = RequestCheckVar(request("mm2"),2)
	dd1	  = RequestCheckVar(request("dd1"),2)
	dd2	  = RequestCheckVar(request("dd2"),2)
	shopid      = RequestCheckVar(request("shopid"),32)
	statecd     = RequestCheckVar(request("statecd"),3)
	baljucode   = RequestCheckVar(request("baljucode"),16)
	baljunum    = RequestCheckVar(request("baljunum"),16)
	page    = RequestCheckVar(request("page"),10)
	beforeidxclick    = RequestCheckVar(request("beforeidxclick"),10)
	refer = request.ServerVariables("HTTP_REFERER")
	makerid      = RequestCheckVar(request("makerid"),32)

if page="" then page=1

if C_ADMIN_USER then

'/직영점
elseif C_IS_OWN_SHOP then
	shopid = C_STREETSHOPID		'"streetshop011"

'직영/가맹점
elseif (C_IS_SHOP) then

	'/어드민권한 점장 미만
	'if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID		'"streetshop011"
	'end if
else
	''업체인경우
	if (C_IS_Maker_Upche) then
	    makerid = session("ssBctId")
	else
		if not(C_ADMIN_USER) then
		else
		end if
	end if
end if

if (yyyy1="") then
    nowdate = Left(CStr(dateadd("m",-1,now())),10)
	yyyy1   = Left(nowdate,4)
	mm1     = Mid(nowdate,6,2)
	dd1     = Mid(nowdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2   = Left(nowdate,4)
	mm2     = Mid(nowdate,6,2)
	dd2     = Mid(nowdate,9,2)
end if

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim osheet
set osheet = new COrderSheet
	osheet.FCurrPage = page
	osheet.Fpagesize=20
	osheet.FRectBaljuId = shopid
	osheet.FRectStatecd = statecd
	osheet.FRectBaljuCode = baljucode
	osheet.FRectStartDate = fromDate
	osheet.FRectEndDate = toDate
	osheet.FRectbaljunum = baljunum
	osheet.frectmakerid = makerid

	'/본사와 업체일 경우에만 전체 리스트
	if C_ADMIN_USER or C_IS_Maker_Upche then
		osheet.GetOrderSheetList
	else
		if (shopid<>"") then
			osheet.GetOrderSheetList
		else
			response.write "<script type='text/javascript'>"
			response.write "	alert('"& CTX_Please_select & " (" & CTX_SHOP & ")"&"');"
			response.write "</script>"
		end if
	end if
%>

<script type='text/javascript'>

function PopIpgoSheet(v,itype){
	var popwin;
	popwin = window.open('popshopjumunsheet.asp?idx=' + v + '&itype=' + itype,'shopjumunsheet','width=680,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function ChangeJState(idx){
    if (confirm('주문접수로 변경한 이후에는 내역을 수정할 수 없습니다.\n\n주문 접수 상태로 변경 하시겠습니까?')){
        document.upfrm.masteridx.value = idx;
        document.upfrm.mode.value = "jupsuchange";
        document.upfrm.submit();
    }
}

function MakeJumun(){
	location.href='/common/offshop/shop_jumuninput.asp';
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>"><%=CTX_SEARCH%><br><%= CTX_conditional %></td>
	<td align="left">
		<%= CTX_Order_code %> : <input type="text" class="text" name="baljucode" value="<%= baljucode %>" size="10" maxlength="8">
		<%= CTX_real_Order_code %> : <input type="text" class="text" name="baljunum" value="<%= baljunum %>" size="10" maxlength="8">
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				<%= CTX_Brand %> :
				<% drawSelectBoxDesignerwithName "makerid", makerid %>
				<%' drawSelectBoxDesignerOffWitakContract "chargeid", chargeid, shopid, "'B012','B022','B023'", " ReSearch('');" %>
				&nbsp;<%=CTX_SHOP%> : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				<%= CTX_Brand %> :
				<% drawSelectBoxDesignerwithName "makerid", makerid %>
				<%' drawSelectBoxDesignerOffWitakContract "chargeid", chargeid, shopid, "'B012','B022','B023'", " ReSearch('');" %>
				&nbsp;<%= CTX_SHOP %> : <% drawSelectBoxOffShop "shopid",shopid %>
		<%
			end if
		else
			''업체인경우
			if (C_IS_Maker_Upche) then
		%>
				<%= CTX_Brand %> : <%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
				&nbsp;<%= CTX_SHOP %> : <% drawBoxDirectIpchulOffShopByMakerchfg "shopid", shopid, makerid ," NextPage('');","'B012','B022','B023'" %>
		<%
			else
				if (C_ADMIN_USER) then
		%>
					<%= CTX_Brand %> : <% drawSelectBoxDesignerwithName "makerid", makerid %>
					&nbsp;<%= CTX_SHOP %> : <% drawSelectBoxOffShop "shopid",shopid %>
		<%
				end if
			end if
		end if
		%>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="<%=CTX_SEARCH%>" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<%= CTX_Order_Date %> : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
     	&nbsp;<%= CTX_Order_Status %> :
		<select name="statecd" class="select">
			<option value="">ALL
			<option value=" " <% if statecd=" " then response.write "selected" %> ><%= CTX_in_process %>
			<option value="0" <% if statecd="0" then response.write "selected" %> ><%= CTX_Register %>
			<option value="1" <% if statecd="1" then response.write "selected" %> ><%= CTX_Confirmed %>
			<option value="2" <% if statecd="2" then response.write "selected" %> ><%= CTX_Payment_waiting %>
			<option value="5" <% if statecd="5" then response.write "selected" %> ><%= CTX_Packing_in_Process %>
			<option value="6" <% if statecd="6" then response.write "selected" %> ><%= CTX_Shipment_Standby %>
			<option value="7" <% if statecd="7" then response.write "selected" %> ><%= CTX_Shipped %>
			<option value="8" <% if statecd="8" then response.write "selected" %> ><%= CTX_preparing %>
			<option value="9" <% if statecd="9" then response.write "selected" %> ><%= CTX_stocked %>
		</select>
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
		<input type="button" class="button" value="<%= CTX_CREATE_ORDER %>" onClick="MakeJumun();">
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<%= CTX_search_result %> : <b><% = osheet.FTotalCount %></b>
		&nbsp;
		<%= CTX_page %> : <b><%= page %> / <%= osheet.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><%= CTX_Order_code %></td>
	<td><%= CTX_WHOLESALEID %></td>
	<td><%= CTX_an_orderer %><br>(SHOP)</td>
	<td><%= CTX_Order_Date %></td>
	<td><%= CTX_Real_Order_Date %><br><%= CTX_real_Order_code %></td>
	<td><%= CTX_Shipment_Date %></td>
	<td><%= CTX_tax_Bill_date %></td>
	<td><%=CTX_Deposit_Date%></td>
	<td><%= CTX_Status %></td>
	<td><%= CTX_total %><br>(<%= CTX_consumer_price %>)</td>
	<td><%= CTX_total %><br>(<%= CTX_Supply_price %>)</td>
	<td><%= CTX_Cost_confirmed %><br>(<%= CTX_Supply_price %>)</td>
	<!--<td>송장번호</td>	-->
	<td><%= CTX_DETAIL %></td>
	<td><%= CTX_Note %></td>
</tr>
<% if osheet.FResultCount >0 then %>
<% for i=0 to osheet.FResultcount-1 %>
<%
totaljumunsellcash = totaljumunsellcash + osheet.FItemList(i).Fjumunsellcash
totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunsuplycash
totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalsuplycash
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><a href="shop_jumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>&menupos=<%= menupos %>" onfocus="this.blur();"><%= osheet.FItemList(i).Fbaljucode %></a></td>
	<td><%= osheet.FItemList(i).Ftargetid %><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
	<td><%= osheet.FItemList(i).Fbaljuid %><br>(<%= osheet.FItemList(i).Fbaljuname %>)</td>
	<td><%= Left(osheet.FItemList(i).FRegdate,10) %></td>
	<td><%= Left(osheet.FItemList(i).Fbaljudate,10) %><br><%= osheet.FItemList(i).Fbaljunum %></td>
	<td>
		<% if IsNull(osheet.FItemList(i).Fbeasongdate) then %>
			<%= Left(osheet.FItemList(i).Fipgodate,10) %>
		<% else %>
			<%= Left(osheet.FItemList(i).Fbeasongdate,10) %>
		<% end if %>
	</td>
	<td><%= Left(osheet.FItemList(i).Fsegumdate,10) %></td>
	<td><%= Left(osheet.FItemList(i).Fipkumdate,10) %></td>
	<td>
		<font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font>
		<% if (osheet.FItemList(i).FStateCD=" ") then %>
		    <br><input type="button" value="<%= CTX_change_Status %>" class="button" onClick="ChangeJState('<%= osheet.FItemList(i).Fidx %>');">
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsuplycash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalsuplycash,0) %></td>
	<!--<td><%= Left(osheet.FItemList(i).Fsongjangno,10) %></td>	-->
	<td>
		<!--
		<a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>','2');">소</a>/<a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>','1');">공</a>
		-->
		<a href="javascript:ViewOfflineOrderSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexplorer.gif" width=21 border=0></a>
		<a href="javascript:ExcelOfflineOrderSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexcel.gif" width=21 border=0></a>
	</td>
	<td width=150>
		<input type="button" class="button" value="출력" onclick="printbarcode_off('JUMUN', '', '', '', '', '', '<%= osheet.FItemList(i).Fidx %>', '', '');">
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=8><%= CTX_total %></td>
	<td align="right"><%= formatNumber(totaljumunsellcash,0) %></td>
	<td align="right"><%= formatNumber(totaljumunsuply,0) %></td>
	<td align="right"><%= formatNumber(totalfixsuply,0) %></td>
	<td colspan="3"></td>
</tr>
<tr bgcolor="#FFFFFF" height=20 align="center">
	<td colspan="16">
		<% if osheet.HasPreScroll then %>
			<a href="javascript:NextPage('<%= osheet.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + osheet.StartScrollPage to osheet.FScrollCount + osheet.StartScrollPage - 1 %>
			<% if i>osheet.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if osheet.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=20>[ <%= CTX_search_returns_no_results %> ]</td>
</tr>
<% end if %>
<form name="upfrm" method="post" action="common_shopjumun_process.asp" >
	<input type="hidden" name="masteridx" value="">
	<input type="hidden" name="mode" value="jupsuchange">
</form>
</table>

<%
set osheet = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->