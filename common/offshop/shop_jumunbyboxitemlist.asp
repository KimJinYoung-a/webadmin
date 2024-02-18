<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 샵별패킹내역(상품별) 공용페이지
' History : 2011.01.18 이상구 생성
'			2012.08.14 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_balju.asp"-->
<%

if (Not C_IS_SHOP) and (Not C_ADMIN_USER) then
    %><html>
    <script type='text/javascript'>
    alert("세션이 종료되었습니다. \n재로그인후 사용하실수 있습니다.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
end if

dim page, shopid, chulgoyn, showdeleted, showmichulgo, michulgoreason
dim statecd, itemid, brandid, shopdiv, baljucode , boxno ,research, i
dim day5chulgo, shortchulgo, tempshort, danjong, etcshort
dim yyyy1,mm1 , dd1, yyyy2, mm2, dd2, fromDate, toDate
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	chulgoyn = requestCheckVar(request("chulgoyn"),1)
	showdeleted = requestCheckVar(request("showdel"),1)		'웹서버 웹나이트가 파라미터중 delete 문구가 있는 경우 막는다.
	showmichulgo = requestCheckVar(request("showmichulgo"),10)
	michulgoreason = requestCheckVar(request("michulgoreason"),32)
	boxno = requestCheckVar(request("boxno"),10)
	statecd = requestCheckVar(request("statecd"),10)
	itemid = requestCheckVar(request("itemid"),10)
	brandid = requestCheckVar(request("brandid"),32)
	shopdiv = requestCheckVar(request("shopdiv"),32)
	baljucode = requestCheckVar(request("baljucode"),32)
	day5chulgo = requestCheckVar(request("day5chulgo"),1)
	shortchulgo = requestCheckVar(request("shortchulgo"),1)
	tempshort = requestCheckVar(request("tempshort"),1)
	danjong = requestCheckVar(request("danjong"),1)
	etcshort = requestCheckVar(request("etcshort"),1)
	research = requestCheckVar(request("research"),2)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)

if (page = "") then
	page = 1
end if

if (research = "") then
	showdeleted = "N"
	michulgoreason = "all"
end if

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

michulgoreason = "|"
if (day5chulgo = "Y") then
	'5일내출고
	michulgoreason = michulgoreason + "5|"
end if
if (shortchulgo = "Y") then
	'재고부족
	michulgoreason = michulgoreason + "S|"
end if
if (tempshort = "Y") then
	'일시품절
	michulgoreason = michulgoreason + "T|"
end if
if (danjong = "Y") then
	'단종
	michulgoreason = michulgoreason + "D|"
end if
if (etcshort = "Y") then
	'기타
	michulgoreason = michulgoreason + "E|"
end if

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))
	if (C_IS_SHOP) then
		'샵은 한달
	    mm1 = Cstr(Month(now()) - 1)
	else
		mm1 = Cstr(Month(now()))
	end if
	dd1 = Cstr(day(now()))

	fromDate = DateSerial(yyyy1, mm1, dd1)
	yyyy1 = CStr(Year(fromDate))
	mm1 = CStr(Month(fromDate))
	dd1 = CStr(Day(fromDate))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oshopbalju
set oshopbalju = new CShopBalju
	oshopbalju.FRectFromDate = fromDate
	oshopbalju.FRectToDate = toDate
	oshopbalju.FRectBaljuId = shopid
	oshopbalju.FRectItemid = itemid
	oshopbalju.FRectBrandid = brandid
	oshopbalju.FRectShopdiv = shopdiv
	oshopbalju.FRectBaljucode = baljucode
	oshopbalju.FRectBoxno = boxno

	if (statecd = "A") then
		oshopbalju.FRectChulgoYN = "N"
	else
		oshopbalju.FRectStatecd = statecd
		if (C_IS_SHOP) then
			oshopbalju.FRectStatecd = ""
		    oshopbalju.FRectChulgoYN = "Y"
		end if
	end if

	oshopbalju.FRectShowDeleted = "N"
	oshopbalju.FRectMichulgoReason = michulgoreason
	oshopbalju.FCurrPage = page
	oshopbalju.Fpagesize = 25
	oshopbalju.GetShopBaljuByItem

%>

<script type='text/javascript'>

function GotoPage(pageno) {
	frm.page.value = pageno;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>"><%=CTX_SEARCH%><br><%= CTX_conditional %></td>
	<td align="left">
		ShopID :
		<% if (C_IS_SHOP) then %>
			<%= shopid %>
		<% else %>
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<% end if %>
		&nbsp;
		<%= CTX_Real_Order_Date %> :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<%= CTX_Order_Status %> :
		<% if (C_IS_SHOP) then %>
			<%= CTX_Shipped %>&nbsp;<%= CTX_after %>&nbsp;ALL
		<% else %>
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
				<option value="">========
				<option value="A" <% if statecd="A" then response.write "selected" %> ><%= CTX_Shipped %>&nbsp;<%= CTX_before %>&nbsp;ALL
				<option value="C" <% if statecd="C" then response.write "selected" %> ><%= CTX_Shipped %>&nbsp;<%= CTX_after %>&nbsp;ALL
			</select>
		<% end if %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="<%=CTX_SEARCH%>" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<%= CTX_Order_code %> : <input type="text" class="text" name="baljucode" value="<%= baljucode %>" size="10" maxlength="8">
		&nbsp;
		<%= CTX_Brand %> : <% drawSelectBoxDesignerwithName "brandid", brandid %>
		&nbsp;
		<%= CTX_Item_Code %> : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="12">
		&nbsp;
		<%= CTX_INNERBOX_NO %> : <input type="text" class="text" name="boxno" value="<%= boxno %>" size="4" maxlength="12">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
     	<%= ctx_shop %>&nbsp;<%= CTX_divide %> :
     	<input type="radio" name="shopdiv" value="" <% if shopdiv="" then response.write "checked" %> >ALL

     	<% if (Not C_IS_SHOP) then %>
			<input type="radio" name="shopdiv" value="direct" <% if shopdiv="direct" then response.write "checked" %> ><%= CTX_direct_store %>
			<input type="radio" name="shopdiv" value="franchisee" <% if shopdiv="franchisee" then response.write "checked" %> ><%= CTX_franchise %>
			<input type="radio" name="shopdiv" value="foreign" <% if shopdiv="foreign" then response.write "checked" %> ><%= CTX_Foreign_store %>
			<input type="radio" name="shopdiv" value="buy" <% if shopdiv="buy" then response.write "checked" %> ><%= CTX_wholesale %>
		<% end if %>
		&nbsp;&nbsp;
		|
		&nbsp;&nbsp;
		<%= CTX_delay %>&nbsp;<%=CTX_cause%> :
		<input type="checkbox" name="day5chulgo" value="Y" <% if day5chulgo="Y" then response.write "checked" %> >5<%= CTX_days %>&nbsp;<%= CTX_release %>
		<input type="checkbox" name="shortchulgo" value="Y" <% if shortchulgo="Y" then response.write "checked" %> ><%= CTX_stock %>&nbsp;<%= CTX_lack %>
		<input type="checkbox" name="tempshort" value="Y" <% if tempshort="Y" then response.write "checked" %> ><%= CTX_temporary %>&nbsp;<%= CTX_sold_out %>
		<input type="checkbox" name="danjong" value="Y" <% if danjong="Y" then response.write "checked" %> ><%= CTX_Discontinued %>
		<input type="checkbox" name="etcshort" value="Y" <% if etcshort="Y" then response.write "checked" %> ><%= CTX_etc %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<%= CTX_search_result %> : <b><%= oshopbalju.FTotalCount %></b>
		&nbsp;
		<%= CTX_page %> : <b><%= page %> / <%= oshopbalju.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><%= CTX_Real_Order_Date %></td>
	<td><%= CTX_INNERBOX_NO %></td>
	<td><%= CTX_real_Order_code %></td>
	<td><%= CTX_Order_code %></td>
	<td><%= CTX_release %>&nbsp;<%=CTX_Status%></td>
	<td><%= CTX_Brand %></td>
	<td width=55><%= CTX_Image %></td>
	<td><%= CTX_Item_Code %></td>
	<td><%= CTX_Description %><br><font color="blue">[<%= CTX_Description_Option %>]</font></td>
	<td><%= CTX_request %><br><%= CTX_quantity %></td>
	<td><%= CTX_release %><br><%= CTX_quantity %></td>
	<td><%= CTX_Note %></td>
</tr>
<% if oshopbalju.FResultCount >0 then %>
<% for i=0 to oshopbalju.FResultcount-1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= oshopbalju.FItemList(i).Fbaljudate %></td>
	<td align="center">
		<%
		if (oshopbalju.FItemList(i).Fboxno <> "0") then
			response.write oshopbalju.FItemList(i).Fboxno
		end if
		%>
	</td>
	<td align="center"><%= oshopbalju.FItemList(i).Fbaljunum %></td>
	<td align="center"><%= oshopbalju.FItemList(i).Fbaljucode %></td>
	<td align="center">
		<font color="<%= oshopbalju.FItemList(i).GetStateColor %>"><%= oshopbalju.FItemList(i).GetStateName %></font>
		<% if (oshopbalju.FItemList(i).Frealitemno > 0) then %>
			<br><%= oshopbalju.FItemList(i).FAlinkCode %>
		<% end if %>
	</td>
	<td align="center"><%= oshopbalju.FItemList(i).Fmakerid %></td>
	<td align="center"><img src="<%= oshopbalju.FItemList(i).Fmainimageurl %>" width="50"></td>
	<td align="center">
		<%= oshopbalju.FItemList(i).Fitemgubun %><%= CHKIIF(oshopbalju.FItemList(i).Fitemid>=1000000,Format00(8,oshopbalju.FItemList(i).Fitemid),Format00(6,oshopbalju.FItemList(i).Fitemid)) %><%= oshopbalju.FItemList(i).Fitemoption %>
	</td>
	<td align="left">
		<%= oshopbalju.FItemList(i).Fitemname %>
		<% if (oshopbalju.FItemList(i).Fitemoption <> "0000") then %>
			<br><font color="blue">[<%= oshopbalju.FItemList(i).Fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td align="center"><%= oshopbalju.FItemList(i).Fbaljuitemno %></td>
	<td align="center"><% if (oshopbalju.FItemList(i).Fbaljuitemno <> oshopbalju.FItemList(i).Frealitemno) then %><font color=red><b><% end if %><%= oshopbalju.FItemList(i).Frealitemno %></td>
	<td align="center">
		<%= oshopbalju.FItemList(i).Fcomment %>
		<%= oshopbalju.FItemList(i).Fipgoflag %>
	</td>
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=15 align=center>[<%= CTX_search_returns_no_results %>]</td>
</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oshopbalju.HasPreScroll then %>
			<a href="javascript:GotoPage(<%= oshopbalju.StartScrollPage-1 %>)">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oshopbalju.StartScrollPage to oshopbalju.FScrollCount + oshopbalju.StartScrollPage - 1 %>
			<% if i>oshopbalju.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:GotoPage(<%= i %>)">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oshopbalju.HasNextScroll then %>
			<a href="javascript:GotoPage(<%= i %>)">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set oshopbalju = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
