<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<!--

제한사항 :

 - 브라우저에서 window.open() 함수를 이용해 열려야 한다.

 window.open("pop_item_search_one.asp");

 - 부모 창은 다음의 함수를 제공해야 한다.

function ReActItemOne(itemid, itemoption);

-->

<%
dim page, mode, makerid, shopid, itemgubun, itemid, research, itemname, saleprice
dim nubeasong, nuitem, nuitemoption
dim onoffgubun, idx
dim onlineonly
dim pitemid, pitemoption, opentype                        ''' eastone 추가. 'R' :: 교환쪽에서 사용
dim sellyn

shopid = requestCheckvar(request("shopid"),32)
page = requestCheckvar(request("page"),10)
mode = requestCheckvar(request("mode"),32)
makerid = requestCheckvar(request("makerid"),32)
saleprice = requestCheckvar(request("saleprice"),10)

itemgubun = requestCheckvar(request("itemgubun"),10)
itemid = requestCheckvar(request("itemid"),10)
itemname = requestCheckvar(request("itemname"),32)
research = requestCheckvar(request("research"),10)
nubeasong = requestCheckvar(request("nubeasong"),10)
nuitem = requestCheckvar(request("nuitem"),10)
nuitemoption = requestCheckvar(request("nuitemoption"),10)
onoffgubun = requestCheckvar(request("onoffgubun"),10)
onlineonly = requestCheckvar(request("onlineonly"),10)
sellyn = requestCheckvar(request("sellyn"),10)

idx = requestCheckvar(request("idx"),10)
opentype   = requestCheckvar(request("opentype"),10)
pitemid     = requestCheckvar(request("pitemid"),10)
pitemoption = requestCheckvar(request("pitemoption"),10)
if (research<>"on") and (nuitem="") then
	nuitem = "on"
end if

if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if

if (onoffgubun="online") then
	itemgubun = "10"
end if

if (onoffgubun="online") then
	itemgubun = "10"
end if

if (saleprice <> "") then
	if (Not IsNumeric(saleprice)) then
		response.write "<script>alert('금액은 숫자만 가능합니다.'); history.back();</script>"
		response.end
	end if
end if

if (itemid <> "") then
	if (Not IsNumeric(itemid)) then
		response.write "<script>alert('상품코드는 숫자만 가능합니다.'); history.back();</script>"
		response.end
	end if
end if

if page="" then page=1



'==============================================================================
dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 50
ioffitem.FCurrPage = page
ioffitem.FRectDesigner = makerid
ioffitem.FRectNoSearchUpcheBeasong = nubeasong
ioffitem.FRectNoSearchNotusingItem = nuitem
ioffitem.FRectSellYN = sellyn

ioffitem.FRectItemgubun = itemgubun
ioffitem.FRectNoSearchNotusingItemOption = nuitemoption
ioffitem.FRectItemid = itemid
ioffitem.FRectItemName = itemname
ioffitem.FRectPriceRow = saleprice

if onoffgubun="offline" then
	ioffitem.GetOffShopItemList
else
	if (makerid = "") and (itemid = "") and (itemname = "") then

	else
		ioffitem.GetOnLineJumunByBrand
	end if
end if

dim i, shopsuplycash, buycash
%>
<script language='javascript'>

function NextPage(page) {
	frm.page.value=page;
	frm.submit();
}

function search(frm){
	frm.submit();
}

function SelectThisItem(itemid, itemoption){
	var frm = document.frm;

	opener.ReActItemOne(itemid, itemoption);

	opener.focus();
	window.close();
}

function SelectThisItem2(itemid, itemoption, itemname, itemoptionname, makerid, itemimg){
    var frm = document.frm;

	opener.ReActItemOne('<%= pitemid %>','<%= pitemoption %>',itemid, itemoption, itemname, itemoptionname, makerid, itemimg);

	opener.focus();
	window.close();
}
</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="shopid" value="<%= shopid %>" >
	<input type="hidden" name="opentype" value="<%= opentype %>" >
	<input type="hidden" name="pitemid" value="<%= pitemid %>" >
	<input type="hidden" name="pitemoption" value="<%= pitemoption %>" >
	<input type="hidden" name="nubeasong" value="<%= nubeasong %>" >

	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드:<% drawSelectBoxDesignerwithName "makerid", makerid %>
			&nbsp;
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=6 maxlength=7 onKeyPress="if (event.keyCode == 13) { document.frm.submit(); }">
			&nbsp;
			판매가 : <input type="text" class="text" name="saleprice" value="<%= saleprice %>" size=6 maxlength=7 onKeyPress="if (event.keyCode == 13) { document.frm.submit(); }">
			&nbsp;
			상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size=30 maxlength=64 onKeyPress="if (event.keyCode == 13) { document.frm.submit(); }">
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<!--
			<input type=checkbox name="nubeasong" <% if nubeasong="on" then response.write "checked" %> readonly>업체배송검색안함
			-->
			<input type=checkbox name="nuitem" <% if nuitem="on" then response.write "checked" %> >사용상품만
			<input type=checkbox name="nuitemoption" <% if nuitemoption="on" then response.write "checked" %> >사용옵션만
			<input type=checkbox name="sellyn" value="Y" <% if sellyn="Y" then response.write "checked" %> > 판매중지(일시품절 포함)상품 제외
		</td>
	</tr>
	</form>
</table>

<p>

<% if nubeasong="on" then %>
* 텐배 다른브랜드 선택가능, 업체배송상품 검색안함
<% end if %>

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ioffitem.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="9">
			검색결과 : <b><%= ioffitem.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= Page %> / <%= ioffitem.FTotalPage %></b>
		</td>
	</tr>
	<% end if %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="55">이미지</td>
		<td width="120">브랜드ID</td>
		<td width="100">상품코드</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="90">판매가<br>(할인가)</td>
		<td width="90">매입가</td>
		<td width="45">선택</td>
		<td width="50">비고</td>
	</tr>
	<% for i=0 to ioffitem.FResultCount -1 %>

	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">

	<tr bgcolor="#FFFFFF">
		<td align="center"><img src="<%= ioffitem.FItemList(i).FimageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
		<td ><%= ioffitem.FItemList(i).FMakerid %></td>
		<td align="center"><%= ioffitem.FItemList(i).GetBarCodeBoldStr %></td>
		<td ><%= ioffitem.FItemList(i).FShopItemName %></td>
		<td ><%= ioffitem.FItemList(i).FShopItemOptionName %></td>

		<td align=right style="padding-right:10px">
			<%= FormatNumber(ioffitem.FItemList(i).FShopItemOrgprice + ioffitem.FItemList(i).Foptaddprice,0) %>
			<% if (ioffitem.FItemList(i).FShopItemOrgprice > ioffitem.FItemList(i).Fshopitemprice) then %>
				<br><font color=red>(할)<%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice + ioffitem.FItemList(i).Foptaddprice,0) %></font>
			<% end if %>

			<% if ioffitem.FItemList(i).FitemCouponYn="Y" then %>
				<% if (CStr(ioffitem.FItemList(i).FitemCouponType) = "1") then %>
					<br><font color=green>(쿠)<%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice - (ioffitem.FItemList(i).Fshopitemprice * ioffitem.FItemList(i).FItemCouponValue / 100) + ioffitem.FItemList(i).Foptaddprice,0) %></font>
				<% elseif (CStr(ioffitem.FItemList(i).FitemCouponType) = "2") then %>
					<br><font color=green>(쿠)<%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice - ioffitem.FItemList(i).FItemCouponValue + ioffitem.FItemList(i).Foptaddprice,0) %></font>
				<% end if %>
			<% end if %>
		</td>
		<td align=right style="padding-right:10px">
			<%= FormatNumber(ioffitem.FItemList(i).Fshopsuplycash + ioffitem.FItemList(i).Foptaddbuyprice,0) %>
			<% if ioffitem.FItemList(i).FitemCouponYn="Y" then %>
				<% if ioffitem.FItemList(i).Fcouponbuyprice <> 0 then %>
					<br><font color=green><%= FormatNumber(ioffitem.FItemList(i).Fcouponbuyprice + ioffitem.FItemList(i).Foptaddbuyprice,0) %></font>
				<% end if %>
			<% end if %>
		</td>

		<td align="center">
		    <% if (opentype="R") then %>
		    <input type="button" class="button" value="선택" onclick="SelectThisItem2(<%= ioffitem.FItemList(i).Fshopitemid %>, '<%= ioffitem.FItemList(i).Fitemoption %>','<%= ioffitem.FItemList(i).FShopItemName %>','<%= ioffitem.FItemList(i).FShopitemoptionName %>','<%= ioffitem.FItemList(i).FMakerid %>','<%= ioffitem.FItemList(i).FimageSmall %>')">
		    <% else %>
			<input type="button" class="button" value="선택" onclick="SelectThisItem(<%= ioffitem.FItemList(i).Fshopitemid %>, '<%= ioffitem.FItemList(i).Fitemoption %>')">
			<% end if %>
		</td>
		<td>
			<% if ioffitem.FItemList(i).Foptusing="N" then %>
			<font color="red">옵션x</font><br>
			<% end if %>
			<% if ioffitem.FItemList(i).IsSoldOut then %>
			<font color="red">판매중지</font><br>
			<% end if %>
			<% if ioffitem.FItemList(i).IsTempSoldOut then %>
			<font color="red">일시품절</font><br>
			<% end if %>
			<% if ioffitem.FItemList(i).Flimityn="Y" then %>
			<font color="blue">한정(<%= ioffitem.FItemList(i).getLimitNo %>)</font><br>
			<% end if %>
		</td>
	</tr>
	</form>
	<% next %>
</table>



<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if ioffitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ioffitem.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
				<% if i>ioffitem.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ioffitem.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
