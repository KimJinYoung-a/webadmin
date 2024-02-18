<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/stock/upcheorderitemcls.asp"-->

<%
const C_STOCK_DAY=7

dim page, mode, suplyer,shopid,itemid, research
dim nubeasong, nuitem, nuitemoption, nudanjong, soldoutover7days
dim onoffgubun, idx, ShortageType

shopid = request("shopid")
page = request("page")
mode = request("mode")
suplyer = request("suplyer")
itemid = request("itemid")
research = request("research")
nubeasong = request("nubeasong")
nuitem = request("nuitem")
nuitemoption = request("nuitemoption")
nudanjong = request("nudanjong")
soldoutover7days = request("soldoutover7days")
onoffgubun = request("onoffgubun")
idx = request("idx")
ShortageType = request("ShortageType")

if (research<>"on") and (nubeasong="") then
	nubeasong = "on"
end if

if (research<>"on") and (nuitem="") then
	nuitem = "on"
end if

if (research<>"on") and (nuitemoption="") then
	nuitemoption = "on"
end if

if (research<>"on") and (nudanjong="") then
	nudanjong = "on"
end if

if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if

if page="" then page=1
if mode="" then mode="bybrand"

'상품코드 유효성 검사(2008.07.31;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim ojumunitem
set ojumunitem  = new CUpcheOrderItem
ojumunitem.FPageSize = 50
ojumunitem.FCurrPage = page
ojumunitem.FRectDesigner = suplyer
ojumunitem.FRectNoSearchUpcheBeasong = nubeasong
ojumunitem.FRectNoSearchNotusingItem = nuitem
ojumunitem.FRectNoSearchNotusingItemOption = nuitemoption
ojumunitem.FRectNoSearchDanjong = nudanjong
ojumunitem.FRectNoSearchSoldoutover7days = soldoutover7days
ojumunitem.FRectItemid = itemid
ojumunitem.FRectShortage7days = chkIIF(ShortageType="7day","on","")
ojumunitem.FRectShortage14days = chkIIF(ShortageType="14day","on","")
ojumunitem.FRectShortageRealStock = chkIIF(ShortageType="5under","on","")

if onoffgubun="offline" then
	ojumunitem.GetOffShopItemList
else
	if (suplyer<>"") or (itemid<>"") then
		ojumunitem.GetOnLineJumunByBrand
	end if
end if

dim i, shopsuplycash, buycash
%>
<script language='javascript'>
function popOffItemEdit(ibarcode){
	var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}


function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'adminitemselledit','width=500,height=600,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function enablebrand(bool){
	//document.frm.designer.disabled = bool;
}

function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function search(frm){
	/*
	if ((frm.suplyer.value.length<1)){
		if ((frm.mode[0].checked)&&(frm.designer.value.length<1)){
			alert('브랜드를 선택 하세요.');
			frm.designer.focus();
			return;
		}
	}
	*/

	frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function AddArr(){
	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.itemnamearr.value = "";
	upfrm.itemoptionnamearr.value = "";
	upfrm.designerarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsInteger(frm.itemno.value)){
					alert('갯수는 정수만 가능합니다.');
					frm.itemno.focus();
					return;
				}

				if (frm.itemno.value=="0"){
					alert('수량을 입력하세요.');
					frm.itemno.focus();
					return;
				}

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
				upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
				upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
				upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
				upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
				upfrm.mwdivarr.value = upfrm.mwdivarr.value + frm.mwdiv.value + "|";

			}
		}
	}


	opener.ReActItems('<%= idx %>', upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value,upfrm.mwdivarr.value);


	//초기화
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				frm.cksel.checked = false;
				frm.itemno.value="0"


			}
		}
	}

}
</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="page" value="<%= page %>">
	<% if (request("changesuplyer") <> "Y") then %>
	<input type="hidden" name="suplyer" value="<%= suplyer %>" >
	<% else %>
	<input type="hidden" name="changesuplyer" value="Y" >
	<% end if %>
	<input type="hidden" name="shopid" value="<%= shopid %>" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<% if (request("changesuplyer") <> "Y") then %>
			브랜드 : <b><%= suplyer %></b>
			<% else %>
			브랜드 : <% drawSelectBoxDesignerwithName "suplyer", suplyer %>
			<% end if %>
			&nbsp;
			<input type=checkbox name="nubeasong" <% if nubeasong="on" then response.write "checked" %> >업체배송제외
			<input type=checkbox name="nuitem" <% if nuitem="on" then response.write "checked" %> >사용상품만
			<input type=checkbox name="nuitemoption" <% if nuitemoption="on" then response.write "checked" %> >사용옵션만
			<input type=checkbox name="nudanjong" <% if nudanjong="on" then response.write "checked" %> >단종제외
			<input type=checkbox name="soldoutover7days" <% if soldoutover7days="on" then response.write "checked" %> >재고부족제외



			<br>
			구분 : <input type="radio" name="onoffgubun" value="online" <% if onoffgubun="online" then response.write "checked" %> >온라인
			<input type="radio" name="onoffgubun" value="offline" <% if onoffgubun="offline" then response.write "checked" %> >오프라인
			&nbsp;&nbsp;
			상품코드로검색 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=6 maxlength=7>
            &nbsp;&nbsp;
            부족구분:
            <input type="radio" name="ShortageType" value="" <% if ShortageType="" then response.write "checked" %> >전체
            <input type="radio" name="ShortageType" value="7day" <% if ShortageType="7day" then response.write "checked" %> ><%= C_STOCK_DAY %>일후재고부족
			<input type="radio" name="ShortageType" value="14day" <% if ShortageType="14day" then response.write "checked" %> ><%= C_STOCK_DAY*2 %>일후재고부족
            <input type="radio" name="ShortageType" value="5under" <% if ShortageType="5under" then response.write "checked" %> >실사재고 5이하
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:search(frm);">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<!-- 상단바 시작 -->
	<% if ojumunitem.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						검색결과 : <b><%= ojumunitem.FTotalCount %></b>
						&nbsp;
						페이지 : <b><%= Page %> / <%= ojumunitem.FTotalPage %></b>
					</td>
					<td align="right">
						<input type="button" class="button" value="전체선택" onClick="AnSelectAllFrame(true)">
        				<input type="button" class="button" value="선택 아이템 추가" onclick="AddArr()">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<% end if %>

	<!-- 상단바 끝 -->
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
		<td width="50">이미지</td>
		<td width="80">브랜드ID</td>
		<td width="90">상품코드</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="45">판매가</td>
		<td width="45">매입가</td>
		<td width="45">마진</td>
		<td width="45">수량</td>
		<td>비고</td>
	</tr>
	<% for i=0 to ojumunitem.FResultCount -1 %>

	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ojumunitem.FItemList(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ojumunitem.FItemList(i).Fitemid %>">
	<input type="hidden" name="itemoption" value="<%= ojumunitem.FItemList(i).Fitemoption %>">
	<input type="hidden" name="itemname" value="<%= ojumunitem.FItemList(i).FItemName %>">
	<input type="hidden" name="itemoptionname" value="<%= ojumunitem.FItemList(i).FItemOptionName %>">
	<input type="hidden" name="desingerid" value="<%= ojumunitem.FItemList(i).FMakerid %>">
	<input type="hidden" name="sellcash" value="<%= ojumunitem.FItemList(i).Fsellcash %>">
	<input type="hidden" name="suplycash" value="<%= ojumunitem.FItemList(i).FBuycash %>">
	<input type="hidden" name="buycash" value="<%= ojumunitem.FItemList(i).FBuycash %>">
	<input type="hidden" name="mwdiv" value="<%= ojumunitem.FItemList(i).Fmwdiv %>">

	<% if (ojumunitem.FItemList(i).Foptusing="N") or (ojumunitem.FItemList(i).Fisusing="N") then %>
	<tr bgcolor="<%= adminColor("gray") %>">
	<% else %>
	<tr bgcolor="#FFFFFF">
	<% end if %>
		<td rowspan=2><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td rowspan=2><img src="<%= ojumunitem.FItemList(i).FimageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
		<td ><%= ojumunitem.FItemList(i).FMakerid %></td>
		<% if ojumunitem.FItemList(i).FItemGubun<>"10" then %>
		<td ><a href="javascript:popOffItemEdit('<%= ojumunitem.FItemList(i).GetBarCode %>')"><%= ojumunitem.FItemList(i).GetBarCodeBoldStr %></a></td>
		<% else %>
		<td ><a href="javascript:PopItemSellEdit('<%= ojumunitem.FItemList(i).FItemID %>');"><%= ojumunitem.FItemList(i).GetBarCodeBoldStr %></a></td>
		<% end if %>
		<td ><a href="/admin/stock/itemcurrentstock.asp?itemid=<%= ojumunitem.FItemList(i).FItemID %>&itemoption=<%= ojumunitem.FItemList(i).FItemOption %>" target=_blank ><%= ojumunitem.FItemList(i).FItemName %></a></td>
		<td ><%= ojumunitem.FItemList(i).FItemOptionName %></td>
		<td rowspan=2 align=right><%= FormatNumber(ojumunitem.FItemList(i).FSellcash,0) %></td>
		<td rowspan=2 align=right><%= FormatNumber(ojumunitem.FItemList(i).FBuycash,0) %></td>
		<td rowspan=2 align=center>
		<font color="<%= ojumunitem.FItemList(i).getMwDivColor %>"><%= ojumunitem.FItemList(i).getMwDivName %></font><br>
		<% if ojumunitem.FItemList(i).FSellcash<>0 then %>
		<%= 100-(CLng(ojumunitem.FItemList(i).FBuycash/ojumunitem.FItemList(i).FSellcash*10000)/100) %> %
		<% end if %>
		</td>
		<td rowspan=2>
			<% if ojumunitem.FItemList(i).Frealstock<0 and ojumunitem.FItemList(i).Fsell7days=0 then %>
			<input type="text" class="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
		    <% elseif ojumunitem.FItemList(i).GetNdayShortageNo(14) < 0 then %>
		    <input type="text" class="text" name="itemno" value="<%= (ojumunitem.FItemList(i).GetNdayShortageNo(14))*-1 %>" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
		    <% else %>
		    <input type="text" class="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
		    <% end if %>
		</td>
		<td rowspan=2 >
			<%= fnColor(ojumunitem.FItemList(i).Fdanjongyn,"dj") %>
			<br>
			<% if ojumunitem.FItemList(i).Foptusing="N" then %>
			<font color="red">옵션x</font><br>
			<% end if %>
			<% if ojumunitem.FItemList(i).IsSoldOut then %>
			<font color="red">판매중지</font><br>
			<% end if %>
			<% if ojumunitem.FItemList(i).Flimityn="Y" then %>
			<font color="blue">한정(<%= ojumunitem.FItemList(i).getOptionLimitNo %>)</font><br>
			<% end if %>
			<% if ojumunitem.FItemList(i).Fpreorderno<>0 then %>
				기주문:
				<% if ojumunitem.FItemList(i).Fpreorderno<>ojumunitem.FItemList(i).Fpreordernofix then response.write "</br>" + CStr(ojumunitem.FItemList(i).Fpreorderno) + "->" %>
					<%= ojumunitem.FItemList(i).Fpreordernofix %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=4>
			<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td>입고</td>
					<td>판매</td>
					<td>출고</td>
					<td>기타</td>
					<td>CS</td>
					<td>불량</td>
					<td>오차</td>

					<td>실사재고</td>
					<td bgcolor="<%= adminColor("green") %>">출고이전</td>
					<td>예상재고</td>

					<% if ojumunitem.FItemList(i).Fmaxsellday<>7 then %>
					<td bgcolor="<%= adminColor("green") %>">On<font color="#CC1111"><%= ojumunitem.FItemList(i).Fmaxsellday %></font>일</td>
					<td bgcolor="<%= adminColor("green") %>">Off<font color="#CC1111"><%= ojumunitem.FItemList(i).Fmaxsellday %></font>일</td>
					<% else %>
					<td bgcolor="<%= adminColor("green") %>">On<%= ojumunitem.FItemList(i).Fmaxsellday %>일</td>
					<td bgcolor="<%= adminColor("green") %>">Off<%= ojumunitem.FItemList(i).Fmaxsellday %>일</td>
					<% end if %>

					<td><%= C_STOCK_DAY %>일</td>
					<td><%= C_STOCK_DAY*2 %>일</td>
					<!--
					<td>OFF준비</td>
					-->
				</tr>
				<tr bgcolor="#FFFFFF" align=center>
					<td><%= ojumunitem.FItemList(i).Ftotipgono %></td>
					<td><%= ojumunitem.FItemList(i).Ftotsellno %></td>
					<td><%= ojumunitem.FItemList(i).Ftotchulgono %></td>
					<td></td>
					<td></td>
					<td><%= ojumunitem.FItemList(i).Ferrbaditemno %></td>
					<td><%= ojumunitem.FItemList(i).Ferrrealcheckno %></td>

					<td>
						<b>
						<% if ojumunitem.FItemList(i).Frealstock<1 then %>
						<font color="#CC1111"><b><%= ojumunitem.FItemList(i).GetCheckStockNo %></b></font>
						<% else %>
						<%= ojumunitem.FItemList(i).Frealstock %>
						<% end if %>
						</b>
					</td>

					<td>
					    <!-- 출고이전 -->
					    <%= ojumunitem.FItemList(i).GetReqNotChulgoNo %></td>
					</td>
					<td>
						<b>
						<% if ojumunitem.FItemList(i).Frealstock + ojumunitem.FItemList(i).GetReqNotChulgoNo < 1 then %>
						<font color="#CC1111"><%= ojumunitem.FItemList(i).Frealstock + ojumunitem.FItemList(i).GetReqNotChulgoNo %></b></font>
						<% else %>
						<%= ojumunitem.FItemList(i).Frealstock + ojumunitem.FItemList(i).GetReqNotChulgoNo %>
						<% end if %>
						</b>
					</td>
					<td><%= ojumunitem.FItemList(i).Fsell7days %></td>
					<td><%= ojumunitem.FItemList(i).Foffchulgo7days %></td>


					<td>
					    <!-- 7일 -->
						<% if ojumunitem.FItemList(i).Fshortageno< 1 then %>
						<font color="#CC1111"><b><%= ojumunitem.FItemList(i).Fshortageno %></b></font>
						<% else %>
						<%= ojumunitem.FItemList(i).Fshortageno %>
						<% end if %>
					</td>
					<td>
					    <!-- N일 필요 -->
						<% if (ojumunitem.FItemList(i).GetNdayShortageNo(14))< 1 then %>
						<font color="#CC1111"><b><%= ojumunitem.FItemList(i).GetNdayShortageNo(14) %></b></font>
						<% else %>
						<%= ojumunitem.FItemList(i).GetNdayShortageNo(14) %>
						<% end if %>
					</td>
					<!--
					<td><%= ojumunitem.FItemList(i).Foffconfirmno %></td>
					-->
				</tr>
			</table>
		</td>
	</tr>
	</form>
	<% next %>

	<!-- 하단바 시작 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if ojumunitem.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojumunitem.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ojumunitem.StarScrollPage to ojumunitem.FScrollCount + ojumunitem.StarScrollPage - 1 %>
				<% if i>ojumunitem.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ojumunitem.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
	<!-- 하단바 끝 -->
</table>



<form name="frmArrupdate" method="post" action="">
<input type="hidden" name="mode" value="arrins">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="itemnamearr" value="">
<input type="hidden" name="itemoptionnamearr" value="">
<input type="hidden" name="designerarr" value="">
<input type="hidden" name="mwdivarr" value="">
</form>
<script language='javascript'>
//alert('수정중');
</script>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->