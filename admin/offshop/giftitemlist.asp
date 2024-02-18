<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 사은품 등록
' Hieditor : 2013.01.15 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer,page,usingyn , research, mageview, imageview, itemgubun, itemid, itemname
dim cdl, cdm, cds, i, PriceDiffExists , IsDirectIpchulContractExistsBrand ,publicbarcode
dim weightYn, sizeYn
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),9)
	usingyn     = RequestCheckVar(request("usingyn"),1)
	research    = RequestCheckVar(request("research"),9)
	imageview   = RequestCheckVar(request("imageview"),9)
	itemgubun   = RequestCheckVar(request("itemgubun"),16)
	itemid      = RequestCheckVar(request("itemid"),9)
	itemname    = RequestCheckVar(request("itemname"),32)
	publicbarcode    = RequestCheckVar(request("publicbarcode"),20)
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
    weightYn    = RequestCheckVar(request("weightYn"),3)
    sizeYn      = RequestCheckVar(request("sizeYn"),3)
	if page="" then page=1
	if research<>"on" then usingyn="Y"

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 100
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = designer
	ioffitem.FRectOnlyUsing = usingyn
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemID = itemid
	ioffitem.FRectItemName = html2db(itemname)
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.FRectpublicbarcode = publicbarcode

    ioffitem.FRectIsWeight		= weightYn
    ioffitem.FRectSizeYn = sizeYn

	ioffitem.GetOffNOnLineGiftItemList
%>

<script language='javascript'>

function NotUsingCheckAll(){
    var frm;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
		    if (frm.isusing[0].checked==true){
		        frm.isusing[1].checked = true;
		        frm.cksel.checked = true;
		        AnCheckClick(frm.cksel);
		    }
		}
	}
}

//수정
function pop_itemedit_gift_edit(ibarcode){
	var pop_itemedit_gift_edit = window.open('/admin/offshop/pop_itemedit_gift_edit.asp?barcode=' + ibarcode,'pop_itemedit_gift_edit','width=1400,height=800,resizable=yes,scrollbars=yes');
	pop_itemedit_gift_edit.focus();
}

//등록
function pop_itemedit_gift_new(){
	var pop_itemedit_gift_new;

	pop_itemedit_gift_new = window.open('/admin/offshop/pop_itemedit_gift_edit.asp','pop_itemedit_gift_new','width=1400,height=800,scrollbars=yes,resizable=yes');
	pop_itemedit_gift_new.focus();
}

function ReSearch(page){
	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('상품번호는 숫자만 가능합니다.');
			frm.itemid.focus();
			return;
		}
	}

	frm.page.value = page;
	frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function ChargeIdAvail(ichargeid){
	var comp = document.frm.designer;

	if (ichargeid=="10x10"){
		return true
	}

	for (var i=0;i<comp.length;i++){
		if (comp[i].value==ichargeid){
			return true
		}
	}

	return false;
}

function ModiArr(){
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
	upfrm.itempricearr.value = "";
	upfrm.itemsuplyarr.value = "";
	upfrm.onofflinkynarr.value = "";
	upfrm.extbarcodearr.value = "";

	upfrm.shopbuypricearr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (frm.tx_sellcash.value > 0) {
					alert("사은품은 판매가가 0이하여야 합니다.");
					frm.tx_sellcash.focus();
					return;
				}

				if (frm.tx_orgsellprice.value > 0) {
					alert("사은품은 소비자가가 0이하여야 합니다.");
					frm.tx_orgsellprice.focus();
					return;
				}

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.orgsellpricearr.value = upfrm.orgsellpricearr.value + frm.tx_orgsellprice.value + "|";
				upfrm.itempricearr.value = upfrm.itempricearr.value + frm.tx_sellcash.value + "|";

				upfrm.itemsuplyarr.value = upfrm.itemsuplyarr.value + frm.tx_suplycash.value + "|";
				upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.tx_shopbuyprice.value + "|";

				upfrm.extbarcodearr.value = upfrm.extbarcodearr.value + frm.extbarcode.value + "|";
				upfrm.onofflinkynarr.value = upfrm.onofflinkynarr.value + frm.onofflinkyn.value + "|";

				if (frm.isusing[0].checked){
					upfrm.isusingarr.value = upfrm.isusingarr.value + "Y" + "|";
				}else{
					upfrm.isusingarr.value = upfrm.isusingarr.value + "N" + "|";
				}
			}
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		upfrm.mode.value = "arrmodi";
		upfrm.submit();
	}
}

function samePriceALL(){
    var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
		    samePrice(frm);
		}
	}
}

function samePrice(frm){
    frm.tx_orgsellprice.value=frm.oldonlineorgprice.value*1 + frm.oldonlineOptAddprice.value*1;  //소비자가
	frm.tx_sellcash.value=frm.oldonlineprice.value*1 + frm.oldonlineOptAddprice.value*1;         //판매가

	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function EventPrice0(){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (frm.tx_discountsellprice.value!=0){
				frm.tx_discountsellprice.value=0;
				frm.cksel.checked=true;
				AnCheckClick(frm.cksel);
			}
		}
	}
}

function BuyPrice0(){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
    			if (frm.tx_suplycash.value!=0){
    				frm.tx_suplycash.value=0;
    				frm.cksel.checked=true;
    				AnCheckClick(frm.cksel);
    			}
			}
		}
	}
}

function ShopSuplyPrice0(){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
    			if (frm.tx_shopbuyprice.value!=0){
    				frm.tx_shopbuyprice.value=0;
    				frm.cksel.checked=true;
    				AnCheckClick(frm.cksel);
    			}
    		}
		}
	}
}

function PopItemWeightEdit(iitemid){
	var popwin = window.open('/warehouse/pop_ItemWeightEdit.asp?itembarcode=' + iitemid + '&menupos=1103','itemWeightEdit','width=800,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드 :
		<% drawSelectBoxDesignerwithName "designer",designer  %>
		&nbsp;
		구분:
		<input type="radio" name="itemgubun" value="" <% if itemgubun = "" then response.write " checked" %>> 전체
		<input type="radio" name="itemgubun" value="85" <% if itemgubun = "85" then response.write " checked" %>> ON사은품(85)
		<input type="radio" name="itemgubun" value="80" <% if itemgubun = "80" then response.write " checked" %>> OFF사은품(80)
		&nbsp;
     	사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;
		무게입력 : <% drawSelectBoxUsingYN "weightYn", weightYn %>
		&nbsp;
		사이즈입력 : <% drawSelectBoxUsingYN "sizeYn", sizeYn %>
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="ReSearch('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" style="IME-MODE: disabled" />
		&nbsp;
		상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
		&nbsp;
		범용바코드 : <input type="text" class="text" name="publicbarcode" value="<%= publicbarcode %>" size="20" maxlength="20">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
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
		<input type="button" class="button" value="사은품 등록" onclick="pop_itemedit_gift_new()">
	</td>
	<td align="right">
		<% if ioffitem.FresultCount>0 then %>
			<input type="button" class="button" value="선택상품 매입가 0 설정" onclick="BuyPrice0()">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= ioffitem.FTotalcount %></b>
		<% if ioffitem.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>

		<b><%= page %> / <%= ioffitem.FTotalpage %></b>

		<% if (ioffitem.FTotalpage - ioffitem.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<% if (imageview<>"") then %>
	<td width="50">이미지</td>
	<% end if %>
	<td width="70">브랜드ID</td>
	<td width="90">상품코드</td>
	<td>상품명</td>
	<td>옵션명</td>

	<td width="90">매입가</td>
	<td width="30">센터<br>매입<br>구분</td>

	<td width="90">사용 여부</td>

    <td width="70">중량</td>
    <td width="100">사이즈</td>

	<td>비고</td>
</tr>
<% for i=0 to ioffitem.FresultCount -1 %>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
<input type="hidden" name="oldonlineprice" value="<%= ioffitem.FItemlist(i).FOnLineItemprice %>">
<input type="hidden" name="oldonlineorgprice" value="<%= ioffitem.FItemlist(i).FOnLineItemOrgprice %>">
<input type="hidden" name="oldonlineOptAddprice" value="<%= ioffitem.FItemlist(i).FOnlineOptaddprice %>">
<input type="hidden" name="extbarcode" value="<%= ioffitem.FItemlist(i).FextBarcode %>">
<input type="hidden" name="onofflinkyn" value="<%= ioffitem.FItemlist(i).Fonofflinkyn %>">

<% if ioffitem.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<% if (imageview<>"") then %>
		<td width="50">
			<img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0>
		</td>
	<% end if %>
	<td ><%= ioffitem.FItemlist(i).FMakerID %></td>
	<td>
		<a href="javascript:pop_itemedit_gift_edit('<%= ioffitem.FItemlist(i).GetBarCode %>')" onfocus="this.blur()">
		<%= ioffitem.FItemlist(i).GetBarCode %>
		</a>
	</td>
	<td>
		<a href="javascript:pop_itemedit_gift_edit('<%= ioffitem.FItemlist(i).GetBarCode %>')" onfocus="this.blur()">
		<%= ioffitem.FItemlist(i).FShopItemName %>
		</a>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).FShopitemOptionname %>
		<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
		    <br>옵션추가금액: <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>
		<% end if %>
	</td>
	<td align="right" >
		<input type="text" name="tx_suplycash" value="<%= ioffitem.FItemlist(i).Fshopsuplycash %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
	</td>
    <td align="center" ><%= ioffitem.FItemlist(i).FCenterMwDiv %></td>
	<td align="center" >
		<% if ioffitem.FItemlist(i).Fisusing="Y" then %>
			<input type="radio" name="isusing" value="Y" checked onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
			<input type="radio" name="isusing" value="N" onclick="CheckThis(frmBuyPrc_<%= i %>)">N
		<% else %>
			<input type="radio" name="isusing" value="Y" onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
			<input type="radio" name="isusing" value="N" checked onclick="CheckThis(frmBuyPrc_<%= i %>)"><font color="red">N</font>
		<% end if %>
	</td>
    <td align="center"><%= ioffitem.FItemlist(i).FitemWeight %> g</td>
    <td align="center"><%= ioffitem.FItemlist(i).FvolX %> * <%= ioffitem.FItemlist(i).FvolY %> * <%= ioffitem.FItemlist(i).FvolZ %> cm</td>
	<td align="center">
        <input type="button" class="button" value="사이즈수정" onclick="PopItemWeightEdit('<%= ioffitem.FItemlist(i).GetBarCode %>')">
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="25" align="center">
	<% if ioffitem.HasPreScroll then %>
		<a href="javascript:ReSearch('<%= ioffitem.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
		<% if i>ioffitem.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:ReSearch('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ioffitem.HasNextScroll then %>
		<a href="javascript:ReSearch('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<form name="frmArrupdate" method="post" action="/admin/offshop/shopitem_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="onofflinkynarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="orgsellpricearr" value="">
	<input type="hidden" name="itempricearr" value="">
	<input type="hidden" name="itemsuplyarr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="extbarcodearr" value="">
</form>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" class="button" value="사은품 등록" onclick="pop_itemedit_gift_new()">
	</td>
	<td align="right">
		<% if ioffitem.FresultCount>0 then %>
			<input type="button" class="button" value="선택상품 매입가 0 설정" onclick="BuyPrice0()">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
