<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->


<%
dim designer,page,usingyn
dim research,pricediff,imageview
dim itemgubun, itemid, itemname
dim cdl, cdm, cds
dim onexpire
dim shopid, offgubun

designer    = RequestCheckVar(request("designer"),32)
page        = RequestCheckVar(request("page"),9)
usingyn     = RequestCheckVar(request("usingyn"),1)
research    = RequestCheckVar(request("research"),9)
pricediff   = RequestCheckVar(request("pricediff"),9)
imageview   = RequestCheckVar(request("imageview"),9)
onexpire    = RequestCheckVar(request("onexpire"),9)

itemgubun   = RequestCheckVar(request("itemgubun"),2)
itemid      = RequestCheckVar(request("itemid"),9)
itemname    = RequestCheckVar(request("itemname"),32)

cdl         = RequestCheckVar(request("cdl"),3)
cdm         = RequestCheckVar(request("cdm"),3)
cds         = RequestCheckVar(request("cds"),3)

''차후 session("ssAdminPsn")="6" : 부서번호만 사용할것.
if session("ssBctDiv")="201" or session("ssAdminPsn")="6" then
	shopid = "cafe002"
	offgubun = "CAF"
	designer = "menu002"
elseif session("ssBctDiv")="301" or session("ssAdminPsn")="16" then
	shopid = "cafe003"
	offgubun = "CAF"
	designer = "menu003"
else
    ''계정별 fix
    if ((session("ssBctDiv")>9) and (session("ssBctBigo")<>"")) then shopid=session("ssBctBigo")
    designer = "menu091"
end if



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

ioffitem.FRectOnlineExpiredItem = onexpire

if pricediff="on" then
	ioffitem.GetOffShopPriceDiffItemList
else
	ioffitem.GetOffNOnLineShopItemList
end if

dim i, PriceDiffExists


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

function popOffItemEdit(ibarcode){
	var popwin = window.open('popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popOffImageEdit(ibarcode){
	var popwin = window.open('popoffimageedit.asp?barcode=' + ibarcode,'popoffimageedit','width=500,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function OffItemReg(idesigner){

	var subwin;

	if (confirm('코드관리에 문제가 있으니 온라인에 등록되있거나\n입점예정인 상품은\n\n----------------절대------------- \n\n등록하지 말아 주세요. 계속하시겠습니까?')){
		subwin = window.open('shopoffitemreg.asp?designer=' + idesigner,'window_reg','width=500,height=300,scrollbars=yes,status=no');
		subwin.focus();
	}
}

function OffEtcItemReg(makerid){

	var subwin;

	if (confirm('오프라인 기타 메뉴를 등록하는 작업입니다.계속하시겠습니까?')){
		subwin = window.open('popoffitemreg_Etc.asp?makerid=' + makerid,'window_reg','width=500,height=300,scrollbars=yes,resizable=yes,status=no');
		subwin.focus();
	}
}

function ReSearch(page){
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

	upfrm.extbarcodearr.value = "";

	upfrm.shopbuypricearr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

                if (!IsDigit(frm.tx_orgsellprice.value)){
					alert('소비자가가는 숫자만 가능합니다.');
					frm.tx_orgsellprice.focus();
					return;
				}

				if (!IsDigit(frm.tx_sellcash.value)){
					alert('판매가는 숫자만 가능합니다.');
					frm.tx_sellcash.focus();
					return;
				}

				if (frm.tx_sellcash.value<10){
					if (!confirm('판매가는 10원보다 커야 합니다. 계속 진행하시겠습니까?')){
					    frm.tx_sellcash.focus();
					    return;
					}
				}

                if (frm.tx_orgsellprice.value*1<frm.tx_sellcash.value*1){
					alert('소비자가는 판매가보다 커야합니다..');
					frm.tx_orgsellprice.focus();
					return;
				}

                //매입가 공급가 체크
                if ((frm.tx_suplycash.value*1!=0)&&(frm.tx_suplycash.value*1!=0)){
                    if (frm.tx_suplycash.value*1>frm.tx_shopbuyprice.value*1){
    					alert('샵 공급가는 매입가보다 커야합니다..');
    					frm.tx_suplycash.focus();
    					return;
    				}
				}

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";

				upfrm.orgsellpricearr.value = upfrm.orgsellpricearr.value + frm.tx_orgsellprice.value + "|";
				upfrm.itempricearr.value = upfrm.itempricearr.value + frm.tx_sellcash.value + "|";
				upfrm.itemsuplyarr.value = upfrm.itemsuplyarr.value + frm.tx_suplycash.value + "|";

				upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.tx_shopbuyprice.value + "|";

				upfrm.extbarcodearr.value = upfrm.extbarcodearr.value + frm.extbarcode.value + "|";

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
</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 : <% drawSelectBoxDesignerwithName "designer",designer  %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			<br>
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9">
			&nbsp;
			상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		<!--
			상품구분:<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
	     	&nbsp;
	     -->
	     	오프사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;
			<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
		<!--
			&nbsp;
			<input type="checkbox" name="pricediff" value="on" <% if pricediff="on" then response.write "checked" %> >가격상이만 보기
			&nbsp;
			<input type="checkbox" name="onexpire" value="on" <% if onexpire="on" then response.write "checked" %> >ON품절+단종+사용안함(신상품제외)
		-->
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->


<p>

<!-- 액션 시작 -->

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		    <input type="button" class="button" value="메뉴 상품등록" onclick="OffEtcItemReg('<%= designer %>')">
		    <% if ioffitem.FresultCount>0 then %>
		    &nbsp;
			<input type="button" class="button" value="선택상품 매입가 0 설정" onclick="BuyPrice0()">
			&nbsp;
			<input type="button" class="button" value="선택상품 샵공급가 0 설정" onclick="ShopSuplyPrice0()">
			&nbsp;
			<% end if %>
		</td>
		<td align="right">
		    <% if ioffitem.FresultCount>0 then %>
			<input type="button" class="button" value="선택상품 일괄수정" onclick="ModiArr()">
			<% end if %>
		</td>
	</tr>
</table>

<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18">

			검색결과 : <b><%= ioffitem.FTotalcount %></b>
			&nbsp;
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
    	<td>상품명<font color="blue">[옵션명]</font></td>
    	<td width="20"></td>
    	<td width="50">소비자가</td>
    	<td width="50">판매가</td>
    	<td width="40">할인율<br>(%)</td>
    	<td width="50">매입가</td>
    	<td width="50">샾공급가</td>
    	<td width="30">매입<br>마진</td>
    	<td width="30">공급<br>마진</td>
    	<td width="30">센터<br>매입<br>구분</td>
    	<td width="30">ON<br>판매</td>
    	<td width="30">ON<br>단종</td>
    	<td width="60">범용바코드</td>
    	<td width="70">사용 여부<br><input class="button" type="button" value="사용안함" onClick="NotUsingCheckAll();"></td>
	</tr>
	<% for i=0 to ioffitem.FresultCount -1 %>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
	<input type="hidden" name="oldonlineprice" value="<%= ioffitem.FItemlist(i).FOnLineItemprice %>">
	<input type="hidden" name="oldonlineorgprice" value="<%= ioffitem.FItemlist(i).FOnLineItemOrgprice %>">
	<input type="hidden" name="oldonlineOptAddprice" value="<%= ioffitem.FItemlist(i).FOnlineOptaddprice %>">

	<% if ioffitem.FItemlist(i).Fisusing="N" then %>
	<tr bgcolor="#EEEEEE">
	<% else %>
	<tr bgcolor="#FFFFFF">
	<% end if %>
  		<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
  		<% if (imageview<>"") then %>
  		<td width="50"><a href="javascript:popOffImageEdit('<%= ioffitem.FItemlist(i).GetBarCode %>')"><img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></a></td>
  		<% end if %>
  		<td ><%= ioffitem.FItemlist(i).FMakerID %></td>
  		<td><a href="javascript:popOffItemEdit('<%= ioffitem.FItemlist(i).GetBarCode %>');"><%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %></a></td>
  		<td>
  			<%= ioffitem.FItemlist(i).FShopItemName %>
  			<% if ioffitem.FItemlist(i).Fitemoption<>"0000" then %>
  				<font color="blue">[<%= ioffitem.FItemlist(i).FShopitemOptionname %>]</font>
  			<% end if %>

  			<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
  			    <br>옵션추가금액: <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>
  			<% end if %>
  		</td>
        <% PriceDiffExists = false %>
  		<td align="center" >
  		    <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
  		    <% if (ioffitem.FItemlist(i).FOnlineitemorgprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemOrgprice) or (ioffitem.FItemlist(i).FOnLineItemprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemprice) then %>
  		    <input type="button" class="button" value=">" onclick="samePrice(frmBuyPrc_<%= i %>);">
  		    <% PriceDiffExists = true %>
  		    <% end if %>
  		    <% end if %>
  		</td>
        <td align="right" >
            <input type="text" class="text" name="tx_orgsellprice" value="<%= ioffitem.FItemlist(i).FShopItemOrgprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
            <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
            <% if (ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemOrgprice)  then %>
                <font color="red"><strong><%= ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
            <% else %>
                <% if (PriceDiffExists) then %>
                <%= ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice %>
                <% end if %>
            <% end if %>
            <% end if %>
        </td>
  		<td align="right" >
  		    <input type="text" class="text" name="tx_sellcash" value="<%= ioffitem.FItemlist(i).FShopItemprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
  		    <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
            <% if (ioffitem.FItemlist(i).FOnLineItemprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemprice)  then %>
  		        <font color="red"><strong><%= ioffitem.FItemlist(i).FOnLineItemprice + ioffitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
  		    <% else %>
  		        <% if (PriceDiffExists) then %>
  		        <%= ioffitem.FItemlist(i).FOnLineItemprice + ioffitem.FItemlist(i).FOnlineOptaddprice %>
  		        <% end if %>
            <% end if %>
            <% end if %>
  		</td>
  		<td align="center" >
            <% if (ioffitem.FItemlist(i).FShopItemOrgprice<>0) then %>
                <% if ioffitem.FItemlist(i).FShopItemOrgprice<>ioffitem.FItemlist(i).FShopItemprice then %>
                OFF:<font color="#FF3333"><%= CLng((ioffitem.FItemlist(i).FShopItemOrgprice-ioffitem.FItemlist(i).FShopItemprice)/ioffitem.FItemlist(i).FShopItemOrgprice*100*100)/100 %></font>
                <% end if %>
  		    <% end if %>

  		    <% if (ioffitem.FItemlist(i).FOnlineitemorgprice<>0) then %>
  		        <% if ioffitem.FItemlist(i).FOnlineitemorgprice<>ioffitem.FItemlist(i).FOnLineItemprice then %>
                ON:<font color="#FF3333"><%= CLng((ioffitem.FItemlist(i).FOnlineitemorgprice-ioffitem.FItemlist(i).FOnLineItemprice)/ioffitem.FItemlist(i).FOnlineitemorgprice*100*100)/100 %></font>
                <% end if %>
  		    <% end if %>
  		</td>

  		<td align="right" ><input type="text" name="tx_suplycash" value="<%= ioffitem.FItemlist(i).Fshopsuplycash %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>
  		<td align="right" ><input type="text" name="tx_shopbuyprice" value="<%= ioffitem.FItemlist(i).Fshopbuyprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>

  		<td align="right" >
  		<% if (ioffitem.FItemlist(i).FShopItemprice<>0) and (ioffitem.FItemlist(i).Fshopsuplycash<>0) then %>
  			<font color="blue"><%= CLng((ioffitem.FItemlist(i).FShopItemprice-ioffitem.FItemlist(i).Fshopsuplycash)/ioffitem.FItemlist(i).FShopItemprice*100) %>%</font>
  		<% end if %>
  		</td>
  		<td align="right" >
  		<% if (ioffitem.FItemlist(i).FShopItemprice<>0) and (ioffitem.FItemlist(i).Fshopbuyprice<>0) then %>
  			<font color="blue"><%= CLng((ioffitem.FItemlist(i).FShopItemprice-ioffitem.FItemlist(i).Fshopbuyprice)/ioffitem.FItemlist(i).FShopItemprice*100) %>%</font>
  		<% end if %>
  	    </td>
  	    <td align="center" ><%= ioffitem.FItemlist(i).FCenterMwDiv %></td>
  	    <td align="center" ><%= fnColor(ioffitem.FItemlist(i).Fsellyn,"sellyn") %></td>
  	    <td align="center" ><%= fnColor(ioffitem.FItemlist(i).FonLineDanjongyn,"dj") %></td>
  		<td align="right" ><input type="text" name="extbarcode" value="<%= ioffitem.FItemlist(i).FextBarcode %>" size="12" maxlength="20" style="border:1px #999999 solid; " onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>
  		<td align="left" >
  		<% if ioffitem.FItemlist(i).Fisusing="Y" then %>
  		<input type="radio" name="isusing" value="Y" checked onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
  		<input type="radio" name="isusing" value="N" onclick="CheckThis(frmBuyPrc_<%= i %>)">N
  		<% else %>
  		<input type="radio" name="isusing" value="Y" onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
  		<input type="radio" name="isusing" value="N" checked onclick="CheckThis(frmBuyPrc_<%= i %>)"><font color="red">N</font>
  		<% end if %>
  		</td>
  	</tr>
  	</form>
  	<% next %>
  	<tr bgcolor="#FFFFFF">
		<td colspan="18" align="center">
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
</table>

<form name="frmArrupdate" method="post" action="shopitem_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="orgsellpricearr" value="">
<input type="hidden" name="itempricearr" value="">
<input type="hidden" name="itemsuplyarr" value="">
<input type="hidden" name="shopbuypricearr" value="">
<input type="hidden" name="isusingarr" value="">
<input type="hidden" name="extbarcodearr" value="">

</form>
<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->