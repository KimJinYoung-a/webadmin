<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/new_offshopitemcls.asp"-->
<%

dim makerid, page, imageview, ckonlyusing,research
makerid = request("makerid")
page = request("page")
imageview = request("imageview")
ckonlyusing = request("ckonlyusing")
research= request("research")

if page="" then page=1
if research<>"on" then ckonlyusing="on"

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 100
ioffitem.FCurrPage = page
ioffitem.FRectMakerid = makerid
ioffitem.FRectOnlyUsing = ckonlyusing

ioffitem.GetDiffItemMarginList

dim i
%>
<script language='javascript'>
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

function ReSearch(page){
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

	//upfrm.chargeidarr.value = "";
	upfrm.extbarcodearr.value = "";
	upfrm.shopitemnamearr.value = "";

	upfrm.discountsellpricearr.value = "";
	upfrm.shopbuypricearr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				//if (!ChargeIdAvail(frm.tx_charge.value)){
				//	alert(frm.tx_charge.value + '는 올바른 아이디가 아닙니다.');
				//	frm.tx_charge.focus();
				//	return;

				//}

				if (!IsDigit(frm.tx_sellcash.value)){
					alert('판매가는 숫자만 가능합니다.');
					frm.tx_sellcash.focus();
					return;
				}

				if (frm.tx_sellcash.value<10){
					alert('판매가는 10원보다 커야 합니다.');
					frm.tx_sellcash.focus();
					return;
				}


				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.itempricearr.value = upfrm.itempricearr.value + frm.tx_sellcash.value + "|";
				upfrm.itemsuplyarr.value = upfrm.itemsuplyarr.value + frm.tx_suplycash.value + "|";

				upfrm.discountsellpricearr.value = upfrm.discountsellpricearr.value + frm.tx_discountsellprice.value + "|";
				upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.tx_shopbuyprice.value + "|";

				//upfrm.chargeidarr.value = upfrm.chargeidarr.value + frm.tx_charge.value + "|";
				upfrm.extbarcodearr.value = upfrm.extbarcodearr.value + frm.extbarcode.value + "|";

				upfrm.shopitemnamearr.value = upfrm.shopitemnamearr.value + frm.shopitemname.value + "|";

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
	frm.tx_sellcash.value=frm.oldonlineprice.value;
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

			if (frm.tx_suplycash.value!=0){
				frm.tx_suplycash.value=0;
				frm.cksel.checked=true;
				AnCheckClick(frm.cksel);
			}
		}
	}
}

function ShopSuplyPrice0(){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (frm.tx_shopbuyprice.value!=0){
				frm.tx_shopbuyprice.value=0;
				frm.cksel.checked=true;
				AnCheckClick(frm.cksel);
			}
		}
	}
}
</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>오프샾 개별 마진 상품관리</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			업체 기본 마진과 상품 마진이 다를 경우 (이벤트 상품, 할인 상품 등)<br>
			매입가 또는 공급가를 <b>0이 아닌 값</b>으로 입력하시면<br>
			개별적으로 업체 매입가 또는 샾 공급가를 설정할 수 있습니다.
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>
<br>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name=frm method=get>
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">

   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	업체:<% drawSelectBoxDesignerOffShopContract "makerid",makerid  %>
	        	&nbsp;&nbsp;
	        	<input type="checkbox" name="ckonlyusing" value="on" <% if ckonlyusing="on" then response.write "checked" %> >사용중인상품만
	        	&nbsp;&nbsp;
				<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
	        </td>
	        <td valign="top" align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr bgcolor="#DDDDFF">
    	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
    	<% if (imageview<>"") then %>
    	<td width="50">이미지</td>
    	<% end if %>
    	<td width="70">브랜드</td>
    	<td width="70">텐바이텐<br>바코드</td>
    	<td >상품명</td>
    	<td width="40">옵션명</td>
    	<td width="60">OnLine<br>판매가</td>
    	<td width="60">판매가</td>
    	<td width="60">할인<br>판매가</td>
    	<td width="60">매입가</td>
    	<td width="60">샾<br>공급가</td>
    	<td width="30">매입<br>마진</td>
    	<td width="30">공급<br>마진</td>
    	<td width="34">사용<br>여부</td>
	</tr>
    <% for i=0 to ioffitem.FResultCount - 1 %>
    <form name="frmBuyPrc_<%= i %>" >
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
  		<td><a href="javascript:popOffItemEdit('<%= ioffitem.FItemlist(i).GetBarCode %>');"><%= ioffitem.FItemlist(i).Fitemgubun %><%=  Format00(6,ioffitem.FItemlist(i).Fshopitemid) %><%= ioffitem.FItemlist(i).Fitemoption %></a></td>
  		<td ><%= replace(ioffitem.FItemlist(i).FShopItemName,"|","") %></td>
  		<td align="center" ><%= ioffitem.FItemlist(i).FShopitemOptionname %></td>
  		<% if ioffitem.FItemlist(i).FOnLineItemprice>ioffitem.FItemlist(i).FShopItemprice then %>
  		<td align="center" ><font color="red"><b><%= ioffitem.FItemlist(i).FOnLineItemprice %></b></font></td>
  		<% elseif ioffitem.FItemlist(i).FOnLineItemprice<ioffitem.FItemlist(i).FShopItemprice then %>
  		<td align="center" ><font color="red"><%= ioffitem.FItemlist(i).FOnLineItemprice %></font></td>
  		<% else %>
  		<td align="center" ><%= ioffitem.FItemlist(i).FOnLineItemprice %></td>
  		<% end if %>
  		<td align="right" ><input type="text" name="tx_sellcash" value="<%= ioffitem.FItemlist(i).FShopItemprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>
  		<td align="right" ><input type="text" name="tx_discountsellprice" value="<%= ioffitem.FItemlist(i).Fdiscountsellprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>

  		<td align="right" ><input type="text" name="tx_suplycash" value="<%= ioffitem.FItemlist(i).Fshopsuplycash %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>
  		<td align="right" ><input type="text" name="tx_shopbuyprice" value="<%= ioffitem.FItemlist(i).Fshopbuyprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>

  		<td align="right" >
  		<% if ioffitem.FItemlist(i).FShopItemprice<>0 then %>
  			<%= Fix((1-ioffitem.FItemlist(i).Fshopsuplycash/ioffitem.FItemlist(i).FShopItemprice)*100) %>%
  		<% end if %>
  		</td>
  		<td align="right" >
  		<% if ioffitem.FItemlist(i).FShopItemprice<>0 then %>
  			<%= Fix((1-ioffitem.FItemlist(i).Fshopbuyprice/ioffitem.FItemlist(i).FShopItemprice)*100) %>%
  		<% end if %>
  		</td>
  		<td align="left" >
  		<% if ioffitem.FItemlist(i).Fisusing="Y" then %>
  		<input type="radio" name="isusing" value="Y" checked >Y
  		<input type="radio" name="isusing" value="N">N
  		<% else %>
  		<input type="radio" name="isusing" value="Y">Y
  		<input type="radio" name="isusing" value="N" checked ><font color="red">N</font>
  		<% end if %>
  		</td>
  	</tr>
  	</form>
    <% next %>
    <tr bgcolor="#FFFFFF">
		<td colspan="<% if (imageview<>"") then response.write "16" else response.write "15" end if %>" align="center">
		<% if ioffitem.HasPreScroll then %>
			<a href="javascript:ReSearch('<%= ioffitem.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ioffitem.StarScrollPage to ioffitem.FScrollCount + ioffitem.StarScrollPage - 1 %>
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

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="10">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set ioffitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->