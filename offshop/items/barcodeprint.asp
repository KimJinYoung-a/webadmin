<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->


<%
dim shopid
dim designer
dim imageview
dim cdl, cdm, cds
dim itemid, itemname, shopitemname
dim prdcode, generalbarcode
dim tmpimgurl



'==============================================================================
shopid 			= RequestCheckVar(request("shopid"),32)

designer 		= RequestCheckVar(request("designer"),32)
imageview 		= RequestCheckVar(request("imageview"),32)

cdl         	= RequestCheckVar(request("cdl"),3)
cdm         	= RequestCheckVar(request("cdm"),3)
cds         	= RequestCheckVar(request("cds"),3)

itemid 			= RequestCheckVar(request("itemid"),32)
itemname 		= RequestCheckVar(request("itemname"),32)
shopitemname 	= RequestCheckVar(request("shopitemname"),32)

prdcode 		= RequestCheckVar(request("prdcode"),32)
generalbarcode 	= RequestCheckVar(request("generalbarcode"),32)



'==============================================================================
''유저구분이 가맹점인경우 박아 넣는다
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if



'==============================================================================
dim obarcode

set obarcode = new COffShopItem

obarcode.FRectShopid = shopid
obarcode.FRectDesigner = designer
'obarcode.FRectBarCode = barcode
'obarcode.FRectItemId = itemid

obarcode.FRectCDL = cdl
obarcode.FRectCDM = cdm
obarcode.FRectCDS = cds

obarcode.FRectItemID = itemid

obarcode.FRectItemName = html2db(itemname)
obarcode.FRectShopItemName = html2db(shopitemname)

obarcode.FRectPrdCode = prdcode
obarcode.FRectGeneralBarcode = generalbarcode

if (designer<>"") or (itemid<>"") or (itemname<>"") or (shopitemname<>"") or (prdcode<>"") or (generalbarcode<>"") then
	obarcode.GetBarCodeList
end if

dim i
%>
<script language='javascript'>
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function AddData(itemid, itemoption, itemname, itemoptionname, brand, itemprice, itemtype, itemno){
	iaxobject.AddData(itemid, itemoption, itemname, itemoptionname, brand, itemprice, itemtype, itemno);
}

//AddData(v,'0000','아이템명','옵션명','브랜드',3000,'T','5')
function AddArr(){
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
	iaxobject.ClearItem();
	//iaxobject.setTitleVisible(true);
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
			    if (frm.itemid.value*1>=1000000){
			        AddData(frm.itemid.value,frm.itemoption.value,frm.itemname.value,frm.itemoptionname.value,frm.brand.value,frm.sellprice.value,frm.itemgubun.value*10,1);
			    }else{
				    AddData(frm.itemid.value,frm.itemoption.value,frm.itemname.value,frm.itemoptionname.value,frm.brand.value,frm.sellprice.value,frm.itemgubun.value,1);
				}
			}
		}
	}
	iaxobject.ShowFrm();
}

</script>
<OBJECT
	  id=iaxobject
	  classid="clsid:5D776FEA-8C6B-4C53-8EC3-3585FC040BDB"
	  codebase="http://webadmin.10x10.co.kr/common/cab/tenbarPrint.cab#version=1,0,0,29"
	  width=0
	  height=0
	  align=center
	  hspace=0
	  vspace=0
>
</OBJECT>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    매장 :
		<% if (C_IS_SHOP) then %>
		<%= shopid %><input type="hidden" name="shopid" value="<%= shopid %>">
		<% else %>
		<% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
		<% end if %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		브랜드 : <% drawSelectBoxDesignerwithName "designer",designer %>
		&nbsp;
		상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) 	frm.submit();">
		&nbsp;
		상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) frm.submit();">
		<!--
		&nbsp;
		샵별상품명 : <input type="text" class="text" name="shopitemname" value="<%= shopitemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) frm.submit();">
		-->
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		물류코드 :
		<input type="text" class="text" name="prdcode" value="<%= prdcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) frm.submit();">
		&nbsp;
		범용바코드 :
		<input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>
* 상품명에 특수문자가 있는 경우 검색되지 않습니다.
<p>

<% if obarcode.FResultCount>0 then %>
<table width="800" cellspacing="1" class="a" >
	<tr bgcolor="#FFFFFF">
      <td colspan="6" align="left"><input type="button" class="button" value="선택내역출력" onclick="AddArr()"></td>
    </tr>
</table>
<% end if %>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="<%= adminColor("tabletop") %>">
      <td width="20" align="center"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
<% if imageview="on" then %>
	  <td width="55" align="center">이미지</td>
<% end if %>
      <td width="100" align="center">물류코드</td>
      <td align="center">상품명</td>
      <td width="200" align="center">옵션명</td>
      <td width="80" align="center">판매가</td>
    </tr>
    <% if obarcode.FResultCount<1 then %>
    <tr bgcolor="#FFFFFF">
      <td colspan="6" align="center">검색 결과가 없습니다.</td>
    </tr>
    <% else %>

	    <% for i=0 to obarcode.FResultCount-1 %>
	    	<%

	    	if (obarcode.FItemList(i).FimageSmall <> "") then
	    		tmpimgurl = obarcode.FItemList(i).FimageSmall
	    	else
	    		tmpimgurl = obarcode.FItemList(i).FOffimgSmall
	    	end if

	    	%>
	    <form name="frmBuyPrc_<%= i %>" >
	    <input type="hidden" name="itemid" value="<%= obarcode.FItemList(i).Fshopitemid %>">
	    <input type="hidden" name="itemoption" value="<%= obarcode.FItemList(i).Fitemoption %>">
	    <input type="hidden" name="itemname" value="<%= obarcode.FItemList(i).Fshopitemname %>">
	    <input type="hidden" name="itemoptionname" value="<%= obarcode.FItemList(i).Fshopitemoptionname %>">
	    <input type="hidden" name="brand" value="<%= obarcode.FItemList(i).FSocName %>">
	    <input type="hidden" name="sellprice" value="<%= obarcode.FItemList(i).Fshopitemprice %>">
	    <input type="hidden" name="itemgubun" value="<%= obarcode.FItemList(i).Fitemgubun %>">
	    <tr bgcolor="#FFFFFF">
	      <td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			<% if imageview="on" then %>
	  	  <td align="center"><img src="<%= tmpimgurl %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></td>
			<% end if %>
	      <td align="center"><%= obarcode.FItemList(i).GetBarCode %></td>
	      <td ><%= obarcode.FItemList(i).Fshopitemname %></td>
	      <td ><%= obarcode.FItemList(i).Fshopitemoptionname %></td>
	      <td align="center"><%= FormatNumber(obarcode.FItemList(i).Fshopitemprice,0) %></td>
	    </tr>
	    </form>
	    <% next %>
    <% end if %>
</table>
<%
set obarcode = Nothing
%>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->