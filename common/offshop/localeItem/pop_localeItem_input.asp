<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 지역별용 상품 리스트
' History : 2010.08.05 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer,page,usingyn ,research,pricediff,imageview, pricelow
dim itemgubun, itemid, itemname ,cdl, cdm, cds ,onexpire ,shopid , strparm
dim i, PriceDiffExists
	designer    = RequestCheckVar(request.form("designer"),32)
	page        = RequestCheckVar(request.form("page"),10)
	usingyn     = RequestCheckVar(request.form("usingyn"),1)
	research    = RequestCheckVar(request.form("research"),9)
	pricediff   = RequestCheckVar(request.form("pricediff"),9)
	pricelow    = RequestCheckVar(request.form("pricelow"),9)
	imageview   = RequestCheckVar(request.form("imageview"),9)
	onexpire    = RequestCheckVar(request.form("onexpire"),9)
	itemgubun   = RequestCheckVar(request.form("itemgubun"),2)
	itemid      = RequestCheckVar(request.form("itemid"),9)
	itemname    = RequestCheckVar(request.form("itemname"),32)
	cdl         = RequestCheckVar(request.form("cdl"),3)
	cdm         = RequestCheckVar(request.form("cdm"),3)
	cds         = RequestCheckVar(request.form("cds"),3)
	shopid    = RequestCheckVar(request.form("shopid"),10)

	if page="" then page=1
	if research<>"on" then usingyn="Y"
	strparm = "designer="&designer&"&usingyn="&usingyn&""
	strparm = strparm & "&research="&research&"&pricediff="&pricediff&"&pricelow="&pricelow&"&imageview="&imageview&"&onexpire="&onexpire&""
	strparm = strparm & "&itemgubun="&itemgubun&"&itemid="&itemid&"&itemname="&itemname&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds

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
	    ioffitem.FRectPriceRow = pricelow
		ioffitem.GetOffShopPriceDiffItemList
	else
		ioffitem.GetOffNOnLineShopItemList
	end if
%>

<script language="javascript">

function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

//선택상품 추가
function iteminsert(upfrm){

if (upfrm.shopid.value.length < 1){
	alert('저장하실 매장을 선택해주세요');
	return;
}

if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "," ;
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "," ;
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "," ;

				}
			}
		}

	upfrm.mode.value='itemadd';
	upfrm.action = "/admin/offshop/localeitem/localeItem_process.asp";
	upfrm.submit();
}

function reg(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="sType" >
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemoptionarr">
<input type="hidden" name="itemgubunarr">
<input type="hidden" name="page">
<input type="hidden" name="mode">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드 : <% drawSelectBoxDesignerwithName "designer",designer  %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;
		상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		상품구분:<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
     	&nbsp;
     	오프사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
     	&nbsp;
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
		&nbsp;
		<input type="checkbox" name="pricediff" value="on" <% if pricediff="on" then response.write "checked" %> >가격상이만 보기
		&nbsp;
		<input type="checkbox" name="pricelow" value="on" <% if pricelow="on" then response.write "checked" %> >온라인보다 작은가격
		&nbsp;
		<input type="checkbox" name="onexpire" value="on" <% if onexpire="on" then response.write "checked" %> >ON품절+단종+사용안함(신상품제외)
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
<tr>
	<td valign="bottom">
		매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
		<input type="button" value="선택상품 추가" onClick="iteminsert(frm)" class="button">
		※이미 등록되어 있는 상품은 빼고 저장 됩니다
	</td>
</tr>
</form>
</table>

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
<% if ioffitem.FresultCount > 0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<% if (imageview<>"") then %>
	<td width="50">이미지</td>
	<% end if %>
	<td width="70">브랜드ID</td>
	<td width="90">상품코드</td>
	<td>상품명<font color="blue">[옵션명]</font></td>
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
</tr>

<% for i=0 to ioffitem.FresultCount -1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">

<% if ioffitem.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<td  align="center">
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		<input type="hidden" name="shopid" value="<%=shopid%>">
		<input type="hidden" name="itemid" value="<%=ioffitem.FItemlist(i).FShopitemid%>">
		<input type="hidden" name="itemoption" value="<%=ioffitem.FItemlist(i).Fitemoption%>">
		<input type="hidden" name="itemgubun" value="<%=ioffitem.FItemlist(i).fitemgubun%>">
	</td>
	<% if (imageview<>"") then %>
	<td width="50"><img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></td>
	<% end if %>
	<td ><%= ioffitem.FItemlist(i).FMakerID %></td>
	<td><%= ioffitem.FItemlist(i).Fitemgubun %>-<%=  FormatCode(ioffitem.FItemlist(i).Fshopitemid) %>-<%= ioffitem.FItemlist(i).Fitemoption %></td>
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
    <td align="right" >
        <%= FormatNumber(ioffitem.FItemlist(i).FShopItemOrgprice,0) %>
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
	    <%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %>
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

	<td align="right" ><%= FormatNumber(ioffitem.FItemlist(i).Fshopsuplycash,0) %></td>
	<td align="right" ><%= FormatNumber(ioffitem.FItemlist(i).Fshopbuyprice,0) %></td>

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
	<td align="right" ><%= ioffitem.FItemlist(i).FextBarcode %></td>
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if ioffitem.HasPreScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=ioffitem.StartScrollPage-1%>);">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ioffitem.StartScrollPage to ioffitem.StartScrollPage + ioffitem.FScrollCount - 1 %>
			<% if (i > ioffitem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ioffitem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ioffitem.HasNextScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->