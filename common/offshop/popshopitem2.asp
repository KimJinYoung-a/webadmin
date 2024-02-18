<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 입출고 리스트 상품추가
' Hieditor : 2009.04.07 서동석 생성
'			 2010.08.04 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%

dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False

'' =============================================================================
'' 아래 3개의 메뉴 검색조건은 기본적으로 동일해야 한다.
'' (매장 적정재고 부족상품, 주문관리(물류), 주문관리(업 체))
'' =============================================================================
'' /common/offshop/stock/shortagestock_shop.asp
'' /common/offshop/popshopitem2.asp
'' /common/offshop/popshopjumunitem.asp
'' =============================================================================

dim PriceEditEnable ,page, chargeid, shopid ,isusing, itemname, imageon, research , i
dim mode , onlyActive, itemid
dim ipgo, logicsipgo, sell7days, includepreorder, shortagetype, onlinemwdiv, ordby
dim cp_idx
dim forcemakerid

PriceEditEnable = false

''업체인경우, 오프샵 관리자인경우
if (C_IS_Maker_Upche) then
	chargeid = session("ssBctID")
else
	chargeid = RequestCheckVar(request("chargeid"),32)
end if

''if Not (C_IS_SHOP) and Not (C_IS_Maker_Upche) then PriceEditEnable = true
	itemid  = RequestCheckVar(request("itemid"),255)
	onlyActive = RequestCheckVar(request("onlyActive"),32)
	mode = RequestCheckVar(request("mode"),32)
	page = RequestCheckVar(request("page"),10)
	shopid  = RequestCheckVar(request("shopid"),32)
	isusing = RequestCheckVar(request("isusing"),1)
	itemname = RequestCheckVar(request("itemname"),124)
	imageon = RequestCheckVar(request("imageon"),2)
	research = RequestCheckVar(request("research"),2)
	ipgo = RequestCheckVar(request("ipgo"),10)
    logicsipgo = RequestCheckVar(request("logicsipgo"),32)
	sell7days = RequestCheckVar(request("sell7days"),10)
	includepreorder = RequestCheckVar(request("includepreorder"),10)
	shortagetype = RequestCheckVar(request("shortagetype"),10)
	onlinemwdiv = RequestCheckVar(request("onlinemwdiv"),1)
	ordby = RequestCheckVar(request("ordby"),10)
	cp_idx = RequestCheckVar(request("cp_idx"),20)
    forcemakerid = RequestCheckVar(request("forcemakerid"),32)

if page="" then page=1
if research="" then imageon="on"
if research="" and isusing="" then isusing="Y"
if mode="" then mode="bybrand"

if research = "" Then
	If (cp_idx = "") Then
		'ipgo = "on"
		logicsipgo = "on"
		includepreorder = "on"
		shortagetype = "7"
		mode = "all"
	Else
		'
	End If
end if

if C_ADMIN_USER or C_IS_OWN_SHOP then

'' 매장인경우
elseif (C_IS_SHOP) then
    isusing="Y"
	IS_HIDE_BUYCASH = True
end if

if (research<>"on") and (ordby="") then
    ordby = "BI"
end if

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,",",chr(10))
  	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim ioffitem
	set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 50
	ioffitem.FCurrPage = page
    if forcemakerid = "" then
	    ioffitem.FRectDesigner = chargeid
    else
        ioffitem.FRectDesigner = forcemakerid
    end if
	ioffitem.FRectshopid = shopid
	ioffitem.FRectOnlyUsing = isusing
	ioffitem.FRectItemName = Html2Db(itemname)
	ioffitem.FRectOnlyActive = onlyActive
	ioffitem.FRectOrder = mode
	ioffitem.FRectItemid = itemid
	ioffitem.FRectIpGoOnly = ipgo
    ioffitem.FRectLogicsIpGoOnly = logicsipgo
	ioffitem.FRectSell7days = sell7days
	ioffitem.FRectIncludePreOrder = includepreorder
	ioffitem.FRectShortageType = shortagetype
	ioffitem.FRectOnlineMWdiv = onlinemwdiv
	If (cp_idx <> "") Then
		ioffitem.FPageSize = 250
		ioffitem.FRectCopyIdx = cp_idx
	End If

	''ioffitem.FRectOrder = ordby

	if chargeid<>"" then
		if (shopid<>"") then
	        ioffitem.GetOffShopItemList
	    else
	        response.write "<script type='text/javascript'>alert('매장이 지정 되지 않았습니다. ');</script>"
	    end if
	end if
%>

<script type='text/javascript'>

function ReSearch(page){
	frm.page.value= page;
	frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function RefreshParent(){
	opener.ReAct();
}

function AnSearch(frm){
	frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
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
	upfrm.shopbuypricearr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.itemnamearr.value = "";
	upfrm.itemoptionnamearr.value = "";
	upfrm.designerarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if (!IsDigit(frm.sellcash.value)){
					alert('판매가는 숫자만 가능합니다.');
					frm.sellcash.focus();
					return;
				}

				if (frm.suplycash.value*0 != 0) {
					alert('공급가는 숫자만 가능합니다.');
					frm.suplycash.focus();
					return;
				}

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
				upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.shopbuyprice.value + "|";
				upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
				upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
				upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
				upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
			}
		}
	}

	opener.ReActItems(upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.shopbuypricearr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value);

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
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="chargeid" value="<%= chargeid %>">
<input type="hidden" name="cp_idx" value="<%= cp_idx %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        <% if C_ADMIN_AUTH then %>
        * 브랜드[관리자] <input type="text" class="text" name="forcemakerid" value="<%= forcemakerid %>" size="16" maxlength="32">
        &nbsp;&nbsp;
        <% end if %>
		* 상품코드 : <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;&nbsp;
		* 상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="30" maxlength="32">
		<% if not(C_IS_SHOP) then %>
			&nbsp;&nbsp;
			* 상품사용여부 : <% drawSelectBoxUsingYN "isusing", isusing %>
		<% end if %>

		<% if (Not C_IS_SHOP) then %>
		&nbsp;&nbsp;
		ON매입구분 :
		<select class="select" name="onlinemwdiv">
			<option></option>
			<option value="M" <% if (onlinemwdiv = "M") then %>selected<% end if %> >매입</option>
			<option value="W" <% if (onlinemwdiv = "W") then %>selected<% end if %> >위탁</option>
			<option value="U" <% if (onlinemwdiv = "U") then %>selected<% end if %> >업체</option>
		</select>
		<% end if %>

		&nbsp;&nbsp;
		<input type="checkbox" name="imageon" value="on" <% if imageon="on" then response.write "checked" %> > 이미지표시
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    	<input type=checkbox name="ipgo" <% if ipgo = "on" then response.write " checked" %>>입고된것만(매장)
		<input type=checkbox name="logicsipgo" <% if logicsipgo = "on" then response.write " checked" %>>입고된것만(물류)
        <input type=checkbox name="sell7days" <% if sell7days = "on" then response.write " checked" %>>최근7일판매내역있는것만
        <input type=checkbox name="includepreorder" <% if includepreorder = "on" then response.write " checked" %>>기주문포함부족만&nbsp;
        재고부족 : <input type="radio" name="shortagetype" value="" <% if shortagetype="" then response.write " checked" %>>전체&nbsp;
        <input type="radio" name="shortagetype" value="3" <% if shortagetype="3" then response.write " checked" %>>3일후&nbsp;
        <input type="radio" name="shortagetype" value="7" <% if shortagetype="7" then response.write " checked" %>>7일후&nbsp;
        <input type="radio" name="shortagetype" value="14" <% if shortagetype="14" then response.write " checked" %>>14일후&nbsp;
		&nbsp;
		정렬순서 :
		<select class="select" name="ordby">
			<option value="BI" <% if (ordby = "BI") then %>selected<% end if %> >브랜드</option>
			<option value="I" <% if (ordby = "I") then %>selected<% end if %> >상품코드 역순</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="radio" name="mode" value="all" <% if mode="all" then response.write "checked" %> >전체
		<input type="radio" name="mode" value="by7sell" <% if mode="by7sell" then response.write "checked" %> >이전7일 판매상품
		<input type="radio" name="mode" value="byevent" <% if mode="byevent" then response.write "checked" %> disabled ><font color=gray>텐바이텐 기획상품[준비중]</font>
		<input type="radio" name="mode" value="byrecent" <% if mode="byrecent" then response.write "checked" %> >신상품
		<input type="radio" name="mode" value="byshopfav" <% if mode="byshopfav" then response.write "checked" %> disabled ><font color=gray>관심상품[준비중]</font>
		<input type="radio" name="mode" value="byetc" <% if mode="byetc" then response.write "checked" %> >기타소모품 <!-- 70 -->
		<br>
		<input type="radio" name="mode" value="bybrand" <% if mode="bybrand" then response.write "checked" %> >브랜드별
		<input type="radio" name="mode" value="byonbest" <% if mode="byonbest" then response.write "checked" %> >온라인 베스트
		<!-- <input type="radio" name="mode" value="byonfav" <% if mode="byonfav" then response.write "checked" %> >온라인 인기상품 -->
		<input type="radio" name="mode" value="byoffbest" <% if mode="byoffbest" then response.write "checked" %> >오프라인 베스트
		<input type="radio" name="mode" value="byoffbestAll" <% if mode="byoffbestAll" then response.write "checked" %> >오프라인 베스트(전샵)
		&nbsp;&nbsp;
		<input type="checkbox" name="onlyActive" <% if onlyActive="on" then response.write "checked" %>> 온라인 및 오프 사용중인 상품만보기
	</td>
</tr>
</form>
</table>
<p>
<!-- 액션 시작 -->
<% if ioffitem.FresultCount>0 then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="선택 상품 추가" onclick="AddArr()">
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- 액션 끝 -->
<p>
<% end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ioffitem.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		검색결과 : <b><%= ioffitem.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= Page %> / <%= ioffitem.FTotalPage %></b>
	</td>
</tr>
<% end if %>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<% if imageon="on" then %>
	<td width="50">이미지</td>
	<% end if %>
	<td>브랜드</td>
	<td>상품코드</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td width="60">판매가</td>

	<% if (Not C_IS_Maker_Upche) then %>
	<td width="60">출고가</td>
	<% end if %>
	<% if (Not C_IS_SHOP) then %>
	<td width="60">매입가</td>
	<% end if %>

	<% if (Not C_IS_SHOP) then %>
	<td width="40">ON<br>매입<br>구분</td>
	<td width="40">센터<br>매입<br>구분</td>
	<% end if %>

	<td width="40">수량</td>
	<td width="70">비고</td>
</tr>
<% for i=0 to ioffitem.FResultCount -1 %>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemList(i).Fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">
<input type="hidden" name="itemname" value="<%= ioffitem.FItemList(i).FShopItemName %>">
<input type="hidden" name="itemoptionname" value="<%= ioffitem.FItemList(i).FShopItemOptionName %>">
<input type="hidden" name="desingerid" value="<%= ioffitem.FItemList(i).FMakerid %>">
<% if Not (PriceEditEnable) then %>
<input type="hidden" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>">
<% if IS_HIDE_BUYCASH = True then %>
<input type="hidden" name="suplycash" value="-1">
<% else %>
<input type="hidden" name="suplycash" value="<%= ioffitem.FItemList(i).GetOfflineBuycash %>">
<% end if %>
<input type="hidden" name="shopbuyprice" value="<%= ioffitem.FItemList(i).GetOfflineSuplycash %>">
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="2"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<% if imageon="on" then %>
	<td rowspan="2"><img src="<%= ioffitem.FItemList(i).FimageSmall %>" width="50" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
	<% end if %>
	<td height="25"><%= ioffitem.FItemList(i).Fmakerid %></td>
	<td>
		<!--
		<a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= ioffitem.FItemList(i).Fitemgubun %>&itemid=<%= ioffitem.FItemList(i).Fshopitemid %>&itemoption=<%= ioffitem.FItemList(i).Fitemoption %>" target=_blank ><%= ioffitem.FItemList(i).GetBarCode %></a>
		-->
		<a href="/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=<%= shopid %>&barcode=<%= ioffitem.FItemList(i).GetBarCode %>" target=_blank ><%= ioffitem.FItemList(i).GetBarCode %></a>
	</td>
	<td align="left"><%= ioffitem.FItemList(i).FShopItemName %></td>
	<td align="left"><%= ioffitem.FItemList(i).FShopItemOptionName %></td>
	<% if Not (PriceEditEnable) then %>
		<td rowspan="2" align="right"><%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice,0) %></td>
		<% if (Not C_IS_Maker_Upche) then %>
		<td rowspan="2" align="right">
			<%= FormatNumber(ioffitem.FItemList(i).GetOfflineSuplycash,0) %>
			<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
			<br>(<%= 100-Clng(ioffitem.FItemList(i).GetOfflineSuplycash/ioffitem.FItemList(i).Fshopitemprice*100*100)/100 %> %)
			<% end if %>
		</td>
		<% end if %>
		<% if (Not C_IS_SHOP) then %>
		<td rowspan="2" align="right"><%= FormatNumber(ioffitem.FItemList(i).GetOfflineBuycash,0) %></td>
		<% end if %>
	<% else %>
	<td rowspan="2"><input type="text" class="text" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>" size="5" maxlength="9"></td>
	<td rowspan="2"><input type="text" class="text" name="shopbuyprice" value="<%= ioffitem.FItemList(i).GetOfflineSuplycash %>" size="5" maxlength="9"></td>
	<td rowspan="2"><input type="text" class="text" name="suplycash" value="<%= ioffitem.FItemList(i).GetOfflineBuycash %>" size="5" maxlength="9"></td>
	<% end if %>

	<% if (Not C_IS_SHOP) then %>
	<td rowspan="2"><%= ioffitem.FItemList(i).Fmwdiv %></td>
	<td rowspan="2"><%= ioffitem.FItemList(i).Fcentermwdiv %></td>
	<% end if %>

	<td rowspan="2"><input type="text" class="text" name="itemno" value="<%= ioffitem.FItemList(i).Fitemno %>" size="3" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);"></td>
	<td rowspan="2">
        <% if ioffitem.FItemList(i).Fpreorderno>0 or ioffitem.FItemList(i).Fpreorderno<0 then %>
        	기주문:
            <% if ioffitem.FItemList(i).Fpreorderno<>ioffitem.FItemList(i).Fpreordernofix then response.write CStr(ioffitem.FItemList(i).Fpreorderno) + " -> " %>
        	<%= ioffitem.FItemList(i).Fpreordernofix %><br>
        <% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="4" >
		<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>센터<br>입고</td>
				<td>업체<br>입고</td>
				<td>판매</td>
				<td>시스템<br>재고</td>
				<td>오차</td>
				<td>실사<br>재고</td>
				<td>샘플</td>
				<td>유효<br>재고</td>
				<td>OFF<br>3일</td>
				<td>OFF<br>7일</td>
				<td>3일</td>
				<td>7일</td>
				<td>14일</td>
			</tr>
			<tr align="center" bgcolor="#FFFFFF" height="25">
				<td>
					<%= ioffitem.FItemlist(i).flogicsipgono + ioffitem.FItemlist(i).flogicsreipgono %>    <!--센터입고반품-->
				</td>
				<td>
					<%= ioffitem.FItemlist(i).fbrandipgono + ioffitem.FItemlist(i).fbrandreipgono %>		<!--브랜드입고반품-->
				</td>
				<td>
					<%= ioffitem.FItemlist(i).fsellno+ioffitem.FItemlist(i).fresellno %>       <!--총판매현황 -->
				</td>
				<td bgcolor="#EEEEFF">
					<b><%= ioffitem.FItemlist(i).fsysstockno %></b>       <!--시스템재고-->
				</td>
				<td>
					<%= ioffitem.FItemlist(i).Ferrrealcheckno %>       <!--오차-->
				</td>
				<td bgcolor="#EEEEFF">
					<b><%= ioffitem.FItemlist(i).frealstockno %></b>          <!-- 실사재고 -->
				</td>
				<td>
					<%= ioffitem.FItemlist(i).ferrsampleitemno %>      <!--샘플-->
				</td>
				<td bgcolor="#EEEEFF">
					<b><%= ioffitem.FItemlist(i).getAvailStock %></b>     <!--유효재고-->
				</td>

				<td><%= ioffitem.FItemlist(i).fsell3days %></td>		<!--판매수량-->
				<td><%= ioffitem.FItemlist(i).fsell7days %></td>

				<td>													<!-- 총필요수량 -->
					<% if ioffitem.FItemlist(i).frequire3daystock > 0 then %>
					<font color="red"><%= ioffitem.FItemlist(i).frequire3daystock*-1 %></font>
					<% else %>
					0
					<% end if %>
				</td>
				<td>
					<% if ioffitem.FItemlist(i).frequire7daystock > 0 then %>
					<font color="red"><%= ioffitem.FItemlist(i).frequire7daystock*-1 %></font>
					<% else %>
					0
					<% end if %>
				</td>
				<td>
					<% if ioffitem.FItemlist(i).frequire14daystock > 0 then %>
					<font color="red"><%= ioffitem.FItemlist(i).frequire14daystock*-1 %></font>
					<% else %>
					0
					<% end if %>
				</td>
			</tr>
		</table>
	</td>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="#FFFFFF">
	<td colspan="30" align="center">
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
			<a href="javascript:ReSearch('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if ioffitem.HasNextScroll then %>
			<a href="javascript:ReSearch('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<form name="frmArrupdate" method="post" action="">
	<input type="hidden" name="mode" value="arrins">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="itemnamearr" value="">
	<input type="hidden" name="itemoptionnamearr" value="">
	<input type="hidden" name="designerarr" value="">
</form>

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
