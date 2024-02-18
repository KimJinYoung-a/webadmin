<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  매장/관리자 공통 사용
''				[OFF]오프_물류관리>>오프라인주문서작성/수정시
''				[OFF]오프_물류관리>>업체개별주문서관리 > 수정시
''				매장-입/출/재고관리>>주문서작성  / 주문서관리
''				재고 작성 후 TRUE
' History : 2009.04.07 서동석 생성
'			2010.06.01 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%

dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
if C_IS_OWN_SHOP or C_IS_SHOP then
	IS_HIDE_BUYCASH = True
end if

'' =============================================================================
'' 아래 3개의 메뉴 검색조건은 기본적으로 동일해야 한다.
'' (매장 적정재고 부족상품, 주문관리(물류), 주문관리(업 체))
'' =============================================================================
'' /common/offshop/stock/shortagestock_shop.asp
'' /common/offshop/popshopitem2.asp
'' /common/offshop/popshopjumunitem.asp
'' =============================================================================

CONST C_STOCK_DISP = True

dim menupos, page, mode, designer,suplyer,shopid,itemid,itemname , idx , onlyActive, research
dim i, shopsuplycash, buycash,prdcode, generalbarcode ,cdl, cdm, cds, cwflag ,comm_cd, isusing
Dim currencyunit,loginsite, countrylangcd, sellyn
dim ipgo, logicsipgo, sell7days, includepreorder, shortagetype, onlinemwdiv, ordby
dim cp_idx, foreign_suplycash
dim forcemakerid
	sellyn      = requestCheckvar(request("sellyn"),10)
	shopid  = RequestCheckVar(request("shopid"),32)
	page    = RequestCheckVar(request("page"),9)
	mode    = RequestCheckVar(request("mode"),32)
	designer = RequestCheckVar(request("designer"),32)
	suplyer  = RequestCheckVar(request("suplyer"),32)
	itemid  = request("itemid")
	itemname= RequestCheckVar(request("itemname"),32)
	idx     = RequestCheckVar(request("idx"),32)
	onlyActive = RequestCheckVar(request("onlyActive"),32)
	research = RequestCheckVar(request("research"),32)
	prdcode = RequestCheckVar(request("prdcode"),32)
	generalbarcode = RequestCheckVar(request("generalbarcode"),32)
	cdl = requestCheckVar(request("cdl"),3)
	cdm = requestCheckVar(request("cdm"),3)
	cds = requestCheckVar(request("cds"),3)
	cwflag = requestCheckVar(request("cwflag"),10)
	isusing = requestCheckVar(request("isusing"),1)

	ipgo = RequestCheckVar(request("ipgo"),32)
    logicsipgo = RequestCheckVar(request("logicsipgo"),32)
	sell7days = RequestCheckVar(request("sell7days"),32)
	includepreorder = RequestCheckVar(request("includepreorder"),32)
	shortagetype = RequestCheckVar(request("shortagetype"),32)

	onlinemwdiv = RequestCheckVar(request("onlinemwdiv"),1)
	ordby = RequestCheckVar(request("ordby"),10)
	cp_idx = RequestCheckVar(request("cp_idx"),20)
    forcemakerid = RequestCheckVar(request("forcemakerid"),32)

''suplyer = "10x10" 인경우 매장에서 센터로 주문시 ELSE 센터에서 브랜드로 주문시.
if suplyer<>"10x10" then designer = suplyer
if page="" then page=1

if research = "" And cp_idx = "" then
	''ipgo = "on"
    logicsipgo = "on"
	includepreorder = "on"
	shortagetype = "7"
	mode = "all"
end if

if (research<>"on") and (sellyn="") And cp_idx = "" then
    sellyn = "YS"
end if
if research="" and isusing="" then isusing="Y"
'if (research="") then onlyActive="on"

if C_ADMIN_USER then

elseif (C_IS_SHOP) then
    shopid = C_STREETSHOPID
    isusing="Y"
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

if (prdcode<>"") then
	if Not(isNumeric(prdcode)) then
		Response.Write "<script type='text/javascript'>alert('"& CTX_Type_Mismatch &" ["& CTX_Warehouse & CTX_Barcode &":" & prdcode & "]'); history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

if cwflag = "1" then
	comm_cd = "'B013'"
else
	comm_cd = "'B031','B011'"
end if

if (research<>"on") and (ordby="") then
    ordby = "BI"
end if

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 50
	ioffitem.FCurrPage = page
    if forcemakerid = "" then
	    ioffitem.FRectDesigner = designer
    else
        ioffitem.FRectDesigner = forcemakerid
    end if
	ioffitem.FRectAdminView = "on"
	ioffitem.FRectOnlyActive = onlyActive
	ioffitem.FRectOrder = mode
	ioffitem.FRectItemid = itemid
	ioffitem.FRectItemName = itemname
	ioffitem.FRectPrdCode = prdcode
	ioffitem.FRectGeneralBarcode = generalbarcode
	ioffitem.FRectCDL   	= cdl
	ioffitem.FRectCDM     	= cdm
	ioffitem.FRectCDS   	= cds
	ioffitem.FRectcomm_cd = comm_cd
	ioffitem.frectisusing = isusing
	ioffitem.FRectIpGoOnly = ipgo
    ioffitem.FRectLogicsIpGoOnly = logicsipgo
	ioffitem.FRectSell7days = sell7days
	ioffitem.FRectIncludePreOrder = includepreorder
	ioffitem.FRectShortageType = shortagetype
	ioffitem.FRectOnlineMWdiv = onlinemwdiv
	ioffitem.FRectSellYN       = sellyn

	if (suplyer="10x10") then
	    ''매장에서 센터로 주문시. (업체위탁, 업체매입 제외)
		ioffitem.FRectShopid = shopid
		ioffitem.FRectDesignerjungsangubun = "'2','4','5'"
		''cp_idx
		If (cp_idx <> "") Then
			ioffitem.FRectCopyIdx = cp_idx
			ioffitem.FPageSize = 250
		End If
	else
	    ''센터에서 브랜드로 주문시.(업체 개별 매입/위탁 주문서)
		if (shopid="10x10") or (shopid="") then
			ioffitem.FRectDesigner = suplyer
			ioffitem.FRectShopid = "streetshop800" '-->기본 800
			''''XX ioffitem.FRectDesignerjungsangubun = "'2','4'"
		else
			ioffitem.FRectShopid = shopid
			''''XX ioffitem.FRectDesignerjungsangubun = "'6','8'"
		end if
	end if

	ioffitem.FRectOrder = ordby

	if (itemid<>"") or (prdcode<>"") or (generalbarcode<>"") or (itemname<>"") or (designer <> "") or (ipgo = "on") or (mode="by7sell") or (mode="byrecent") or (mode="byetc") or (mode="byonbest") or (mode="byoffbest") or (mode="byoffbestAll") Or (cp_idx <> "") then
	    ioffitem.GetOffLineJumunItemWithStock

		if shopid <> "" then
			countrylangcd = ioffitem.fcountrylangcd
			currencyunit = ioffitem.Fcurrencyunit
			loginsite = ioffitem.Floginsite
		end if
	end if

'// ==============================================
dim IsShopChulgo : IsShopChulgo = False
if (suplyer = "10x10") then
    IsShopChulgo = True
end if

%>

<script type='text/javascript'>

function enablebrand(bool){
	//document.frm.designer.disabled = bool;
}

function search(frm){
	if ((frm.itemid.value.length<1)&&(frm.generalbarcode.value.length<1)&&(frm.mode[5].checked)&&(frm.designer.value.length<1)){
		alert('<%= CTX_Please_select %> (<%= CTX_Brand %>)');
		frm.designer.focus();
		return;
	}

	frm.submit();
}

function AddArr(){
	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;
	var unreg="";
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('<%= CTX_Please_select %> (ITEM)');
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
	upfrm.foreign_sellcasharr.value = "";
	upfrm.foreign_suplycasharr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsInteger(frm.itemno.value)){
					alert('<%= CTX_Type_Mismatch %> (<%= CTX_Only_numbers %>)');
					frm.itemno.focus();
					return;
				}

				if (frm.itemno.value=="0"){
					alert('<%= CTX_Type_Mismatch %> (0 <%= CTX_except %>)');
					frm.itemno.focus();
					return;
				}

				if(frm.foreign_sellcash.value==0&&(document.frm.loginsite.value=="WSLWEB")){
					if ( unreg == "" ){
						unreg = frm.itemid.value;
					}else{
				 		unreg = unreg + "," + frm.itemid.value;
					}

					//미등록 상품도 일단 넣는다. 차후에 해외상품단에 상품구분이 생기면 입력안되게 막던가 해야함.	//2017.06.12 한용민
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
					upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
					upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
					upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
					upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
					upfrm.foreign_sellcasharr.value = upfrm.foreign_sellcasharr.value+frm.foreign_sellcash.value + "|";
					upfrm.foreign_suplycasharr.value = upfrm.foreign_suplycasharr.value+frm.foreign_suplycash.value + "|";
					upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
					upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
					upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
					upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
				}else{
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
					upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
					upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
					upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
					upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
					upfrm.foreign_sellcasharr.value = upfrm.foreign_sellcasharr.value+frm.foreign_sellcash.value + "|";
					upfrm.foreign_suplycasharr.value = upfrm.foreign_suplycasharr.value+frm.foreign_suplycash.value + "|";
					upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
					upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
					upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
					upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
				}
			}
		}
	}

	if (unreg!=""){
		alert("선택하신 상품 중 상품코드 ["+unreg+"]는 미등록상품입니다. \n상품 등록 후 주문해주세요");
	}

	opener.ReActItems('<%= idx %>',upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value,upfrm.foreign_sellcasharr.value,upfrm.foreign_suplycasharr.value);
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="suplyer" value="<%= suplyer %>" >
<input type="hidden" name="shopid" value="<%= shopid %>" >
<input type="hidden" name="idx" value="<%= idx %>" >
<input type="hidden" name="cp_idx" value="<%= cp_idx %>" >
<input type="hidden" name="page" value="1" >
<input type="hidden" name="cwflag" value="<%= cwflag %>">
<input type="hidden" name="loginsite" value="<%=loginsite%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>"><%=CTX_SEARCH%><br><%= CTX_conditional %></td>
	<td align="left">
		* [<%= CTX_SHOP %> : <%= shopid %>]
		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			&nbsp;&nbsp;
			* 상품사용여부 : <% drawSelectBoxisusingYN "isusing", isusing, " onchange='search(frm)';" %>
		<% end if %>

		&nbsp;&nbsp;
		ON매입구분 :
		<select class="select" name="onlinemwdiv">
			<option></option>
			<option value="M" <% if (onlinemwdiv = "M") then %>selected<% end if %> >매입</option>
			<option value="W" <% if (onlinemwdiv = "W") then %>selected<% end if %> >위탁</option>
			<option value="U" <% if (onlinemwdiv = "U") then %>selected<% end if %> >업체</option>
		</select>
	 	&nbsp;&nbsp;온라인판매여부:<% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:search(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* <%= CTX_Brand %> : <% drawSelectBoxShopjumunDesignerNotUpche "designer",designer,shopid,suplyer,comm_cd %>
		&nbsp;&nbsp;
		* <%= CTX_Description %> :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" onKeyPress="if (event.keyCode == 13) search(frm);">
		&nbsp;&nbsp;
		<!--
		* <%= CTX_Warehouse %>&nbsp;<%= CTX_Barcode %> :
		<input type="text" class="text" name="prdcode" value="<%= prdcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) search(frm);">
		-->
		&nbsp;&nbsp;
		* <%= CTX_Universal %>&nbsp;<%= CTX_Barcode %> :
		<input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) search(frm);">
		<br>
        <% if C_ADMIN_AUTH then %>
        * 브랜드[관리자] <input type="text" class="text" name="forcemakerid" value="<%= forcemakerid %>" size="16" maxlength="32">
        &nbsp;&nbsp;
        <% end if %>
		* <%= CTX_Item_Code %> : <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;&nbsp;
		* 정산기준 :
		<% if (comm_cd = "'B013'") then %>
			출고위탁
		<% else %>
			출고매입+텐바이텐위탁
		<% end if %>
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
		<input type="radio" name="mode" value="by7sell" <% if mode="by7sell" then response.write "checked" %> >7<%= CTX_days %>&nbsp;<%= CTX_Selling %>
		<input type="radio" name="mode" value="byevent" <% if mode="byevent" then response.write "checked" %> disabled ><font color=gray>텐바이텐 기획상품[준비중]</font>
		<input type="radio" name="mode" value="byrecent" <% if mode="byrecent" then response.write "checked" %> ><%= CTX_NEW %>&nbsp;<%= CTX_ITEM %>
		<input type="radio" name="mode" value="byshopfav" <% if mode="byshopfav" then response.write "checked" %> disabled ><font color=gray>관심상품[준비중]</font>
		<input type="radio" name="mode" value="byetc" <% if mode="byetc" then response.write "checked" %> ><%= CTX_consumables %> <!-- 70 -->
		<p>
		<input type="radio" name="mode" value="bybrand" <% if mode="bybrand" then response.write "checked" %> ><%= CTX_Brand %>
		<input type="radio" name="mode" value="byonbest" <% if mode="byonbest" then response.write "checked" %> ><%= CTX_ONLINE %>&nbsp;<%= CTX_BEST %>
		<!-- <input type="radio" name="mode" value="byonfav" <% if mode="byonfav" then response.write "checked" %> >온라인 인기상품 -->
		<input type="radio" name="mode" value="byoffbest" <% if mode="byoffbest" then response.write "checked" %> ><%= CTX_OFFLINE %>&nbsp;<%= CTX_BEST %>
		<input type="radio" name="mode" value="byoffbestAll" <% if mode="byoffbestAll" then response.write "checked" %> ><%= CTX_OFFLINE %>&nbsp;<%= CTX_BEST %>(ALL SHOP)
		&nbsp;&nbsp;
		<input type="checkbox" name="onlyActive" <% if onlyActive="on" then response.write "checked" %>> <%=CTX_ONLINE%>&nbsp;<%= CTX_use_Y %>
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
		* 판매량은 오프라인 판매량만 표시되며, 판매량/현재고 수량은 매일 새벽 3시 기준으로 표시됩니다.
		<br>
		* 업체위탁및 매장매입 주문은 [입/출/재고관리>>개별입출고리스트] 메뉴를 사용하세요.
	</td>
	<td align="right">
		<% if ioffitem.FresultCount>0 then %>
			<input type="button" class="button" value="<%= CTX_Add_new_items %>" onclick="AddArr()">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="#FFFFFF">
				<td>
					검색결과 : <b><%= ioffitem.FTotalCount %></b>
					&nbsp;
					페이지 : <b><%= Page %> / <%= ioffitem.FTotalPage %></b>
				</td>
				<td align="right">
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" onClick="AnSelectAllFrame(this.checked)"></td>
	<td width="50">이미지</td>
	<td>브랜드</td>
	<td width="90">상품코드</td>
	<td>상품명</td>
	<td>옵션</td>

	<% if Not (C_STOCK_DISP) then %>
		<% if mode="byoffbestAll" then %>
			<td width="50">5<%= CTX_days %><br><%= CTX_Selling %></td> <!-- 5일이 맞다. -->
			<td width="50">--</td>
		<% else %>
			<td width="50">7<%= CTX_days %><br><%= CTX_Selling %></td>
			<td width="50">3<%= CTX_days %><br><%= CTX_Selling %></td>
		<% end if %>

		<td width="70"><%= CTX_Now %><br>SHOP&nbsp;<%= CTX_stock %></td>
	<% end if %>

	<td width="50">판매가</td>
	<td width="50">출고가</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td width="50">매입가</td>
		<td width="40">ON<br>매입<br>구분</td>
		<td width="40">센터<br>매입<br>구분</td>
	<% end if %>

	<td width="50">수량</td>
	<td width="70">비고</td>
</tr>
<% for i=0 to ioffitem.FResultCount -1 %>
<%
    ''동도 트레이딩.. 마진 10%..
    if shopid="streetshop881" then
        ioffitem.FItemList(i).Fshopbuyprice = 0
    end if
	shopsuplycash = ioffitem.FItemList(i).GetFranchiseSuplycash
	buycash		  = ioffitem.FItemList(i).GetFranchiseBuycash
    if IsShopChulgo then
        buycash		  = ioffitem.FItemList(i).GetFranchiseBuycashByItemInfo			'// 수익률 분석용 상품정보 매입가 (정산내역에 올라가면 안됨)
    end if

	if ioffitem.Floginsite="WSLWEB" then
		'/ 해외 출고가. 쿼리단에서 상품테이블에 출고가가 없을경우 복잡해서 처리 못한거 넣어줌
		if ioffitem.FItemList(i).Fforeign_suplyprice="" or isnull(ioffitem.FItemList(i).Fforeign_suplyprice) or ioffitem.FItemList(i).Fforeign_suplyprice=0 then
			foreign_suplycash = shopsuplycash
		else
			foreign_suplycash = ioffitem.FItemList(i).Fforeign_suplyprice
		end if
	end if
%>
<form name="frmBuyPrc_<%= i %>" style="margin:0px;">
<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemList(i).Fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">
<input type="hidden" name="itemname" value="<%= ioffitem.FItemList(i).FShopItemName %>">
<input type="hidden" name="itemoptionname" value="<%= ioffitem.FItemList(i).FShopItemOptionName %>">
<input type="hidden" name="desingerid" value="<%= ioffitem.FItemList(i).FMakerid %>">
<input type="hidden" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>">
<input type="hidden" name="suplycash" value="<%= shopsuplycash %>">
<% if IS_HIDE_BUYCASH then %>
<input type="hidden" name="buycash" value="-1">
<% else %>
<input type="hidden" name="buycash" value="<%= buycash %>">
<% end if %>
<input type="hidden" name="foreign_sellcash" value="<%= getdisp_price(ioffitem.FItemList(i).Fforeign_sellprice, ioffitem.fcurrencyChar) %>">
<input type="hidden" name="foreign_suplycash" value="<%= getdisp_price(foreign_suplycash, ioffitem.fcurrencyChar) %>">
<tr align="center" bgcolor="#FFFFFF">
	<td <%= ChkIIF(C_STOCK_DISP,"rowspan=2","") %> ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td <%= ChkIIF(C_STOCK_DISP,"rowspan=2","") %> ><img src="<%= ioffitem.FItemList(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
	<td height="25"><%= ioffitem.FItemList(i).FMakerid %></td>
	<td>
		<!--
		<a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= ioffitem.FItemList(i).Fitemgubun %>&itemid=<%= ioffitem.FItemList(i).Fshopitemid %>&itemoption=<%= ioffitem.FItemList(i).Fitemoption %>" target=_blank ><%= ioffitem.FItemList(i).GetBarCode %></a>
		-->
		<a href="/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=<%= shopid %>&barcode=<%= ioffitem.FItemList(i).GetBarCode %>" target=_blank ><%= ioffitem.FItemList(i).GetBarCode %></a>
	</td>
	<td align="left">
		<%= ioffitem.FItemList(i).FShopItemName %>
    </td>
	<td align="left">
		<% if right(ioffitem.FItemList(i).GetBarCode,4)<>"0000" then %>
			<%= ioffitem.FItemList(i).FShopItemOptionName %>
		<% end if %>
    </td>

    <% if Not (C_STOCK_DISP) then %>
		<td><%= ioffitem.FItemList(i).FOffsell7days %></td>
		<td><%= ioffitem.FItemList(i).Fsell3days %></td>
		<td><%= ioffitem.FItemList(i).Frealstockno %></td>
	<% end if %>

	<td <%= ChkIIF(C_STOCK_DISP,"rowspan=2","") %> align="right">
		<%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice,0) %>

		<br><font color="Gray"><%= getdisp_price_currencyChar(ioffitem.FItemList(i).Fforeign_sellprice, ioffitem.fcurrencyChar) %></font>
	</td>
	<td <%= ChkIIF(C_STOCK_DISP,"rowspan=2","") %> align="right"><%= FormatNumber(shopsuplycash,0) %><br>
		<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
			(<%= 100-(CLng(shopsuplycash/ioffitem.FItemList(i).Fshopitemprice*10000)/100) %> %)
		<% end if %>

		<br><font color="gray"><%= getdisp_price_currencyChar(foreign_suplycash, ioffitem.fcurrencyChar) %></font>
	</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td <%= ChkIIF(C_STOCK_DISP,"rowspan=2","") %> align="right"><%= FormatNumber(buycash,0) %></td>
		<td <%= ChkIIF(C_STOCK_DISP,"rowspan=2","") %> ><%= ioffitem.FItemList(i).Fmwdiv %></td>
		<td <%= ChkIIF(C_STOCK_DISP,"rowspan=2","") %> ><%= ioffitem.FItemList(i).Fcentermwdiv %></td>
	<% end if %>

	<td <%= ChkIIF(C_STOCK_DISP,"rowspan=2","") %> ><input type="text" class="text" name="itemno" value="<%= ioffitem.FItemList(i).Fitemno %>" size="2" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);"></td>
	<td <%= ChkIIF(C_STOCK_DISP,"rowspan=2","") %> >
        <% if ioffitem.FItemList(i).Fpreorderno>0 or ioffitem.FItemList(i).Fpreorderno<0 then %>
        	기주문:
            <% if ioffitem.FItemList(i).Fpreorderno<>ioffitem.FItemList(i).Fpreordernofix then response.write CStr(ioffitem.FItemList(i).Fpreorderno) + " -> " %>
        	<%= ioffitem.FItemList(i).Fpreordernofix %><br>
        <% end if %>
		<% if ioffitem.FItemList(i).IsSoldOut then %>
			<font color="red"><%= CTX_sold_out %></font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).Flimityn="Y" then %>
			<font color="blue"><%= CTX_Limit %>(<%= ioffitem.FItemList(i).getLimitNo %>)</font><br />
		<% end if %>
		물류(<%= ioffitem.FItemList(i).FLogicsRealStock %>)
	</td>
</tr>

<% if (C_STOCK_DISP) then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="4">
			<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td><%= CTX_Center %><br><%= CTX_Warehousing %></td>
					<td><%= CTX_Upche %><br><%= CTX_Warehousing %></td>
					<td><%= CTX_Sell %></td>
					<td><%= CTX_system %><br><%= CTX_stock %></td>
					<td><%= CTX_error %></td>
					<td><%= CTX_Inspection %><br><%= CTX_stock %></td>
					<td><%= CTX_Sample %></td>
					<td><%= CTX_Available %><br><%= CTX_stock %></td>
					<td>OFF<br>3<%= CTX_days %></td>
					<td>OFF<br>7<%= CTX_days %></td>
					<td>3<%= CTX_days %></td>
					<td>7<%= CTX_days %></td>
					<td>14<%= CTX_days %></td>
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
<% end if %>

</form>

<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="13" align="center">
	<% if ioffitem.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ioffitem.StartScrollPage-1 %>');">[pre]</a>
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
</tr>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<% if ioffitem.FresultCount>0 then %>
			<input type="button" class="button" value="<%= CTX_Add_new_items %>" onclick="AddArr()">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frmArrupdate" method="post" action="">
	<input type="hidden" name="mode" value="arrins">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="buycasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="foreign_sellcasharr" value="">
	<input type="hidden" name="foreign_suplycasharr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="itemnamearr" value="">
	<input type="hidden" name="itemoptionnamearr" value="">
	<input type="hidden" name="designerarr" value="">
</form>

<%
set ioffitem = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
