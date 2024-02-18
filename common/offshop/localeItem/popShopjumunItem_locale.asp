<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  매장/관리자 공통 사용
''				[OFF]오프_물류관리>>오프라인주문서작성/수정시
''				[OFF]오프_물류관리>>업체개별주문서관리 > 수정시
''				매장-입/출/재고관리>>주문서작성  / 주문서관리
''				재고 작성 후 TRUE
' History : 2010.08.12 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%

dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
if C_IS_OWN_SHOP or C_IS_SHOP then
	IS_HIDE_BUYCASH = True
end if

CONST C_STOCK_DISP = FALSE

dim menupos, page, mode, designer,suplyer,shopid,itemid,itemname , idx , onlyActive, research
dim i, shopsuplycash, buycash , makertr , cdl, cdm, cds , l,m,s , cdlname , cdmname, isusing
dim oLcate ,oMcate ,oScate, cwflag ,comm_cd
dim currencyunit, loginsite, foreign_suplycash
	isusing = request("isusing")
	shopid  = RequestCheckVar(request("shopid"),32)
	page    = RequestCheckVar(request("page"),9)
	mode    = RequestCheckVar(request("mode"),32)
	designer = RequestCheckVar(request("designer"),32)
	suplyer  = RequestCheckVar(request("suplyer"),32)
	itemid  = RequestCheckVar(request("itemid"),32)
	itemname= RequestCheckVar(request("itemname"),32)
	idx     = RequestCheckVar(request("idx"),32)
	onlyActive = RequestCheckVar(request("onlyActive"),32)
	research = RequestCheckVar(request("research"),32)
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
	cwflag = request("cwflag")

''suplyer = "10x10" 인경우 매장에서 센터로 주문시 ELSE 센터에서 브랜드로 주문시.
if suplyer<>"10x10" then designer = suplyer
if page="" then page=1
if mode="" then mode="bybrand"
if (research="") then onlyActive="on"
if research="" and isusing="" then isusing="Y"

if C_ADMIN_USER then

elseif (C_IS_SHOP) then
    shopid = C_STREETSHOPID
    isusing="Y"
end if

'상품코드 유효성 검사(2008.08.01;허진원)
if (itemid<>"") then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

if cwflag = "1" then
	comm_cd = "'B013'"
else
	comm_cd = "'B031','B011'"
end if

'상품추가 : 해외주문일 경우 해외등록상품인지 체크
dim sqlStr
if shopid <> "" then
	sqlStr = "select currencyunit, loginsite from db_shop.dbo.tbl_shop_user where userid = '" + shopid + "' "
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		currencyunit = rsget("currencyunit")
		loginsite = rsget("loginsite")
	end if
	rsget.Close
end if

'//상품리스트
dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 20
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = designer
	ioffitem.FRectAdminView = "on"
	ioffitem.FRectOnlyActive = onlyActive
	ioffitem.FRectOrder = mode
	ioffitem.FRectItemid = itemid
	ioffitem.FRectItemName = itemname
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.FRectcomm_cd = comm_cd
	ioffitem.frectisusing = isusing

	if (suplyer="10x10") then
	    ''매장에서 센터로 주문시. (업체특정, 업체매입 제외)
		ioffitem.FRectShopid = shopid
		ioffitem.FRectDesignerjungsangubun = "'2','4','5'"
	else
	    ''센터에서 브랜드로 주문시.(업체 개별 매입/특정 주문서)
		if (shopid="10x10") or (shopid="") then
			ioffitem.FRectDesigner = suplyer
			ioffitem.FRectShopid = "streetshop800" '-->기본 800
			''''XX ioffitem.FRectDesignerjungsangubun = "'2','4'"
		else
			ioffitem.FRectShopid = shopid
			''''XX ioffitem.FRectDesignerjungsangubun = "'6','8'"
		end if
	end if

	if (itemid<>"") or (itemname<>"") or (designer<>"") or (mode="by7sell") or (mode="byrecent") or (mode="byetc") or (mode="byonbest") or (mode="byoffbest") or (mode="byoffbestAll") or (cdl<>"") or (cdm<>"cdm") or (cds<>"cds")then
	    ioffitem.GetOffLineJumunItemWithStock_locale
	end if

'//대카테고리 리스트
set oLcate = new CCatemanager
	oLcate.GetNewCateMaster()


'// ==============================================
dim IsShopChulgo : IsShopChulgo = False
if (suplyer = "10x10") then
    IsShopChulgo = True
end if

%>

<script language='javascript'>

function enablebrand(bool){
	//document.frm.designer.disabled = bool;
}

function search(frm){
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
		alert('선택 아이템이 없습니다.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.foreign_sellcasharr.value = "";
	upfrm.foreign_suplycasharr.value = "";
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
					upfrm.foreign_sellcasharr.value = upfrm.foreign_sellcasharr.value + frm.foreign_sellcash.value + "|";
					upfrm.foreign_suplycasharr.value = upfrm.foreign_suplycasharr.value + frm.foreign_suplycash.value + "|";
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
					upfrm.foreign_sellcasharr.value = upfrm.foreign_sellcasharr.value + frm.foreign_sellcash.value + "|";
					upfrm.foreign_suplycasharr.value = upfrm.foreign_suplycasharr.value + frm.foreign_suplycash.value + "|";
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

function OnOffLeftSubMenu(fnm,sw){
	var leftcate = document.all(fnm);
	if(sw=="on")
		leftcate.style.visibility = "visible";
	else
		leftcate.style.visibility = 'hidden';
}

function catesearch(cdl,cdm,cds){

	frm.cdl.value=cdl;
	frm.cdm.value=cdm;
	frm.cds.value=cds;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="suplyer" value="<%= suplyer %>" >
<input type="hidden" name="shopid" value="<%= shopid %>" >
<input type="hidden" name="idx" value="<%= idx %>" >
<input type="hidden" name="page" value="1" >
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="cwflag" value="<%= cwflag %>">
<input type="hidden" name="loginsite" value="<%=loginsite%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : <% drawSelectBoxShopjumunDesignerNotUpche "designer",designer,shopid,suplyer,comm_cd %>
		[주문매장 : <%= shopid %>]
		<p>
		* 상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=6 maxlength=7>
		* 상품명 :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;&nbsp;
		<% if not(C_IS_SHOP) then %>
			* 상품사용여부 : <% drawSelectBoxisusingYN "isusing", isusing, " onchange='search(frm)';" %>
		<% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:search(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" valign="top">
<%
'//카테고리 리스트 출력
if oLcate.FResultCount > 0 then

for l=0 to oLcate.FResultCount-1

set oMcate = new CCatemanager
	oMcate.GetNewCateMasterMid oLcate.FItemList(l).Fcdlarge
%>
	<td onmouseout="OnOffLeftSubMenu('cateory<%=oLcate.FItemList(l).FCdlarge%>','off')" onmouseover="OnOffLeftSubMenu('cateory<%=oLcate.FItemList(l).Fcdlarge%>','on')">
		<div id='cateory<%=oLcate.FItemList(l).Fcdlarge%>' style='position:absolute; width:70px; margin-top:15px; margin-left:0px;visibility:hidden;'>
		<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<%
		'//중카테고리 리스트 출력
		if oMcate.FResultCount > 0 then

		for m=0 to oMcate.FResultCount-1
		%>
		<tr align="center" bgcolor="#FFFFFF" valign="top">
			<td>
				<% if oLcate.FItemList(l).Fcdlarge = cdl and oMcate.FItemList(m).Fcdmid = cdm then %><b><% end if %>
				<a href="javascript:catesearch('<%=oLcate.FItemList(l).Fcdlarge%>','<%= oMcate.FItemList(m).Fcdmid %>','');" onfocus="this.blur();"><%= oMcate.FItemList(m).Fnmlarge %></a>
				<% if oLcate.FItemList(l).Fcdlarge = cdl and oMcate.FItemList(m).Fcdmid = cdm then %></b><% end if %>
			</td>
		</tr>
		<%
		if oLcate.FItemList(l).Fcdlarge = cdl and oMcate.FItemList(m).Fcdmid = cdm then
			cdmName = oMcate.FItemList(m).Fnmlarge
		end if
		%>
		<%
		next

		end if
		%>
		</table>
		</div>
		<% if oLcate.FItemList(l).Fcdlarge = cdl then %><b><% end if %>
			<a href="javascript:catesearch('<%=oLcate.FItemList(l).Fcdlarge%>','','');" onfocus="this.blur();"><%= oLcate.FItemList(l).Fnmlarge %></a>
		<% if oLcate.FItemList(l).Fcdlarge = cdl then %></b><% end if %>
		<%
		if oLcate.FItemList(l).Fcdlarge = cdl then
			cdlname = oLcate.FItemList(l).Fnmlarge
		end if
		%>
	</td>
<%
next

end if
%>
</tr>
<tr align="left" bgcolor="#FFFFFF">
	<td colspan=20>
		<%
		set oScate = new CCatemanager

		if cdl<>"" and cdm<>"" then
			oScate.GetNewCateMasterSmall cdl,cdm

		%>
			[ <%= cdlname %> . <%= cdmName %> ]
		<%
			'//소카테고리 리스트 출력
			if oScate.FResultCount > 0 then

			for s=0 to oScate.FResultCount - 1 %>
				<% if oScate.FItemList(s).Fcdsmall = cds then %><b><% end if %>
				<a href="javascript:catesearch('<%=cdl%>','<%= cdm %>','<%= oScate.FItemList(s).Fcdsmall %>');" onfocus="this.blur();"><%= oScate.FItemList(s).Fnmlarge %></a>
				<% if oScate.FItemList(s).Fcdsmall = cds then %></b><% end if %>
				<% if oScate.FResultCount-1 <> s then response.write " . " %>
			<%
			next

			end if
		end if
		%>
	</td>
</tr>
</table>

<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if ioffitem.FresultCount>0 then %>
			<input type="button" class="button" value="선택 아이템 추가" onclick="AddArr()">
			<% end if %>
		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan=10>
		<input type="checkbox" name="ckall" onClick="AnSelectAllFrame(this.checked)">전체선택
		&nbsp; 검색결과 : <b><%= ioffitem.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= Page %> / <%= ioffitem.FTotalPage %></b>
		※최대 3000건 까지 검색 됩니다
	</td>
</tr>
<% if ioffitem.fresultcount > 0 then %>
<tr align='center' bgcolor='#FFFFFF' >
	<td style='padding:10 0 0 0'>
		<table width="100%" align="center" cellpadding="2" cellspacing="0" class="a" >
		<%
		for i=0 to ioffitem.FResultCount -1

		makertr = makertr + 1

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
			<td width=200 >
				<table width="100%" border=0 cellspacing="0" cellpadding="0" class="a">
				<form name="frmBuyPrc_<%= i %>" >
				<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemList(i).Fitemgubun %>">
				<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
				<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">
				<input type="hidden" name="itemname" value="<%= ioffitem.FItemList(i).FShopItemName %>">
				<input type="hidden" name="itemoptionname" value="<%= ioffitem.FItemList(i).FShopItemOptionName %>">
				<input type="hidden" name="desingerid" value="<%= ioffitem.FItemList(i).FMakerid %>">
				<input type="hidden" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>">
				<input type="hidden" name="suplycash" value="<%= shopsuplycash %>">
				<% if IS_HIDE_BUYCASH = True then %>
				<input type="hidden" name="buycash" value="-1">
				<% else %>
				<input type="hidden" name="buycash" value="<%= buycash %>">
				<% end if %>
				<input type="hidden" name="foreign_sellcash" value="<%= getdisp_price(ioffitem.FItemList(i).Fforeign_sellprice, ioffitem.fcurrencyChar) %>">
				<input type="hidden" name="foreign_suplycash" value="<%= getdisp_price(foreign_suplycash, ioffitem.fcurrencyChar) %>">
				<tr align="left">
					<td height=280 valign="top">
						<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
						<img src="<%= ioffitem.FItemList(i).GetImageList %>" width=100 height=100 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" >
						<br>[<%= ioffitem.FItemList(i).FMakerid %>]
						<br><%= ioffitem.FItemList(i).GetBarCode %>
						<br><%= ioffitem.FItemList(i).FShopItemName %>
						<% if right(ioffitem.FItemList(i).GetBarCode,4)<>"0000" then %>
							<font color="blue">[<%= ioffitem.FItemList(i).FShopItemOptionName %>]</font>
						<% end if %>

						<br>판매가:<%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice,0) %>
						<br><font color="Gray">해외판매가:<%= getdisp_price_currencyChar(ioffitem.FItemList(i).Fforeign_sellprice, ioffitem.fcurrencyChar) %></font>

						<br>출고가:<%= FormatNumber(shopsuplycash,0) %>
						<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
							(<%= 100-(CLng(shopsuplycash/ioffitem.FItemList(i).Fshopitemprice*10000)/100) %>%)
						<% end if %>
						<br><font color="gray">해외출고가:<%= getdisp_price_currencyChar(foreign_suplycash, ioffitem.fcurrencyChar) %></font>

						<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
							<br>매입가:<%= FormatNumber(buycash,0) %>
							<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
								(<%= 100-(CLng(buycash/ioffitem.FItemList(i).Fshopitemprice*10000)/100) %>%)
							<% end if %>
						<% end if %>

						<% if ioffitem.FItemList(i).IsSoldOut then %>
						<font color="red">품절</font><br>
						<% end if %>
						<% if ioffitem.FItemList(i).Flimityn="Y" then %>
						<font color="blue">한정(<%= ioffitem.FItemList(i).getLimitNo %>)</font>
						<% end if %>

					    <% if Not (C_STOCK_DISP) then %>
						<br>7일판매량:<%= ioffitem.FItemList(i).FOffsell7days %>
						<br>3일판매량:<%= ioffitem.FItemList(i).Fsell3days %>
						<% end if %>
						<br><input type="text" class="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
					</td>
				</tr>
				</form>

				</table>
			</td>
		<%
		if makertr = 5 and ioffitem.fresultcount <> i + 1 then
			response.write "</tr><tr align='center' bgcolor='#FFFFFF'>"
			makertr = 0
		end if


		next

		if (ioffitem.fresultcount mod 5) = 1 then response.write "<td width=200></td><td width=200></td><td width=200></td><td width=200></td>"
		if (ioffitem.fresultcount mod 5) = 2 then response.write "<td width=200></td><td width=200></td><td width=200></td>"
		if (ioffitem.fresultcount mod 5) = 3 then response.write "<td width=200></td><td width=200></td>"
		if (ioffitem.fresultcount mod 5) = 4 then response.write "<td width=200></td>"
		%>
		</tr>
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
		<% else %>

		<tr bgcolor="#FFFFFF">
			<td align="center"><img src="http://fiximage.10x10.co.kr/web2008/category/list_none.gif " border="0"></td>
		</tr>

		<% end if %>
		</table>
	</td>

</tr>
</table>

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
set oLcate = nothing
set oMcate = nothing
set oScate = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
