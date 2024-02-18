<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 매장적정재고부족상품
' Hieditor : 2011.07.13 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shortagestock_cls.asp" -->
<!-- #include virtual="/lib/classes/items/adminshoppingbag/adminshoppingbag_cls.asp" -->
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

dim page , shopid , isusing , makerid , itemid , itemname , generalbarcode , i , sell7days
dim cdl , cdm , cds , shortagetype , comm_cd ,includepreorder ,research , parameter , ipgo
dim onlinemwdiv, ordby, sellyn
dim forcemakerid
    page = requestCheckVar(getNumeric(request("page")),10)
    research = requestCheckVar(request("research"),2)
    isusing = requestCheckVar(request("isusing"),1)
    makerid = requestCheckVar(request("makerid"),32)
    itemid = requestCheckVar(request("itemid"),10)
    itemname = requestCheckVar(request("itemname"),64)
    generalbarcode = requestCheckVar(request("generalbarcode"),20)
    comm_cd = requestCheckVar(request("comm_cd"),16)
    cdl = requestCheckVar(getNumeric(request("cdl")),3)
    cdm = requestCheckVar(getNumeric(request("cdm")),3)
    cds = requestCheckVar(getNumeric(request("cds")),3)
    shortagetype = requestCheckVar(request("shortagetype"),10)
    includepreorder = requestCheckVar(request("includepreorder"),2)
    sell7days = requestCheckVar(request("sell7days"),2)
    ipgo = requestCheckVar(request("ipgo"),2)
	shopid = requestCheckVar(request("shopid"),32)
	onlinemwdiv = requestCheckVar(request("onlinemwdiv"),1)
	ordby = requestCheckVar(request("ordby"),32)
	sellyn      = requestCheckvar(request("sellyn"),2)
    forcemakerid = RequestCheckVar(request("forcemakerid"),32)

if page="" then page=1
if (research<>"on") and (sellyn="") then
    sellyn = "YS"
end if
if (research<>"on") and (includepreorder="") then
    includepreorder = "on"
end if
if (research<>"on") and (ipgo="") then
    ipgo = "on"
end if
if (research<>"on") and (shortagetype="") then
    shortagetype = 7
end if
if (research<>"on") and (isusing="") then
    isusing = "Y"
end if

if (research<>"on") and (ordby="") then
    ordby = "BI"
end if

if C_ADMIN_USER then

'/매장일경우 본인 매장만 사용가능
elseif (C_IS_SHOP) then
	'/어드민권한 점장 미만
	'if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	'end if
else
	if (C_IS_Maker_Upche) then
		shopid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then

		else

		end if
	end if
end if

'if shopid = "" then shopid = "streetshop011"

parameter = "page="&page&"&research="&research&"&shopid="&shopid&"&isusing="&isusing&"&makerid="&makerid&"&itemid="&itemid&"&itemname="&itemname&"&sell7days="&sell7days&""
parameter = parameter & "&generalbarcode="&generalbarcode&"&comm_cd="&comm_cd&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&shortagetype="&shortagetype&"&includepreorder="&includepreorder
parameter = parameter & "&ipgo="&ipgo&"&sellyn="&sellyn&""

dim oshortage
set oshortage  = new cshortagestock_list
    oshortage.FPageSize = 100
    oshortage.FCurrPage = page
    oshortage.frectcdl = cdl
    oshortage.frectcdm = cdm
    oshortage.frectcds = cds
    oshortage.frectincludepreorder = includepreorder
    oshortage.frectsell7days = sell7days
    oshortage.Frectshopid = shopid
    oshortage.Frectisusing = isusing
    if forcemakerid = "" then
        oshortage.Frectmakerid = makerid
    else
        oshortage.Frectmakerid = forcemakerid
    end if
    oshortage.Frectitemid = itemid
    oshortage.Frectitemname = itemname
    oshortage.Frectcomm_cd = comm_cd
    oshortage.Frectgeneralbarcode = generalbarcode
    oshortage.Frectshortagetype = shortagetype
    oshortage.Frectipgo = ipgo
	oshortage.FRectOnlineMWdiv = onlinemwdiv
	oshortage.FRectOrder = ordby
	oshortage.FRectSellYN       = sellyn

    if shopid <> "" then
        oshortage.fshortagestock_list
    else
        response.write "<script language='javascript'>"
        response.write "    alert('매장을 선택해 주세요');"
        response.write "</script>"
    end if

dim IsUpcheWitakItem
if (makerid<>"") and (shopid<>"") then
    IsUpcheWitakItem = (GetShopBrandContract(shopid,makerid)="B012")
end if
%>

<script type="text/javascript">

//엑셀다운로드
function exceldownload(){
    var exceldownload = window.open('/common/offshop/stock/shortagestock_shop_excel.asp?<%=parameter%>','exceldownload','width=1024,height=768,scrollbars=yes,resizable=yes');
    exceldownload.focus();
}

//필요수량클릭시
function inputiteno(shortitemno,formi){
    formi.itemno.value=shortitemno;

    formi.cksel.checked=true;
    AnCheckClick(formi.cksel);
}

//브랜드클릭시
function searchmakerid(makerid){
    frm.makerid.value=makerid;
    frm.submit();
}

function CheckThis(frm){
    frm.cksel.checked=true;
    AnCheckClick(frm.cksel);
}

//검색버튼
function reg(page){

    if(frm.itemid.value!=''){
        if (!IsDouble(frm.itemid.value)){
            alert('상품코드는 숫자만 가능합니다.');
            frm.itemid.focus();
            return;
        }
    }

    frm.page.value=page;
    frm.submit();
}

//실사재고입력
function popOffErrInput(shopid,itemgubun,itemid,itemoption){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
        alert('권한이 없습니다. - 업체특정 상품만 재고 수정 가능.');
        return; //업체특정 상품인 경우?
    <% else %>
        var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popAdmOffrealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
        popwin.focus();
    <% end if %>
}

var jumunitem = null;

//텐바이텐 주문서 작성
function jumundirect(){
    var upfrm = document.frmArrupdate;
    var frm; var tmpshopid = '';
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

    upfrm.itemgubunarr2.value = "";
    upfrm.itemidadd2.value = "";
    upfrm.itemoptionarr2.value = "";
    upfrm.sellcasharr2.value = "";
    upfrm.suplycasharr2.value = "";
    upfrm.buycasharr2.value = "";
    upfrm.itemnoarr2.value = "";
    upfrm.itemnamearr2.value = "";
    upfrm.itemoptionnamearr2.value = "";
    upfrm.designerarr2.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

                if (tmpshopid==''){
                	tmpshopid = frm.shopid.value;
                } else {
                	if (tmpshopid != frm.shopid.value){
	                    alert('서로 틀린 매장이 선택되어 있습니다 \n매장(주문자)이 동일 해야 합니다.');
	                    return;
                	}
                }

                if (frm.comm_cd.value=="B012" || frm.comm_cd.value=="B022"){
                    alert('업체특정이나 업체매입은 주문 하실수 없습니다.');
                    frm.itemno.focus();
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

                upfrm.itemgubunarr2.value = upfrm.itemgubunarr2.value + frm.itemgubun.value + "|";
                upfrm.itemidadd2.value = upfrm.itemidadd2.value + frm.itemid.value + "|";
                upfrm.itemoptionarr2.value = upfrm.itemoptionarr2.value + frm.itemoption.value + "|";
                upfrm.sellcasharr2.value = upfrm.sellcasharr2.value + frm.shopitemprice.value + "|";

                //[db_storage].[dbo].tbl_ordersheet_master에 들어가는 내용의 경우 센터매입가와 , 매장매입가가 꺼꾸로임
                upfrm.suplycasharr2.value = upfrm.suplycasharr2.value + frm.shopbuyprice.value + "|";
                upfrm.buycasharr2.value = upfrm.buycasharr2.value + frm.shopsuplycash.value + "|";
                upfrm.itemnoarr2.value = upfrm.itemnoarr2.value + frm.itemno.value + "|";
                upfrm.itemnamearr2.value = upfrm.itemnamearr2.value + frm.itemname.value + "|";
                upfrm.itemoptionnamearr2.value = upfrm.itemoptionnamearr2.value + frm.itemoptionname.value + "|";
                upfrm.designerarr2.value = upfrm.designerarr2.value + frm.makerid.value + "|";
            }
        }
    }

	//폼 전송 형식... 사용안함... 이전 상품 내역 날라감..
    var jumunitem = window.open('','jumunitem','width=1024,height=768,scrollbars=yes,resizable=yes');

	<% if C_ADMIN_USER then %>
    	upfrm.action='/admin/fran/jumuninput.asp';
    <% else %>
    	upfrm.action='/common/offshop/shop_jumuninput.asp';
    <% end if %>

    upfrm.target='jumunitem';
    upfrm.shopid.value=tmpshopid;
    upfrm.submit();
    jumunitem.focus();

    //기존 팝업이 띄어 있는 경우
	//if(jumunitem != null){
    //}

    //기존 팝업이 없는경우 팝업 띄운다
    //else {
    //	jumunitem = window.open('/admin/fran/jumuninput.asp?suplyer=10x10&shopid=<%=shopid%>','jumunitem','width=1024,height=768,scrollbars=yes,resizable=yes');
	//}

	//팝업 로드후 0.1후 뒤에.. 상품 때려넣음... 이렇게 해야 기존 상품 내역 남음..
	//window.setTimeout("jumunitem.ReActItems('0',frmArrupdate.itemgubunarr2.value,frmArrupdate.itemidadd2.value,frmArrupdate.itemoptionarr2.value,frmArrupdate.sellcasharr2.value,frmArrupdate.suplycasharr2.value,frmArrupdate.buycasharr2.value,frmArrupdate.itemnoarr2.value,frmArrupdate.itemnamearr2.value,frmArrupdate.itemoptionnamearr2.value,frmArrupdate.designerarr2.value);",500)
	//jumunitem.focus();

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
                frm.cksel.checked = false;
                frm.itemno.value="0"
                dL(frm.cksel);
            }
        }
    }
}

//업체 주문서 작성
function jumundirect_upche(){
    var upfrm = document.frmArrupdate;
    var frm; var tmpshopid = ''; var tmpmakerid = '';
    var pass = false;
    var searchfrm = document.frm;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;

    //if (searchfrm.makerid.value == ''){
    //    alert('브랜드(공급처)를 선택해 주세요.');
    //    return;
    //}

    if (!pass) {
        alert('선택 아이템이 없습니다.');
        return;
    }

    upfrm.itemgubunarr2.value = "";
    upfrm.itemidadd2.value = "";
    upfrm.itemoptionarr2.value = "";
    upfrm.sellcasharr2.value = "";
    upfrm.suplycasharr2.value = "";
    upfrm.shopbuypricearr2.value = "";
    upfrm.itemnoarr2.value = "";
    upfrm.itemnamearr2.value = "";
    upfrm.itemoptionnamearr2.value = "";
    upfrm.designerarr2.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

                if (tmpshopid==''){
                	tmpshopid = frm.shopid.value;
                } else {
                	if (tmpshopid != frm.shopid.value){
	                    alert('서로 틀린 매장이 선택되어 있습니다 \n매장(주문자)이 동일 해야 합니다.');
	                    return;
                	}
                }

                if (tmpmakerid==''){
                	tmpmakerid = frm.makerid.value;
                } else {
                	if (tmpmakerid != frm.makerid.value){
	                    alert('서로 틀린 브랜드(공급처)가 선택되어 있습니다 \n업체주문의 경우 브랜드(공급처)가 동일해야 합니다');
	                    return;
                	}
                }

                if (frm.comm_cd.value=="B011" || frm.comm_cd.value=="B031"){
                    alert('텐바이텐특정이나 출고매입은 주문 하실수 없습니다.');
                    frm.itemno.focus();
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

                upfrm.itemgubunarr2.value = upfrm.itemgubunarr2.value + frm.itemgubun.value + "|";
                upfrm.itemidadd2.value = upfrm.itemidadd2.value + frm.itemid.value + "|";
                upfrm.itemoptionarr2.value = upfrm.itemoptionarr2.value + frm.itemoption.value + "|";
                upfrm.sellcasharr2.value = upfrm.sellcasharr2.value + frm.shopitemprice.value + "|";
                upfrm.suplycasharr2.value = upfrm.suplycasharr2.value + frm.shopsuplycash.value + "|";
                upfrm.shopbuypricearr2.value = upfrm.shopbuypricearr2.value + frm.shopbuyprice.value + "|";
                upfrm.itemnoarr2.value = upfrm.itemnoarr2.value + frm.itemno.value + "|";
                upfrm.itemnamearr2.value = upfrm.itemnamearr2.value + frm.itemname.value + "|";
                upfrm.itemoptionnamearr2.value = upfrm.itemoptionnamearr2.value + frm.itemoptionname.value + "|";
                upfrm.designerarr2.value = upfrm.designerarr2.value + frm.makerid.value + "|";
            }
        }
    }

	//폼 전송 형식... 사용안함... 이전 상품 내역 날라감..
    var jumunitem = window.open('','jumunitem','width=1024,height=768,scrollbars=yes,resizable=yes');

	<% if C_ADMIN_USER then %>
    	upfrm.action='/common/offshop/shop_ipchulinput.asp';
    <% else %>
    	upfrm.action='/common/offshop/shop_ipchulinput.asp';
    <% end if %>

    upfrm.target='jumunitem';
    upfrm.shopid.value=tmpshopid;
    upfrm.chargeid.value=tmpmakerid;
    upfrm.submit();
    jumunitem.focus();

    //기존 팝업이 띄어 있는 경우
	//if(jumunitem != null){
    //}

    //기존 팝업이 없는경우 팝업 띄운다
    //else {
    //	jumunitem = window.open('/common/offshop/shop_ipchulinput.asp?chargeid=<%=makerid%>&shopid=<%=shopid%>&isreq=Y','jumunitem','width=1024,height=768,scrollbars=yes,resizable=yes');
	//}

	//팝업 로드후 0.1후 뒤에.. 상품 때려넣음... 이렇게 해야 기존 상품 내역 남음..
	//window.setTimeout("jumunitem.ReActItems(frmArrupdate.itemgubunarr2.value,frmArrupdate.itemidadd2.value,frmArrupdate.itemoptionarr2.value,frmArrupdate.sellcasharr2.value,frmArrupdate.suplycasharr2.value,frmArrupdate.shopbuypricearr2.value,frmArrupdate.itemnoarr2.value,frmArrupdate.itemnamearr2.value,frmArrupdate.itemoptionnamearr2.value,frmArrupdate.designerarr2.value);",500)
	//jumunitem.focus();

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
                frm.cksel.checked = false;
                frm.itemno.value="0"
                dL(frm.cksel);
            }
        }
    }
}

</script>

- 텐바이텐 물류센터 주문
<Br>&nbsp; &nbsp; 정산구분이 텐바이텐특정&출고매입 이고, 매장(주문자)이 동일 해야 주문서 작성 가능 합니다.
<br>- 업체 주문
<Br>&nbsp; &nbsp; 정산구분이 업체특정&업체매입 이고 , 브랜드(공급처)와 매장(주문자)이 동일 해야 주문서 작성 가능 합니다.
<br>필요수량(3일) = (3일판매분 x 1) - (유효재고 + 기주문건)
<br>필요수량(7일) = (7일판매분 x 1) - (유효재고 + 기주문건)
<br>필요수량(14일) = (7일판매분 x 2) - (유효재고 + 기주문건)
<!--<br>&nbsp; &nbsp; 필요수량(28일) = (7일판매분 x 4) - (유효재고 + 기주문건)-->

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%=page%>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
    <td align="left">
        매장 :
        <% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
        	<% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
        <% elseif (C_IS_SHOP) then %>
    		<%= shopid %>
    	<% else %>
        	<% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
        <% end if %>
        사용여부:<% drawSelectBoxUsingYN "isusing", isusing %>
        온라인판매여부:<% drawSelectBoxSellYN "sellyn", sellyn %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			&nbsp;
			ON매입구분:
			<select class="select" name="onlinemwdiv">
				<option></option>
				<option value="M" <% if (onlinemwdiv = "M") then %>selected<% end if %> >매입</option>
				<option value="W" <% if (onlinemwdiv = "W") then %>selected<% end if %> >특정</option>
				<option value="U" <% if (onlinemwdiv = "U") then %>selected<% end if %> >업체</option>
			</select>
		<% end if %>

		&nbsp;
        <!-- #include virtual="/common/module/categoryselectbox.asp"-->
    </td>

    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="검색" onClick="javascript:reg('');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        <% if C_ADMIN_AUTH then %>
        * 브랜드[관리자] <input type="text" class="text" name="forcemakerid" value="<%= forcemakerid %>" size="16" maxlength="32">
        &nbsp;&nbsp;
        <% end if %>
        브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
        &nbsp;
        상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) reg('');">
        &nbsp;
        상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
        범용바코드 :
        <input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
        정산기준 : <% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
    	<input type=checkbox name="ipgo" <% if ipgo = "on" then response.write " checked" %> onclick="reg('');">입고된것만(매장)
        <input type=checkbox name="sell7days" <% if sell7days = "on" then response.write " checked" %> onclick="reg('');">최근7일판매내역있는것만
        <input type=checkbox name="includepreorder" <% if includepreorder = "on" then response.write " checked" %> onclick="reg('');">기주문포함부족만&nbsp;
        재고부족 : <input type="radio" name="shortagetype" value="" <% if shortagetype="" then response.write " checked" %> onclick="reg('');">전체&nbsp;
        <input type="radio" name="shortagetype" value="3" <% if shortagetype="3" then response.write " checked" %> onclick="reg('');">3일후&nbsp;
        <input type="radio" name="shortagetype" value="7" <% if shortagetype="7" then response.write " checked" %> onclick="reg('');">7일후&nbsp;
        <input type="radio" name="shortagetype" value="14" <% if shortagetype="14" then response.write " checked" %> onclick="reg('');">14일후&nbsp;
        <!--<input type="radio" name="shortagetype" value="28" <%' if shortagetype="28" then response.write " checked" %> onclick="reg('');">28일후-->
		&nbsp;
		정렬순서 :
		<select class="select" name="ordby">
			<option value="BI" <% if (ordby = "BI") then %>selected<% end if %> >브랜드</option>
			<option value="I" <% if (ordby = "I") then %>selected<% end if %> >상품코드 역순</option>
		</select>
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
		<input type="button" class="button" value="엑셀다운로드" onclick="exceldownload()">
    </td>
    <td align="right">
    	<input type="button" class="button" value="장바구니담기" onclick="adminshoppingbagreg(frmbag,'OFF','<%=shopid%>')">
    	<input type="button" class="button" value="장바구니보기" onclick="adminshoppingbagview(frmbag,'OFF','<%=shopid%>')">
        <% if oshortage.FresultCount>0 then %>
            <input type="button" class="button" value="선택바로주문작성(텐바이텐물류)" onclick="jumundirect()">
        <% end if %>
        <% if oshortage.FresultCount>0 then %>
        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
            	<input type="button" class="button" value="선택바로주문작성(업체)" onclick="jumundirect_upche()">
            <%' end if %>
        <% end if %>
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="30">
        검색결과 : <b><%= oshortage.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= oshortage.FTotalpage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width=20><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
    <td width=80>매장</td>
    <td>계약구분</td>
	<td width=50>이미지</td>
    <td>브랜드</td>
    <td width=80>상품코드</td>
    <td>상품명</td>
	<td>옵션명</td>
	<td width=50>판매가</td>
	<td width=50>출고가</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
    	<td width=50>매입가</td>
    <% end if %>

    <td width=30>센터<br>입고</td>
    <td width=30>업체<br>입고</td>
    <td width=30>판매</td>
    <td width=40>시스템<br>재고</td>
    <td width=30>오차</td>
	<td width=30>실사<br>재고</td>
    <td width=30>샘플</td>
    <td width=30>유효<br>재고</td>

	<td width=30>OFF<br>3일</td>
	<td width=30>OFF<br>7일</td>
	<td width=30>3일</td>
	<td width=30>7일</td>
	<td width=30>14일</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
		<td width=30>ON<br>매입<br>구분</td>
		<td width=30>센터<br>매입<br>구분</td>
	<% end if %>

    <td width=40>주문<br>수량</td>
    <td width=90>비고</td>
</tr>
<% if oshortage.FresultCount > 0 then %>
<% for i=0 to oshortage.FresultCount -1 %>
<form method="get" action="" name="frmBuyPrc<%=i%>">

<% if oshortage.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>
<input type="hidden" name="comm_cd" value="<%= oshortage.FItemlist(i).fcomm_cd %>">
<input type="hidden" name="itemgubun" value="<%= oshortage.FItemlist(i).fitemgubun %>">
<input type="hidden" name="itemid" value="<%= oshortage.FItemlist(i).fitemid %>">
<input type="hidden" name="itemoption" value="<%= oshortage.FItemlist(i).fitemoption %>">
<input type="hidden" name="shopitemprice" value="<%= oshortage.FItemlist(i).fshopitemprice %>">
<% if IS_HIDE_BUYCASH = True then %>
<input type="hidden" name="shopsuplycash" value="-1">
<% else %>
<input type="hidden" name="shopsuplycash" value="<%= oshortage.FItemlist(i).fshopsuplycash %>">
<% end if %>
<input type="hidden" name="shopbuyprice" value="<%= oshortage.FItemlist(i).fshopbuyprice %>">
<input type="hidden" name="itemname" value="<%= oshortage.FItemlist(i).fshopitemname %>">
<input type="hidden" name="itemoptionname" value="<%= oshortage.FItemlist(i).fshopitemoptionname %>">
<input type="hidden" name="makerid" value="<%= oshortage.FItemlist(i).fmakerid %>">
<input type="hidden" name="shopid" value="<%= oshortage.FItemlist(i).fshopid %>">
    <td >
        <input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
    </td>
    <td >
        <%= oshortage.FItemlist(i).fshopid %>
    </td>
    <td>
        <%= GetJungsanGubunName(oshortage.FItemlist(i).fcomm_cd) %>
    </td>
    <td>
        <img src="<%= oshortage.FItemlist(i).GetImageSmall %>" width=50 height=50 border=0>
    </td>
    <td>
        <a href="javascript:searchmakerid('<%= oshortage.FItemlist(i).fmakerid %>');" onfocus="this.blur()"><%= oshortage.FItemlist(i).fmakerid %></a>
    </td>
    <td>
		<!--
        <a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= oshortage.FItemList(i).Fitemgubun %>&itemid=<%= oshortage.FItemList(i).Fitemid %>&itemoption=<%= oshortage.FItemList(i).Fitemoption %>" target=_blank ><%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %></a>
		-->
		<a href="/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=<%= oshortage.FItemlist(i).fshopid %>&barcode=<%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %>" target=_blank ><%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %></a>
    </td>
    <td align="left">
        <%= oshortage.FItemlist(i).fshopitemname %><Br>
    </td>
    <td align="left">
        <% if oshortage.FItemlist(i).fshopitemoptionname <> "" then %>
            <%=oshortage.FItemlist(i).fshopitemoptionname%>
        <% end if %>
    </td>
    <td align="right">
        <%= FormatNumber(oshortage.FItemlist(i).fshopitemprice,0) %>
    </td>
    <td align="right">
        <%= FormatNumber(oshortage.FItemlist(i).fshopbuyprice,0) %>
		<% if oshortage.FItemList(i).Fshopitemprice<>0 then %>
		<br>(<%= 100-Clng(oshortage.FItemList(i).fshopbuyprice/oshortage.FItemList(i).Fshopitemprice*100*100)/100 %> %)
		<% end if %>
    </td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	    <td align="right">
	        <%= FormatNumber(oshortage.FItemlist(i).fshopsuplycash,0) %>
	    </td>
	<% end if %>

    <td align="right">
        <%= oshortage.FItemlist(i).flogicsipgono + oshortage.FItemlist(i).flogicsreipgono %>    <!--센터입고반품-->
    </td>
    <td align="right">
        <%= oshortage.FItemlist(i).fbrandipgono + oshortage.FItemlist(i).fbrandreipgono %>		<!--브랜드입고반품-->
    </td>
    <td align="right">
        <%= oshortage.FItemlist(i).fsellno+oshortage.FItemlist(i).fresellno %>       <!--총판매현황 -->
    </td>
    <td bgcolor="#EEEEFF" align="right">
        <b><%= oshortage.FItemlist(i).fsysstockno %></b>       <!--시스템재고-->
    </td>
    <td align="right">
        <%= oshortage.FItemlist(i).Ferrrealcheckno %>       <!--오차-->
    </td>
    <td bgcolor="#EEEEFF" align="right">
		<b><%= oshortage.FItemlist(i).frealstockno %></b>          <!-- 실사재고 -->
    </td>
    <td align="right">
        <%= oshortage.FItemlist(i).ferrsampleitemno %>      <!--샘플-->
    </td>
    <td bgcolor="#EEEEFF" align="right">
        <b><%= oshortage.FItemlist(i).getAvailStock %></b>     <!--유효재고-->
    </td>
	<td align="right"><%= oshortage.FItemlist(i).fsell3days %></td>		<!--판매수량-->
	<td align="right"><%= oshortage.FItemlist(i).fsell7days %></td>
	<td align="right">													<!-- 총필요수량 -->
        <% if oshortage.FItemlist(i).frequire3daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire3daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()">
            <font color="red"><%= oshortage.FItemlist(i).frequire3daystock*-1 %></font>
            </a>
        <% else %>
            0
        <% end if %>
	</td>
	<td align="right">
        <% if oshortage.FItemlist(i).frequire7daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire7daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()">
            <font color="red"><%= oshortage.FItemlist(i).frequire7daystock*-1 %></font>
            </a>
        <% else %>
            0
        <% end if %>
	</td>
	<td align="right">
        <% if oshortage.FItemlist(i).frequire14daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire14daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()">
            <font color="red"><%= oshortage.FItemlist(i).frequire14daystock*-1 %></font>
            </a>
        <% else %>
            0
        <% end if %>
	</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
	    <td><%= oshortage.FItemlist(i).Fonlinemwdiv %></td>
		<td><%= oshortage.FItemlist(i).fcentermwdiv %></td>
	<% end if %>

    <td>
        <input type="text" class="text" name="itemno" value="0" size="2" maxlength="4" onKeyDown="CheckThis(frmBuyPrc<%= i %>);">
    </td>
    <td>
        <% if oshortage.FItemList(i).Fpreorderno>0 or oshortage.FItemList(i).Fpreorderno<0 then %>
        	기주문:
            <% if oshortage.FItemList(i).Fpreorderno<>oshortage.FItemList(i).Fpreordernofix then response.write CStr(oshortage.FItemList(i).Fpreorderno) + " -> " %>
        	<%= oshortage.FItemList(i).Fpreordernofix %><br>
        <% end if %>
        <% if oshortage.FItemList(i).IsSoldOut then %>
			<font color="red">ON:품절</font><Br>
        <% end if %>
		<% if oshortage.FItemList(i).Flimityn="Y" then %>
			<font color="blue">ON:한정(<%= oshortage.FItemList(i).getLimitNo %>)</font><Br>

		<% end if %>
        <input type="button" class="button" value="실사입력" onclick="popOffErrInput('<%= shopid %>','<%= oshortage.FItemList(i).Fitemgubun %>','<%= oshortage.FItemList(i).Fitemid %>','<%= oshortage.FItemList(i).Fitemoption %>');">     <!--실사재고입력-->
    </td>
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
    <td colspan="30" align="center">
        <% if oshortage.HasPreScroll then %>
            <span class="list_link"><a href="javascript:reg(<%=oshortage.StartScrollPage-1%>)">[pre]</a></span>
        <% else %>
        [pre]
        <% end if %>
        <% for i = 0 + oshortage.StartScrollPage to oshortage.StartScrollPage + oshortage.FScrollCount - 1 %>
            <% if (i > oshortage.FTotalpage) then Exit for %>
            <% if CStr(i) = CStr(oshortage.FCurrPage) then %>
            <span class="page_link"><font color="red"><b><%= i %></b></font></span>
            <% else %>
            <a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
            <% end if %>
        <% next %>
        <% if oshortage.HasNextScroll then %>
            <span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
        <% else %>
        [next]
        <% end if %>
    </td>
</tr>

<% else %>

<tr bgcolor="#FFFFFF">
    <td colspan="30" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
<form name="frmArrupdate" method="post">
    <input type="hidden" name="mode" value="arrins">
    <input type="hidden" name="itemgubunarr2" value="">
    <input type="hidden" name="itemidadd2" value="">
    <input type="hidden" name="itemoptionarr2" value="">
    <input type="hidden" name="sellcasharr2" value="">
    <input type="hidden" name="buycasharr2" value="">
    <input type="hidden" name="suplycasharr2" value="">
    <input type="hidden" name="itemnoarr2" value="">
    <input type="hidden" name="itemnamearr2" value="">
    <input type="hidden" name="itemoptionnamearr2" value="">
    <input type="hidden" name="designerarr2" value="">
    <input type="hidden" name="shopid" value="<%=shopid%>">
    <input type="hidden" name="suplyer" value="10x10">
    <input type="hidden" name="idx" value="0">
    <input type="hidden" name="chargeid" value="<%=makerid%>">
    <input type="hidden" name="shopbuypricearr2" value="">
    <input type="hidden" name="isreq" value="Y">
</form>
<form name="frmbag" method="post">
    <input type="hidden" name="onoffgubun">
    <input type="hidden" name="itemgubunarr">
    <input type="hidden" name="itemidarr">
    <input type="hidden" name="itemoptionarr">
    <input type="hidden" name="itemnoarr">
    <input type="hidden" name="makerid">
    <input type="hidden" name="shopid" >
</form>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td align="left">
		<input type="button" class="button" value="엑셀다운로드" onclick="exceldownload()">
    </td>
    <td align="right">
    	<input type="button" class="button" value="장바구니담기" onclick="adminshoppingbagreg(frmbag,'OFF','<%=shopid%>')">
    	<input type="button" class="button" value="장바구니보기" onclick="adminshoppingbagview(frmbag,'OFF','<%=shopid%>')">
        <% if oshortage.FresultCount>0 then %>
            <input type="button" class="button" value="선택바로주문작성(텐바이텐물류)" onclick="jumundirect()">
        <% end if %>
        <% if oshortage.FresultCount>0 then %>
        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
            	<input type="button" class="button" value="선택바로주문작성(업체)" onclick="jumundirect_upche()">
            <%' end if %>
        <% end if %>
    </td>
</tr>
</table>
<!-- 액션 끝 -->
<%
    set oshortage = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
