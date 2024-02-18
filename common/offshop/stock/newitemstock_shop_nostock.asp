<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 상품주문검색
' Hieditor : 2011.08.22 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
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
dim page , shopid , isusing , makerid , itemid , itemname , generalbarcode , i , sell7days ,shopsuplycash ,buycash
dim cdl , cdm , cds , shortagetype , comm_cd ,includepreorder ,research , parameter , ipgo , order
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
    order = requestCheckVar(request("order"),32)

if page="" then page=1
if (research<>"on") and (includepreorder="") then
    'includepreorder = "on"
end if
if (research<>"on") and (ipgo="") then
    'ipgo = "on"
end if
if (research<>"on") and (shortagetype="") then
    'shortagetype = 7
end if
if (research<>"on") and (order="") then
    'order = "byrecent"
end if
if (research<>"on") and (isusing="") then
    isusing = "Y"
end if

'/매장일경우 본인 매장만 사용가능
if (C_IS_SHOP) then
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
parameter = parameter & "&ipgo="&ipgo&"&order="&order&""

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
    oshortage.Frectmakerid = makerid
    oshortage.Frectitemid = itemid
    oshortage.Frectitemname = itemname
    oshortage.Frectcomm_cd = comm_cd
    oshortage.Frectgeneralbarcode = generalbarcode
    oshortage.Frectshortagetype = shortagetype
    oshortage.Frectipgo = ipgo
    oshortage.Frectorder = order

    if shopid <> "" then
        ''임시 데이타 마트 서버 - 배치로 돌림..;;
        if (LCASE(shopid)="wholesale1043") or (LCASE(shopid)="wholesaletest") then
            if (page>1000) then
                ''1=a
                db3_dbget.close:dbget.Close:response.end
            end if
            oshortage.fnewitemstock_list_datamart
        else
            oshortage.fnewitemstock_list_nostock
        end if
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

<script language='javascript'>

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
function jumundirect(shopgubun){
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
    //매장
    if (shopgubun == 'True'){
    	upfrm.action='/common/offshop/shop_jumuninput.asp';
	//직원
	}else{
		upfrm.action='/admin/fran/jumuninput.asp';
	}
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
function jumundirect_upche(shopgubun){
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
    //매장
    if (shopgubun == 'True'){
    	upfrm.action='/common/offshop/shop_ipchulinput.asp';
	//직원
	}else{
		upfrm.action='/common/offshop/shop_ipchulinput.asp';
	}

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

<!--- 텐바이텐 물류센터 주문
<Br>&nbsp; &nbsp; 정산구분이 텐바이텐특정&출고매입 이고, 매장(주문자)이 동일 해야 주문서 작성 가능 합니다.
<br>- 업체 주문
<Br>&nbsp; &nbsp; 정산구분이 업체특정&업체매입 이고 , 브랜드(공급처)와 매장(주문자)이 동일 해야 주문서 작성 가능 합니다.<br>-->
<!--
필요수량(3일) = (3일판매분 x 1) - (유효재고 + 기주문건)
<br>필요수량(7일) = (7일판매분 x 1) - (유효재고 + 기주문건)
<br>필요수량(14일) = (7일판매분 x 2) - (유효재고 + 기주문건)
-->
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
        <% if (C_IS_SHOP) then %>
    		<%= shopid %>
    	<% else %>
        	<% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
        <% end if %>
        &nbsp;
        사용여부:<% drawSelectBoxUsingYN "isusing", isusing %>
        &nbsp;
        <!-- #include virtual="/common/module/categoryselectbox.asp"-->
        &nbsp;
        매입구분 : <% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
    </td>

    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="검색" onClick="javascript:reg('');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
        &nbsp;
        상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) reg('');">
        &nbsp;
        상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
        범용바코드 :
        <input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">

    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
    	<!--<input type=checkbox name="ipgo" <%' if ipgo = "on" then response.write " checked" %> onclick="reg('');">입고된것만
        <input type=checkbox name="sell7days" <%' if sell7days = "on" then response.write " checked" %> onclick="reg('');">최근7일판매내역있는것만
        <input type=checkbox name="includepreorder" <%' if includepreorder = "on" then response.write " checked" %> onclick="reg('');">기주문포함부족만&nbsp;
        재고부족 :
        <input type="radio" name="shortagetype" value="" <%' if shortagetype="" then response.write " checked" %> onclick="reg('');">전체&nbsp;
        <input type="radio" name="shortagetype" value="3" <%' if shortagetype="3" then response.write " checked" %> onclick="reg('');">3일후&nbsp;
        <input type="radio" name="shortagetype" value="7" <%' if shortagetype="7" then response.write " checked" %> onclick="reg('');">7일후&nbsp;
        <input type="radio" name="shortagetype" value="14" <%' if shortagetype="14" then response.write " checked" %> onclick="reg('');">14일후&nbsp;
		<font color="red">※현재줄의 재고관련검색시 입고된 내역만 검색됩니다</font>
		<br>-->
		<!--<input type="radio" name="shortagetype" value="28" <%' if shortagetype="28" then response.write " checked" %> onclick="reg('');">28일후-->
		상품구분 :
		<input type="radio" name="order" value="" <% if order="" then response.write " checked" %> onclick="reg('');">전체
		<input type="radio" name="order" value="byrecent" <% if order="byrecent" then response.write "checked" %> onclick="reg('');">신상품
		<input type="radio" name="order" value="byonbest" <% if order="byonbest" then response.write "checked" %> onclick="reg('');">온라인 베스트
		<input type="radio" name="order" value="byoffbest" <% if order="byoffbest" then response.write "checked" %> onclick="reg('');">오프라인 베스트
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

    </td>
    <td align="right">
    	<input type="button" class="button" value="장바구니담기" onclick="adminshoppingbagreg(frmbag,'OFF','<%=shopid%>')">
    	<input type="button" class="button" value="장바구니보기" onclick="adminshoppingbagview(frmbag,'OFF','<%=shopid%>')">
        <!--<%' if oshortage.FresultCount>0 then %>
            <input type="button" class="button" value="선택바로주문작성(텐바이텐물류)" onclick="jumundirect('<%=C_IS_SHOP%>')">
        <%' end if %>
        <%' if oshortage.FresultCount>0 then %>
        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
            	<input type="button" class="button" value="선택바로주문작성(업체)" onclick="jumundirect_upche('<%=C_IS_SHOP%>')">
            <%' end if %>
        <%' end if %>-->
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        검색결과 : <b><%= oshortage.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= oshortage.FTotalpage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width=20><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
    <td>
    	공급처
    </td>
    <td>브랜드</td>
    <td>상품코드</td>
    <td>이미지</td>
    <td>상품명<br><font color="blue">[옵션명]</font></td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
    	<td>매입가</td>
    <% end if %>
    <td>판매가</td>
    <td>매장<br>공급가</td>
    <td>공급<br>마진</td>
    <td>총<br>판매량</td>
    <td>7일<br>판매량</td>
    <td>3일<br>판매량</td>
    <td>매장재고</td>
    <td>수량</td>
    <td>비고</td>
</tr>
<% if oshortage.FresultCount > 0 then %>
<%
for i=0 to oshortage.FresultCount -1

shopsuplycash = oshortage.FItemList(i).GetFranchiseSuplycash
buycash		  = oshortage.FItemList(i).GetFranchiseBuycash
%>
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
<input type="hidden" name="shopsuplycash" value="0">
<input type="hidden" name="shopbuyprice" value="<%= oshortage.FItemlist(i).fshopbuyprice %>">
<input type="hidden" name="itemname" value="<%= oshortage.FItemlist(i).fshopitemname %>">
<input type="hidden" name="itemoptionname" value="<%= oshortage.FItemlist(i).fshopitemoptionname %>">
<input type="hidden" name="makerid" value="<%= oshortage.FItemlist(i).fmakerid %>">
<input type="hidden" name="shopid" value="<%= oshortage.FItemlist(i).fshopid %>">
    <td >
        <input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
    </td>
    <td>
        <%= GetdeliverGubunName(oshortage.FItemlist(i).fcomm_cd) %><br>(<%= GetJungsanGubunName(oshortage.FItemlist(i).fcomm_cd) %>)
    </td>
    <td>
        <a href="javascript:searchmakerid('<%= oshortage.FItemlist(i).fmakerid %>');" onfocus="this.blur()"><%= oshortage.FItemlist(i).fmakerid %></a>
    </td>
    <td>
        <%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %>
        <% if oshortage.FItemlist(i).Fitemgubun="10" then %>
        	<Br><a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%=oshortage.FItemlist(i).Fitemid%>" target="_blink" onfocus="this.blur()">[상세]</a>
        <% end if %>
    </td>
    <td><img src="<%= oshortage.FItemlist(i).GetImageSmall %>" width=50 height=50 border=0></td>
    <td>
        <%= oshortage.FItemlist(i).fshopitemname %><Br>
        <% if oshortage.FItemlist(i).fshopitemoptionname <> "" then %>
            <font color="blue">[<%=oshortage.FItemlist(i).fshopitemoptionname%>]<font>
        <% end if %>
    </td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
	    <td>
	        <%= FormatNumber(oshortage.FItemlist(i).fshopsuplycash,0) %>
	    </td>
	<% end if %>
    <td>
        <%= FormatNumber(oshortage.FItemlist(i).fshopitemprice,0) %>
    </td>
    <td>
        <%= FormatNumber(oshortage.FItemlist(i).fshopbuyprice,0) %>
    </td>
    <td>
		<% if oshortage.FItemList(i).Fshopitemprice<>0 then %>
		<%= 100-(CLng(shopsuplycash/oshortage.FItemList(i).Fshopitemprice*10000)/100) %> %
		<% end if %>
    </td>
    <td><%= oshortage.FItemlist(i).fsellno %></td><!--총판매현황 -->
    <td><%= oshortage.FItemlist(i).fsell7days %></td>
    <td><%= oshortage.FItemlist(i).fsell3days %></td>
    <td>
        <%= oshortage.FItemlist(i).frealStockno %>
    </td>
    <td>
        <input type="text" class="text" name="itemno" value="0" size="3" maxlength="4" onKeyDown="CheckThis(frmBuyPrc<%= i %>);">
    </td>
    <td>
        <% if oshortage.FItemList(i).Fpreorderno>0 then %>
        	기주문:
            <% if oshortage.FItemList(i).Fpreorderno<>oshortage.FItemList(i).Fpreordernofix then response.write CStr(oshortage.FItemList(i).Fpreorderno) + " -> " %>
        	<%= oshortage.FItemList(i).Fpreordernofix %><br>
        <% end if %>
        <img src="/images/cartimage.jpg" style="cursor:pointer" onclick="adminshoppingbagregoneitem('OFF','<%=shopid%>',frmBuyPrc<%= i %>)">
    </td>
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
    <td colspan="25" align="center">
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
    <td colspan="25" align="center">[검색결과가 없습니다.]</td>
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

    </td>
    <td align="right">
    	<input type="button" class="button" value="장바구니담기" onclick="adminshoppingbagreg(frmbag,'OFF','<%=shopid%>')">
    	<input type="button" class="button" value="장바구니보기" onclick="adminshoppingbagview(frmbag,'OFF','<%=shopid%>')">
        <!--<%' if oshortage.FresultCount>0 then %>
            <input type="button" class="button" value="선택바로주문작성(텐바이텐물류)" onclick="jumundirect('<%=C_IS_SHOP%>')">
        <%' end if %>
        <%' if oshortage.FresultCount>0 then %>
        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
            	<input type="button" class="button" value="선택바로주문작성(업체)" onclick="jumundirect_upche('<%=C_IS_SHOP%>')">
            <%' end if %>
        <%' end if %>-->
    </td>
</tr>
</table>
<!-- 액션 끝 -->
<%
    set oshortage = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
