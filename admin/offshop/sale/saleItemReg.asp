<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  할인 관리
' History : 2010.12.01 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/sale/sale_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!--<H1><font color=red>수정중</font></H1>-->
<%
Dim sCode, clsSale,clsSaleItem ,acURL ,iTotCnt, arrList,i , shopid , para , adminvspos ,point_rate
Dim sTitle,isRate, isMargin, isStatus,eCode, egCode, dSDay, dEDay, isUsing, dOpenDay,isMValue, smargin ,sellpricemargin
Dim ix,page , sale_shopmargin , sale_shopmarginvalue , sshopmargin , designer , itemid , itemname, itemcontract
	adminvspos = requestCheckVar(Request("adminvspos"),2)
	sCode = requestCheckVar(Request("sC"),10)
	designer    = RequestCheckVar(request("designer"),32)
	itemid      = RequestCheckVar(request("itemid"),9)
	itemname    = RequestCheckVar(request("itemname"),32)
	acURL =Server.HTMLEncode("/admin/offshop/sale/saleitemProc.asp?sC="&sCode)
	page = requestCheckVar(request("page"),10)

if page = "" then page = 1

para = "designer="&designer&"&itemid="&itemid&"&itemname="&itemname&"&menupos="&menupos&"&adminvspos="&adminvspos

'할인 기본정보
set clsSale = new CSale
	clsSale.FSCode  = sCode
	clsSale.fnGetSaleConts

	sTitle 		= clsSale.FSName
	isRate 		= clsSale.FSRate
	point_rate 		= clsSale.fpoint_rate
	isMargin 	= clsSale.FSMargin
	eCode 		= clsSale.FECode
	egCode		= clsSale.FEGroupCode
	dSDay 		= clsSale.FSDate
	dEDay 		= clsSale.FEDate
	isStatus 	= clsSale.FSStatus
	isUsing     = clsSale.FSUsing
	dOpenDay	= clsSale.FOpenDate
	isMValue	= clsSale.FSMarginValue
	sale_shopmargin = clsSale.fsale_shopmargin
	sale_shopmarginvalue	= clsSale.fsale_shopmarginvalue
	shopid = clsSale.fshopid
set clsSale = nothing

'rw isMValue
'rw sale_shopmarginvalue
'할인 상품정보
set clsSaleItem = new CSaleItem
	clsSaleItem.FPageSize = 100
	clsSaleItem.FCurrPage = page
	clsSaleItem.FSCode = sCode
	clsSaleItem.FRectDesigner = designer
	clsSaleItem.FRectItemID = itemid
	clsSaleItem.frectadminvspos = adminvspos
	clsSaleItem.FRectItemName = html2db(itemname)
	clsSaleItem.fnGetSaleItemList()

'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
Dim arrsalemargin, arrsalestatus
	arrsalemargin = fnSetCommonCodeArr_off("salemargin",False)
	arrsalestatus= fnSetCommonCodeArr_off("salestatus",False)
%>

<script language="javascript">

// 페이지 이동
function jsGoPage(iP){
	location.href="saleItemReg.asp?menupos=<%=menupos%>&sC=<%=sCode%>&iC="+iP;
}

// 상품추가(검색) 팝업
function addnewItem(eC,egC){
	var popwin;

	if ( eC > 0 ){
		popwin = window.open("/common/offshop/pop_eventitem_addinfo_off.asp?acURL=<%=acURL%>&eC="+eC+"&egC="+egC, "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	}else{
		popwin = window.open("/common/offshop/pop_itemAddInfo_off.asp?shopid=<%=shopid%>&acURL=<%=acURL%>", "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	}
	popwin.focus();
}

// 상품추가(브랜드) 팝업
function addnewbrand(eC,egC){
	var addnewbrand;

	if ( eC > 0 ){
	}else{
		addnewbrand = window.open("/common/offshop/pop_itembrandAddInfo_off.asp?shopid=<%=shopid%>&acURL=<%=acURL%>", "addnewbrand", "width=600,height=400,scrollbars=yes,resizable=yes");
	}
	addnewbrand.focus();
}

//전체 선택
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}


function CkDisPrice(){
	CkDisOrOrg(true);
}

function CkOrgPrice(){
	CkDisOrOrg(false);
}

//원가 할인가 적용
function CkDisOrOrg(isDisc){
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

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if(isDisc==true){
				    <% if isRate=0  then %>
					    //할인율이 0이면
					    //frm.iDSPrice.value = frm.saleprice.value;
				    <% else %>
						frm.iDSPrice.value = frm.saleprice.value;	// 할인가
					<% end if %>

                    frm.sellpricemargin.value = Math.round(((frm.shopitemprice.value-frm.iDSPrice.value)/frm.shopitemprice.value)*10000)/100;  //할인마진
                    if (frm.calcuMarginValue.value!=0){
            	        frm.iDBPrice.value              = Math.floor(Math.round(frm.iDSPrice.value*(100.0 - frm.calcuMarginValue.value)/100)/10)*10; //할인매입가
            	    }
            	    if (frm.calcushopMarginValue.value!=0){
                        frm.idsaleshopsupplycash.value  = Math.floor(Math.round(frm.iDSPrice.value*(100.0 - frm.calcushopMarginValue.value)/100)/10)*10; //할인매장공급가
                    }

            		frm.iDSMargin.value = Math.round(((frm.iDSPrice.value-frm.iDBPrice.value)/frm.iDSPrice.value)*10000.0)/100;
                    frm.idsshopmargin.value = Math.round(((frm.iDSPrice.value-frm.idsaleshopsupplycash.value)/frm.iDSPrice.value)*10000.0)/100;
                    frm.saleItemStatus.value = 7;
				}else{
				    //원가적용
					frm.iDSPrice.value = frm.orgPrice.value;
					frm.iDBPrice.value = frm.orgSupplyPrice.value;
					frm.idsaleshopsupplycash.value = frm.orgshopbuyprice.value;
					frm.iDSMargin.value= frm.orgMarginValue.value;
					frm.idsshopmargin.value= frm.orgshopMarginValue.value;
					frm.saleItemStatus.value = 9;
				}
			}
		}
	}
}

//선택상품 저장
function saveArr(){
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

	frmarr.itemid.value = "";
	frmarr.itemgubun.value = "";
	frmarr.itemoption.value = "";
	frmarr.sailyn.value = "";
	frmarr.iDSPrice.value ="";
	frmarr.iDBPrice.value ="";
	frmarr.idsaleshopsupplycash.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				//check Not AvaliValue
				if (!IsDigit(frm.iDSPrice.value)){
					alert('할인 판매가는 숫자만 가능합니다.');
					frm.iDSPrice.focus();
					return;
				}

				if (frm.iDSPrice.value<1){
					alert('할인 판매가를 정확히 입력하세요.');
					frm.iDSPrice.focus();
					return;
				}

				if (!IsDigit(frm.iDBPrice.value)){
					alert('할인 매입가는 숫자만 가능합니다.');
					frm.iDBPrice.focus();
					return;
				}

				//정산조건 텐바이텐특정, 업체특정, 출고특정일 경우 정산플래그 특정
				if (frm.comm_cd.value=='B011' || frm.comm_cd.value=='B012' || frm.comm_cd.value=='B013'){
					if (frm.iDBPrice.value<1){
						alert('할인 매입가를 정확히 입력하세요.');
						frm.iDBPrice.focus();
						return;
					}
				}

				if (!IsDigit(frm.idsaleshopsupplycash.value)){
					alert('할인 매장공급가는 숫자만 가능합니다.');
					frm.idsaleshopsupplycash.focus();
					return;
				}

				//정산조건 텐바이텐특정, 업체특정, 출고특정일 경우 정산플래그 특정
				if (frm.comm_cd.value=='B011' || frm.comm_cd.value=='B012' || frm.comm_cd.value=='B013'){
					if (frm.idsaleshopsupplycash.value<1){
						alert('할인 매장공급가를 정확히 입력하세요.');
						frm.idsaleshopsupplycash.focus();
						return;
					}
				}

				frmarr.itemid.value = frmarr.itemid.value + frm.itemid.value + ","
				frmarr.itemgubun.value = frmarr.itemgubun.value + frm.itemgubun.value + ","
				frmarr.itemoption.value = frmarr.itemoption.value + frm.itemoption.value + ","
				//if (frm.sailyn[0].checked){
					//frmarr.sailyn.value = frmarr.sailyn.value + "Y" + ","
				//}else{
					//frmarr.sailyn.value = frmarr.sailyn.value + "N" + ","
				//}
				frmarr.iDSPrice.value = frmarr.iDSPrice.value + frm.iDSPrice.value + ","
				frmarr.iDBPrice.value = frmarr.iDBPrice.value + frm.iDBPrice.value + ","
				frmarr.idsaleshopsupplycash.value = frmarr.idsaleshopsupplycash.value + frm.idsaleshopsupplycash.value + ","
				frmarr.point_ratearr.value = frmarr.point_ratearr.value + frm.point_rate.value + ","
				frmarr.saleItemStatus.value = frmarr.saleItemStatus.value + frm.saleItemStatus.value+","
				frmarr.saleitem_idxarr.value = frmarr.saleitem_idxarr.value + frm.saleitem_idx.value + ","
			}
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		frmarr.submit();
	}
}

function delArr(){
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

	frmdel.itemid.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				frmdel.itemid.value = frmdel.itemid.value + frm.itemid.value + ","
				frmdel.itemgubun.value = frmdel.itemgubun.value + frm.itemgubun.value + ","
				frmdel.itemoption.value = frmdel.itemoption.value + frm.itemoption.value + ","
				frmdel.saleitem_idxarr.value = frmdel.saleitem_idxarr.value + frm.saleitem_idx.value + ","
			}
		}
	}

	var ret = confirm('삭제하시겠습니까?');

	if (ret){
		frmdel.submit();
	}

}

//할인가 수정시 공급가 입력
function reCALDisPrice(fid){
    var frm = document["frmBuyPrc_" + fid];

	if(frm.iDSPrice.value>0) {
	    frm.sellpricemargin.value = Math.round(((frm.shopitemprice.value-frm.iDSPrice.value)/frm.shopitemprice.value)*10000)/100;  //할인마진
        if (frm.calcuMarginValue.value!=0){
	        frm.iDBPrice.value              = Math.round(frm.iDSPrice.value*(100.0 - frm.calcuMarginValue.value)/100); //할인매입가
        }

        if (frm.calcushopMarginValue.value!=0){
            frm.idsaleshopsupplycash.value  = Math.round(frm.iDSPrice.value*(100.0 - frm.calcushopMarginValue.value)/100); //할인매입가
        }

		frm.iDSMargin.value = Math.round(((frm.iDSPrice.value-frm.iDBPrice.value)/frm.iDSPrice.value)*10000.0)/100;
        frm.idsshopmargin.value = Math.round(((frm.iDSPrice.value-frm.idsaleshopsupplycash.value)/frm.iDSPrice.value)*10000.0)/100;
	} else {
	    frm.sellpricemargin.value = 0;
		frm.iDSMargin.value = 0;

	}
}

// 마진율 재계산
function reCALbyPrice(fid) {
	var frm = document["frmBuyPrc_" + fid];

	if(frm.iDSPrice.value>0) {
		frm.iDSMargin.value = Math.round(((frm.iDSPrice.value-frm.iDBPrice.value)/frm.iDSPrice.value)*10000.0)/100;
		//frm.sellpricemargin.value = Math.round(((frm.shopitemprice.value-frm.iDSPrice.value)/frm.shopitemprice.value)*10000.0)/100;
	} else {
		frm.iDSMargin.value = 0;
		//frm.sellpricemargin.value = 0;
	}
}

// 샵 판매 마진율 재계산
function reCALbyshopPrice(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSPrice.value>0) {
		frm.idsshopmargin.value = Math.round(((frm.iDSPrice.value-frm.idsaleshopsupplycash.value)/frm.iDSPrice.value)*10000.0)/100;
	} else {
		frm.idsshopmargin.value = 0;
	}
}

// 매입가 재계산
function reCALbyMargin(fid) {
	var frm = document["frmBuyPrc_" + fid];

	//정산조건 텐바이텐특정, 업체특정, 출고특정일 경우 정산플래그 특정
	if (frm.comm_cd.value=='B011' || frm.comm_cd.value=='B012' || frm.comm_cd.value=='B013'){
		if(frm.iDSMargin.value>0) {
			frm.iDBPrice.value = Math.round(frm.iDSPrice.value*(1-(frm.iDSMargin.value/100)));
		} else {
			frm.iDBPrice.value = frm.iDSPrice.value;
		}
	}else{
		alert('매입상품은 매입가를 변경하실수 없습니다.');
	}
}

// 샵판매가 재계산
function reCALbyshopMargin(fid) {
	var frm = document["frmBuyPrc_" + fid];

	//정산조건 텐바이텐특정, 업체특정, 출고특정일 경우 정산플래그 특정
	if (frm.comm_cd.value=='B011' || frm.comm_cd.value=='B012' || frm.comm_cd.value=='B013'){
		if(frm.iDSMargin.value>0) {
			frm.idsaleshopsupplycash.value = Math.round(frm.iDSPrice.value*(1-(frm.idsshopmargin.value/100)));
		} else {
			frm.idsaleshopsupplycash.value = frm.iDSPrice.value;
		}
	}else{
		alert('매입상품은 매장공급가를 변경하실수 없습니다.');
	}
}

</script>

<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmdummi">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sc" value="<%=sCode%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : <% drawSelectBoxDesignerwithName "designer",designer  %>
		&nbsp;&nbsp;
		* 상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;&nbsp;
		* 상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
		&nbsp;&nbsp;
		<input type="checkbox" name="adminvspos" value="ON" <% if adminvspos = "ON" then response.write " checked" %>>어드민포스할인가격서로상이
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:frmdummi.submit();">
	</td>
</tr>
</form>
</table>
<!---- /검색 ---->
<Br>
※ 현재 포스에서 할인중인 상품이여도 할인기간만 겹치지 않는다면, 등록이 가능합니다.
매장에서 오픈이나, 종료로 <font color="red">즉시반영</font>을 원하시는경우, <font color="red">반드시 실시간적용</font> 버튼을 누르세요.
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td bgcolor="<%= adminColor("tabletop") %>">할인코드</td>
	<td bgcolor="#FFFFFF"><%=sCode%></td>
	<td bgcolor="<%= adminColor("tabletop") %>">할인명</td>
	<td bgcolor="#FFFFFF"><%=sTitle%></td>
	<td bgcolor="<%= adminColor("tabletop") %>">매장</td>
	<td bgcolor="#FFFFFF"><%= shopid %></td>
</tr>
<tr align="center">
	<td bgcolor="<%= adminColor("tabletop") %>">상태</td>
	<td bgcolor="#FFFFFF"><%=fnGetCommCodeArrDesc_off(arrsalestatus,isStatus)%></td>
	<td bgcolor="<%= adminColor("tabletop") %>">기간</td>
	<td bgcolor="#FFFFFF"><%=dSDay%> ~ <%=dEDay%></td>
	<td bgcolor="<%= adminColor("tabletop") %>">포인트적립</td>
	<td bgcolor="#FFFFFF">
		<%= point_rate %>
	</td>
</tr>
<!--<tr align="center">
	<td bgcolor="<%= adminColor("tabletop") %>">이벤트코드(그룹)</td>
	<td bgcolor="#FFFFFF">
		<%' If eCode > 0 THEN %>
			<%'= eCode %>
			<%' If egCode > 0 THEN %>
				(<%'= egCode %>)
			<%' END IF %>
		<%' END IF %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>"></td>
	<td bgcolor="#FFFFFF"></td>
	<td bgcolor="<%= adminColor("tabletop") %>"></td>
	<td bgcolor="#FFFFFF">
	</td>
</tr>-->
<tr align="center">
	<td bgcolor="<%= adminColor("tabletop") %>">할인율</td>
	<td bgcolor="#FFFFFF"><%=isRate%>%</td>
	<td bgcolor="<%= adminColor("tabletop") %>">매입마진</td>
	<td bgcolor="#FFFFFF">
		<%=fnGetCommCodeArrDesc_off(arrsalemargin,isMargin)%>

		<%IF isMargin = 5 THEN%>
			,&nbsp;할인마진율 <font color="blue"><%=isMValue%>%</font>
		<%END IF%>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">샾판매마진</td>
	<td bgcolor="#FFFFFF">
		<%=fnGetCommCodeArrDesc_off(arrsalemargin,sale_shopmargin)%>
		<%IF sale_shopmargin = 5 THEN%>
			,&nbsp;할인마진율 <font color="blue"><%=sale_shopmarginvalue%>%</font>
		<%END IF%>
	</td>
</tr>
</table>

<Br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<input type="button" value="할인적용" onClick="CkDisPrice();" class="button">
		<input type="button" value="원가적용" onClick="CkOrgPrice();" class="button">
		<input type=button value="선택상품수정" onClick="saveArr()" class="button">
		<input type=button value="선택상품삭제" onClick="delArr()" class="button">
    </td>
    <td align="right">
		<% if eCode <> "0" then %>
			<input type="button" value="상품추가(검색)" <% if geteventcheckitem(eCode) then%>onclick="addnewItem(<%=eCode%>,<%=egCode%>);"<% else %>onclick="alert('먼저 이벤트에 상품을 넣어주세요');"<% end if %> class="button">
		<% else %>
			<input type="button" value="상품추가(검색)" onclick="addnewItem(<%=eCode%>,<%=egCode%>);" class="button">
			<input type="button" value="상품추가(브랜드일괄)" onclick="addnewbrand(<%=eCode%>,<%=egCode%>);" class="button">
		<% end if %>
		&nbsp;&nbsp;
		<input type="button" value="뒤로가기" onClick="location.href='salelist.asp?menupos=<%=menupos%>&shopid=<%=shopid%>';" class="button">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="left">검색결과 : <b><%=clsSaleItem.ftotalcount%></b>&nbsp;&nbsp;페이지 : <b><%=page%> / <%=clsSaleItem.FTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ck_all" onclick="SelectCk(this)"></td>
	<td>이미지</td>
	<td>상품번호</td>
	<td>상품명</td>
	<td>브랜드</td>
	<td>센터계약<Br>정산구분</td>
	<td>할인상태</td>
	<td>판매가</td>
	<td>매입가<br>매장공급가</td>
	<td>매입마진<br>매장공급마진</td>
    <td>할인판매가</td>
    <td>할인마진</td>
	<td>할인매입가<br>할인매장공급가</td>
	<td>할인매입마진<br>할인매장공급마진</td>
	<td>
		적립포인트
		<!--<br>사용중지-->
	</td>
</tr>
<%
Dim mSPrice, mSBPrice, iSaleMargin, iOrgMargin ,mSshopPrice,mSBshopPrice, iSaleshopMargin , iOrgshopMargin
	iSaleMargin=0
	iOrgMargin = 0
	iSaleshopMargin = 0
	iOrgshopMargin = 0

Dim icalcuMargin,icalcuShopMargin ''공급가 계산을 위한 마진율

IF clsSaleItem.fresultcount > 0 THEN

For i = 0 To clsSaleItem.fresultcount -1

'//정산조건 텐바이텐특정, 업체특정, 출고특정일 경우 정산플래그 특정
itemcontract = ""
if clsSaleItem.FItemList(i).fcomm_cd="B011" or clsSaleItem.FItemList(i).fcomm_cd="B012" or clsSaleItem.FItemList(i).fcomm_cd="B013" then
	itemcontract="W"
else
	itemcontract="M"
end if

mSPrice = fix(fix( clsSaleItem.FItemList(i).forgsellprice - (clsSaleItem.FItemList(i).forgsellprice*(isRate/100)) )/10)*10	' 할인가
mSBPrice = fnSetSaleSupplyPrice(isMargin,isMValue,clsSaleItem.FItemList(i).forgsellprice,clsSaleItem.FItemList(i).fshopsuplycash,mSPrice,clsSaleItem.FItemList(i).fcomm_cd)
if mSPrice<>0 then iSaleMargin = 100-fix(mSBPrice/mSPrice*10000)/100
if clsSaleItem.FItemList(i).forgsellprice<>0 then iOrgMargin= 100-fix(clsSaleItem.FItemList(i).fshopsuplycash/clsSaleItem.FItemList(i).forgsellprice*10000)/100

mSshopPrice = clsSaleItem.FItemList(i).forgsellprice - (clsSaleItem.FItemList(i).forgsellprice*(isRate/100))
mSBshopPrice = fnSetSaleSupplyPrice(sale_shopmargin,sale_shopmarginvalue,clsSaleItem.FItemList(i).forgsellprice,clsSaleItem.FItemList(i).fshopbuyprice,mSshopPrice,clsSaleItem.FItemList(i).fcomm_cd)
if mSshopPrice<>0 then iSaleshopMargin = 100-fix(mSBshopPrice/mSshopPrice*10000)/100
if clsSaleItem.FItemList(i).forgsellprice<>0 then iOrgshopMargin= 100-fix(clsSaleItem.FItemList(i).fshopbuyprice/clsSaleItem.FItemList(i).forgsellprice*10000)/100

'/할인마진
sellpricemargin = 0
if clsSaleItem.FItemList(i).fshopitemprice<>0 then
	sellpricemargin = 100-fix(clsSaleItem.FItemList(i).fsaleprice/clsSaleItem.FItemList(i).fshopitemprice*10000)/100
end if

icalcuMargin = clsSaleItem.FItemList(i).getCalcuMargin(isMargin,isMValue)
icalcuShopMargin = clsSaleItem.FItemList(i).getCalcuShopMargin(sale_shopmargin,sale_shopmarginvalue)
%>
<form name="frmBuyPrc_<%=i%>" >
<input type="hidden" name="saleitem_idx" value="<%= clsSaleItem.FItemList(i).fsaleitem_idx %>">
<input type="hidden" name="comm_cd" value="<%= clsSaleItem.FItemList(i).fcomm_cd %>">
<input type="hidden" name="itemid" value="<%=clsSaleItem.FItemList(i).fitemid%>">
<input type="hidden" name="itemgubun" value="<%=clsSaleItem.FItemList(i).fitemgubun%>">
<input type="hidden" name="itemoption" value="<%=clsSaleItem.FItemList(i).fitemoption%>">
<input type="hidden" name="saleprice" value='<%= mSPrice %>'>
<input type="hidden" name="saleshopprice" value="<%=mSshopPrice%>">
<input type="hidden" name="salesupplyprice" value="<%=mSBPrice%>">
<input type="hidden" name="saleshopsupplyprice" value="<%=mSBshopPrice%>">
<input type="hidden" name="salemargin" value="<%=iSaleMargin%>">
<input type="hidden" name="saleshopmargin" value="<%=iSaleshopMargin%>">
<input type="hidden" name="orgPrice" value="<%=clsSaleItem.FItemList(i).forgsellprice%>">
<input type="hidden" name="orgSupplyPrice" value="<%=clsSaleItem.FItemList(i).fshopsuplycash%>">
<input type="hidden" name="orgshopbuyprice" value="<%=clsSaleItem.FItemList(i).fshopbuyprice%>">
<input type="hidden" name="orgMarginValue" value="<%=iOrgMargin%>">
<input type="hidden" name="orgshopMarginValue" value="<%=iOrgshopMargin%>">
<input type="hidden" name="calcuMarginValue" value="<%=icalcuMargin%>">
<input type="hidden" name="calcushopMarginValue" value="<%=icalcuShopMargin%>">
<input type="hidden" name="saleItemStatus" value="<%=clsSaleItem.FItemList(i).fsaleItem_status%>">
<input type="hidden" name="shopitemprice" value="<%=clsSaleItem.FItemList(i).fshopitemprice%>">

<% IF cint(clsSaleItem.FItemList(i).fsaleItem_status) >= 8 THEN %>
	<tr align="center" bgcolor="#c1c1c1" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#c1c1c1';>
<% else %>
	<tr align="center" bgcolor="#ffffff" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% end if %>

    <td width=25><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
    <td width=50>
    	<%IF clsSaleItem.FItemList(i).fsmallimage <> "" THEN%>
    		<img src="<%=clsSaleItem.FItemList(i).fsmallimage%>" width=50 height=50>
    	<%END IF%>
    </td>
    <td width=80>
    	<%=clsSaleItem.FItemList(i).fitemgubun%><%=CHKIIF(clsSaleItem.FItemList(i).fitemid>=1000000,Format00(8,clsSaleItem.FItemList(i).fitemid),Format00(6,clsSaleItem.FItemList(i).fitemid))%><%=clsSaleItem.FItemList(i).fitemoption%>
    </td>
    <td align="left">
    	<br><%=db2html(clsSaleItem.FItemList(i).fshopitemname)%>
    </td>
    <td>
    	<%=db2html(clsSaleItem.FItemList(i).fmakerid)%>
    </td>
    <td width=100>
    	<%= fnColor(clsSaleItem.FItemList(i).fcentermwdiv,"mw") %>&nbsp;<%= clsSaleItem.FItemList(i).fcentermwdiv %>
    	<Br><%= clsSaleItem.FItemList(i).fcomm_name %>
    </td>
    <td width=150>
    	포스할인 :
    	<% if isStatus = "8" and clsSaleItem.FItemList(i).fsaleyn = "Y" then %>
    		<font color="red"><%=clsSaleItem.FItemList(i).fsaleyn%> (타할인)</font>
    	<% elseif clsSaleItem.FItemList(i).fsaleyn = "Y" then %>
    		<font color="red"><%=clsSaleItem.FItemList(i).fsaleyn%></font>
    	<% else %>
    		<font color="blue"><%=clsSaleItem.FItemList(i).fsaleyn%></font>
    	<% end if %>

    	<Br>어드민상태(<%=clsSaleItem.FItemList(i).fsaleItem_status%>) : <font color="blue"><%=fnGetCommCodeArrDesc_off(arrsalestatus,clsSaleItem.FItemList(i).fsaleItem_status)%></font>
    </td>
    <td align="right" width=80>
    	<%=formatnumber(clsSaleItem.FItemList(i).fshopitemprice,0)%><!--현재판매가-->
    </td>
    <td align="right" width=80>
    	<%=formatnumber(clsSaleItem.FItemList(i).fshopsuplycash,0)%><!--현재매입가-->
    	<br><%=formatnumber(clsSaleItem.FItemList(i).fshopbuyprice,0)%><!--현재샾판매가-->
    </td>
    <td width=80>
    	<% if clsSaleItem.FItemList(i).fshopitemprice<>0 then %><!--현재마진율-->
			<%= 100-fix(clsSaleItem.FItemList(i).fshopsuplycash/clsSaleItem.FItemList(i).fshopitemprice*10000)/100 %>%
		<% end if %>

    	<% if clsSaleItem.FItemList(i).fshopitemprice<>0 then %><!--현재샾판매마진율-->
			<br><%= 100-fix(clsSaleItem.FItemList(i).Fshopbuyprice/clsSaleItem.FItemList(i).fshopitemprice*10000)/100 %>%
		<% end if %>
	</td>
	<% IF cint(clsSaleItem.FItemList(i).fsaleItem_status) = 8 or  cint(clsSaleItem.FItemList(i).fsaleItem_status) = 9 THEN %>
		<td align="right" width=180>
			<input type="text" name="iDSPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALDisPrice('<%=i%>')">
		</td>
        <td align="right" width=70>
			<input type="text" class="text_ro" name="sellpricemargin" value="0" style="text-align:right;" size="4">%
		</td>
        <td align="right" width=100>
			<input type="text" name="iDBPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyPrice('<%=i%>')">
			<br><input type="text" name="idsaleshopsupplycash" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyshopPrice('<%=i%>')">
		</td>
	    <td align="right" width=110>
	    	<input type="text" name="iDSMargin" value="0" style="text-align:right;" size="4" >%
	    	<br><input type="text" name="idsshopmargin" value="0" style="text-align:right;" size="4" >%
	    </td>
	    <td align="right" width=70>
			<input type="text" name="point_rate" value="0" style="text-align:right;" size="4" maxlength=4 readonly>%
	    </td>
	<% ELSE %>
	    <td align="right" width=180>
	    	<input type="text" name="iDSPrice" size="6" maxlength="9" value="<%= clsSaleItem.FItemList(i).fsaleprice %>" style="text-align:right;" onkeyup="reCALDisPrice('<%=i%>')">
	    	<%
	    	if clsSaleItem.FItemList(i).fsale_status = "6" and clsSaleItem.FItemList(i).fsaleItem_status = "6" and clsSaleItem.FItemList(i).fpossaleprice <> "" then
	    	%>
	    		<font color="red"><br>현재포스적용가격 : <%=formatnumber(clsSaleItem.FItemList(i).fpossaleprice,0)%></font>
	    	<% end if %>
	    </td>
        <td align="right" width=70>
			<input type="text" class="text_ro" name="sellpricemargin" value="<%= sellpricemargin %>" style="text-align:right;" size="4">%
		</td>
        <td align="right" width=140>
	    	<input type="text" name="iDBPrice" size="6" maxlength="9" value="<%=clsSaleItem.FItemList(i).fsalesupplycash%>" style="text-align:right;" onkeyup="reCALbyPrice('<%=i%>')" <% if itemcontract="M" then %>readonly style="background-color:c1c1c1;"<% end if %>>
	    	<br><input type="text" name="idsaleshopsupplycash" size="6" maxlength="9" value="<%=clsSaleItem.FItemList(i).fsaleshopsupplycash%>" style="text-align:right;" onkeyup="reCALbyshopPrice('<%=i%>')" <% if itemcontract="M" then %>readonly style="background-color:c1c1c1;"<% end if %>>

	    	<% if itemcontract="M" then %>
	    		<Br><font color="red">정산완료.매입가변경불가</font>
	    	<% end if %>
	    </td>
	    <td align="right" width=110>
	    	<%
	    	if clsSaleItem.FItemList(i).fsaleprice<>0 then smargin= 100-fix(clsSaleItem.FItemList(i).fsalesupplycash/clsSaleItem.FItemList(i).fsaleprice*10000)/100
	    	if clsSaleItem.FItemList(i).fsaleprice<>0 then sshopmargin= 100-fix(clsSaleItem.FItemList(i).fsaleshopsupplycash/clsSaleItem.FItemList(i).fsaleprice*10000)/100
	    	%>
			<input type="text" class="text_ro" name="iDSMargin" value="<%=smargin%>" style="text-align:right;" size="4" >%
			<br><input type="text" class="text_ro" name="idsshopmargin" value="<%=sshopmargin%>" style="text-align:right;" size="4" >%
	    </td>
	    <td align="right" width=70>
	    	<!--<input type="text" name="point_rate" value="<%'=clsSaleItem.FItemList(i).fpoint_rate%>" style="text-align:right;" size="4" maxlength=4 readonly>%-->
	    	<input type="text" name="point_rate" value="<%=clsSaleItem.FItemList(i).fpoint_rate%>" style="text-align:right;" size="4" maxlength=4>%
	    </td>
	<% END IF %>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if clsSaleItem.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= clsSaleItem.StartScrollPage-1 %>&sc=<%=sCode%>&<%=para%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + clsSaleItem.StartScrollPage to clsSaleItem.StartScrollPage + clsSaleItem.FScrollCount - 1 %>
			<% if (i > clsSaleItem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(clsSaleItem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&sc=<%=sCode%>&<%=para%>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if clsSaleItem.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&sc=<%=sCode%>&<%=para%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% ELSE %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
</tr>
<% END IF %>
<form name="frmarr" method="post" action="saleItemPRoc.asp">
	<input type="hidden" name="mode" value="U">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="sC" value="<%=sCode%>">
	<input type="hidden" name="iC" value="<%=page%>">
	<input type="hidden" name="itemid" value="">
	<input type="hidden" name="itemgubun" value="">
	<input type="hidden" name="itemoption" value="">
	<input type="hidden" name="sailyn" value="">
	<input type="hidden" name="iDSPrice" value="">
	<input type="hidden" name="iDBPrice" value="">
	<input type="hidden" name="idsaleshopsupplycash" value="">
	<input type="hidden" name="point_ratearr" value="">
	<input type="hidden" name="saleItemStatus" value="">
	<input type="hidden" name="saleStatus" value="<%=isStatus%>">
	<input type="hidden" name="designer" value="<%=designer%>">
	<input type="hidden" name="saleitem_idxarr" value="">
</form>
<form name="frmdel" method="post" action="saleItemPRoc.asp">
	<input type="hidden" name="mode" value="D">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="sC" value="<%=sCode%>">
	<input type="hidden" name="itemid" value="">
	<input type="hidden" name="itemgubun" value="">
	<input type="hidden" name="itemoption" value="">
	<input type="hidden" name="saleitem_idxarr" value="">
</form>
</table>

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<input type="button" value="할인적용" onClick="CkDisPrice();" class="button">
		<input type="button" value="원가적용" onClick="CkOrgPrice();" class="button">
		<input type=button value="선택상품수정" onClick="saveArr()" class="button">
		<input type=button value="선택상품삭제" onClick="delArr()" class="button">
    </td>
    <td align="right">
		<% if eCode <> "0" then %>
			<input type="button" value="상품추가(검색)" <% if geteventcheckitem(eCode) then%>onclick="addnewItem(<%=eCode%>,<%=egCode%>);"<% else %>onclick="alert('먼저 이벤트에 상품을 넣어주세요');"<% end if %> class="button">
		<% else %>
			<input type="button" value="상품추가(검색)" onclick="addnewItem(<%=eCode%>,<%=egCode%>);" class="button">
			<input type="button" value="상품추가(브랜드일괄)" onclick="addnewbrand(<%=eCode%>,<%=egCode%>);" class="button">
		<% end if %>
		&nbsp;&nbsp;
		<input type="button" value="뒤로가기" onClick="location.href='salelist.asp?menupos=<%=menupos%>&shopid=<%=shopid%>';" class="button">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<%
set clsSaleItem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->