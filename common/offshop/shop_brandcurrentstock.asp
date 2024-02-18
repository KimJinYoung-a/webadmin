<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인매장 브랜드별 재고 파악
' History : 2011.08.10 서동석 생성
'			2011.10.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim NowDate
NowDate = Left(CStr(Now()),10)

dim shopid, makerid, centermwdiv, itembarcode, usingyn, research, NoZeroStock, showminusOnly
dim itemgubun, itemid, itemoption
dim pagesize
shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
centermwdiv  = RequestCheckVar(request("centermwdiv"),10)
itembarcode  = RequestCheckVar(request("itembarcode"),32)
usingyn      = RequestCheckVar(request("usingyn"),1)
research     = RequestCheckVar(request("research"),2)
NoZeroStock  = RequestCheckVar(request("NoZeroStock"),32)
showminusOnly  	= RequestCheckVar(request("showminusOnly"),32)
pagesize  		= RequestCheckVar(request("pagesize"),32)
''매장
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

''업체
if (C_IS_Maker_Upche) then
    makerid = session("ssBctid")
end if

if (itembarcode <> "") then
    if Not (fnGetItemCodeByPublicBarcode(itembarcode,itemgubun,itemid,itemoption)) then
        if Len(itembarcode)=12 then
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(Mid(itembarcode, 3, 6) + 0)
            itemoption  = Right(itembarcode, 4)
        elseif Len(itembarcode)=14 then
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(Mid(itembarcode, 3, 8) + 0)
            itemoption  = Right(itembarcode, 4)
        else
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(0)
            itemoption  = Right(itembarcode, 4)
        end if
    end if
end if

if (research="") and (usingyn="") then NoZeroStock="on"
if (pagesize = "") then
	pagesize = "100"
end if

dim oOffStock
set oOffStock = new CShopItemSummary
oOffStock.FCurrPage 		= 1
oOffStock.FPageSize 		= pagesize
oOffStock.FRectShopID       = shopid
oOffStock.FRectMakerID      = makerid
oOffStock.FRectCenterMwDiv  = centermwdiv
oOffStock.FRectIsUsing      = usingyn
oOffStock.FRectNoZeroStock  = NoZeroStock
oOffStock.FRectShowMinusOnly  = showminusOnly
if (itembarcode <> "") then
    oOffStock.FRectItemGubun    = itemgubun
    oOffStock.FRectItemId       = itemid
    oOffStock.FRectItemOption   = itemoption
end if

if ((shopid<>"") and (makerid<>"")) or ((shopid<>"") and (itembarcode<>"")) or (showminusOnly <> "") then
    oOffStock.GetShopItemCurrentSummaryList
end if

dim i
dim totsysstock, totavailstock, totrealstock

dim IsUpcheWitakItem
if (makerid<>"") and (shopid<>"") then
    IsUpcheWitakItem = (GetShopBrandContract(shopid,makerid)="B012")
end if
%>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function popShopCurrentStock(shopid,itemgubun,itemid,itemoption){
    var popwin = window.open('/common/offshop/shop_itemcurrentstock.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popShopCurrentStock','width=900,height=600,resizable=yes,scrollbars=yes');
    popwin.focus();
}

function popOffItemEdit(ibarcode){
    <% if C_IS_SHOP then %>
        return;
    <% elseif C_IS_Maker_Upche then %>
        var popwin = window.open('/designer/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
        popwin.focus();
    <% else %>
	    var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	    popwin.focus();
	<% end if %>
}

function popOffErrInput(shopid,itemgubun,itemid,itemoption){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
        alert('권한이 없습니다. - 업체위탁 상품만 재고 수정 가능.');
        return; //업체위탁 상품인 경우?
    <% else %>
        var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popAdmOffrealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
	    popwin.focus();
	<% end if %>
}

function PopOFFBrandStockSheet(){

    var shopid = document.frm.shopid.value;
    var makerid = document.frm.makerid.value;
    var centermwdiv = "";//document.frm.centermwdiv.value;
    var usingyn= document.frm.usingyn.value;
	var NoZeroStock= document.frm.NoZeroStock.value;
	var pgSize= document.frm.pagesize.value;

    if ((shopid.length<1)||(makerid.length<1)){
        alert('먼저 매장과 브랜드를 선택후 출력해 주세요.');
        return;
    }

    var popwin;

    popwin = window.open('/common/pop_offbrandstockprint.asp?shopid=' + shopid + '&makerid=' + makerid + '&centermwdiv=' + centermwdiv + '&usingyn=' + usingyn + '&NoZeroStock=' + NoZeroStock + '&pagesize=' + pgSize ,'pop_offbrandstockprint','width=1200,height=600,scrollbars=yes,resizable=yes')
    popwin.focus();
}

function ForceALLZero(){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
        alert('권한이 없습니다. - 업체위탁 상품만 재고 수정 가능.');
        return;
    <% end if %>

    if (!confirm('전체 재고를 0으로 처리하는 기능입니다. \n\n퇴점 브랜드 등에 사용 \n\n계속 하시겠습니까?')){
        return;
    }

    var frm = document.frmArr;

    if (frm.cksel.length){
        for (i=0;i<frm.cksel.length;i++){
            frm.cksel[i].checked = true;
            frm.Arrrealstock[i].value=0;
            frm.Arrshoprealstock[i].value=0;
            CheckThis(i);
        }
    }else{
        frm.cksel.checked = true;
        frm.Arrrealstock.value=0;
        frm.Arrshoprealstock.value=0;
        CheckThis(0);
    }

}

function RealStockInputArr(){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
        alert('권한이 없습니다. - 업체위탁 상품만 재고 수정 가능.');
        return;
    <% end if %>

    var frm = document.frmArr;
    var ischecked = false;
    var i = 0;
    var stockdate = frmStockDt.stockdate.value;

    if (!frm.cksel) return;

    if (frm.cksel.length){
        for (i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                ischecked = true;
                if (!IsInteger(frm.Arrrealstock[i].value)){
                    alert('정수만 가능합니다.');
                    frm.Arrrealstock[i].focus();
                    return;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            ischecked = true;
            if (!IsInteger(frm.Arrrealstock.value)){
                alert('정수만 가능합니다.');
                frm.Arrrealstock.focus();
                return;
            }
        }
    }

    if (!(ischecked)){
        alert('선택된 상품이 없습니다.');
        return;
    }

    if (confirm('실사 재고를 저장 하시겠습니까?')){
        frm.stockdate.value = stockdate;
        frm.submit();
    }
}

function CheckThis(i){
    var frm = document.frmArr;
    if (frm.cksel.length){
        frm.cksel[i].checked = true;
        AnCheckClick(frm.cksel[i]);
    }else{
        frm.cksel.checked = true;
        AnCheckClick(frm.cksel);
    }
}

function CheckSet(i, stockno) {
	var frm = document.frmArr;
	var realstock, logischulgo, logisreturn, shoprealstock;
    if (frm.cksel.length){
		realstock = frm.Arrrealstock[i];
		logischulgo = frm.Arrlogischulgo[i];
		logisreturn = frm.Arrlogisreturn[i];
		shoprealstock = frm.Arrshoprealstock[i];
    }else{
		realstock = frm.Arrrealstock;
		logischulgo = frm.Arrlogischulgo;
		logisreturn = frm.Arrlogisreturn;
		shoprealstock = frm.Arrshoprealstock;
    }

	if ((shoprealstock.value == "") || (shoprealstock.value*0 != 0)) {
		alert("잘못된 숫자입니다.[" + shoprealstock.value + "]");
		shoprealstock.value = stockno;
		return;
	}

	realstock.value = shoprealstock.value*1 + logischulgo.value*1 + logisreturn.value*1;
}

</script>

<script type="text/javascript">

//2013.04.08 한용민 추가(리스트가 길어질경우 해당내역이 무슨내역인지 몰라서 오래걸린다함. 리스트가 상단에 따라다님)
$(document).ready(function(){
	var currentPosition = parseInt($("#floating").css("top"));

	//스크롤시
	$(window).scroll(function() {
		//레이어 보이기
		$("#floating").show();

		//레이어 시작위치와 현재 레이어 위치를 계산해서 0.1초 간격으로 레이어 따라다님
		var position = $(window).scrollTop();
		windowCenterH = parseInt($(window).height()/2);
		$("#floating").stop().animate({"top":position+currentPosition+"px"},100);

		//현재 레이어 위치가 현재 창크기보다 클경우 레이어 숨김
		if (position+currentPosition < $(window).height()){
			$("#floating").hide();
		}
	});
});

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    <% if (C_IS_SHOP) then %>
		    <input type="hidden" name="shopid" value="<%= shopid %>">
		    매장 : <%= shopid %>
	    <% elseif (C_IS_Maker_Upche) then %>
		    <!-- 계약된 업체 -->
		    매장 : <% if (FALSE) then %><!-- drawSelectBoxOpenOffShop "shopid",shopid --> <!-- 2016/05/02 변경 문재 요청-->  <% end if %>
		    <% drawBoxDirectIpchulOffShopByMaker "shopid",shopid,makerid %>
	    <% else %>
	        매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp; <!-- drawSelectBoxOffShop -->
	    <% end if %>

	    <% if (C_IS_Maker_Upche) then %>
	        <input type="hidden" name="makerid" value="<%= makerid %>">
	    <% else %>
			브랜드 :
			<% drawSelectBoxDesignerwithName "makerid", makerid %> &nbsp;&nbsp;
		<% end if %>

		<!-- 카테고리 :  -->
		상품바코드 :
		<input type="text" class="text" name="itembarcode" value="<%= itembarcode %>" size="20" maxlength="32">
		<br>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		상품 사용구분 : <% drawSelectBoxUsingYN "usingyn", usingyn %> &nbsp;&nbsp;

		<input type="checkbox" name="NoZeroStock" <%= CHKIIF(NoZeroStock="on","checked","") %> > <font color="red">재고0인 상품 검색 안함.</font>
		&nbsp;
		<input type="checkbox" name="showminusOnly" <%= CHKIIF(showminusOnly="on","checked","") %> > 마이너스 재고만.</font>
		&nbsp;
		표시갯수 :
		<select class="select" name="pagesize">
			<option value="100" <%= CHKIIF(pagesize = "100", "selected", "") %>>100</option>
			<option value="500" <%= CHKIIF(pagesize = "500", "selected", "") %>>500</option>
			<option value="1000" <%= CHKIIF(pagesize = "1000", "selected", "") %>>1000</option>
			<option value="2000" <%= CHKIIF(pagesize = "2000", "selected", "") %>>2000</option>
		</select>
		<!--
		센터매입구분 :
		   <select class="select" name="centermwdiv">
           <option value="">전체</option>
           <option value="MW" <%= ChkIIF(centermwdiv="MW","selected","") %> >매입+위탁</option>
           <option value="W"  <%= ChkIIF(centermwdiv="W","selected","") %> >위탁</option>
           <option value="M"  <%= ChkIIF(centermwdiv="M","selected","") %> >매입</option>
           <option value="NULL" <%= ChkIIF(centermwdiv="NULL","selected","") %> >미지정</option>
           </select>
         -->
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br />* 최대 <%= pagesize %>개까지만 표시됩니다.

<!-- 액션 시작 -->
<form name="frmStockDt">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td align="left">
		<% if C_ADMIN_AUTH=true then %>
    	<!--
        <input type="button" class="button" value="브랜드 전체 새로고침" onclick="RefreshIpchulStock();">
        -->
        <% end if %>
    	<input type="button" class="button" name="stock_sheet_print" value="재고파악SHEET출력" onclick="javascript:PopOFFBrandStockSheet();">

    	<input type="button" class="button" value="전체 실사 재고 0 처리" onclick="javascript:ForceALLZero();">
	    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
            (업체위탁 계약 매장만 재고 수정 가능)
        <% end if %>
	</td>
	<td align="right">
	    <input type="text" class="text" name="stockdate" value="<%= NowDate %>" size=11 readonly ><a href="javascript:calendarOpen(frmStockDt.stockdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
		<input type="button" class="button" name="stock_sheet_print" value="선택 상품 실사재고 일괄입력" onclick="RealStockInputArr();">
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<form name="frmArr" method="post" action="/common/offshop/shop_stockrefresh_process.asp">
<input type="hidden" name="mode" value="ArrOfferrcheckupdate">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="stockdate" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20" ></td>
    <td width="30">구분</td>
	<td width="40">상품ID</td>
	<td width="40">옵션</td>
	<td width="50">이미지</td>
	<td>상품명<br>[옵션명]</td>
	<td>
		<% if (C_IS_Maker_Upche) or (C_IS_SHOP) then %>
	    매장<br>매입가
	    <% ELSE %>
	    본사<br>매입가
	    <% END IF %>
	</td>
	<td>판매가</td>
	<!-- td width="40">센터<br>매입<br>구분</td -->
	<td width="40">센터<br>입고</td>
	<td width="40">센터<br>반품</td>
	<td width="40">브랜드<br>입고</td>
	<td width="40">브랜드<br>반품</td>
    <td width="40">매장<br>판매</td>
    <td width="40">매장<br>반품</td>
    <td width="40" bgcolor="F4F4F4">시스템<br>총재고</td>
    <td width="40">총<br>실사<br>오차</td>
    <td width="40" bgcolor="F4F4F4">실사<br>재고</td>
    <td width="60">재고금액<br>(매입가*실사)</td>
	<td width="40">배송중</td>
	<td width="40">반품중</td>
	<td width="40" bgcolor="F4F4F4">매장<br>재고<br>(현재)</td>
	<td width="40">재고<br>입력</td>
	<td width="40">샘플</td>
    <td width="40" bgcolor="F4F4F4">유효<br>재고</td>
    <td width="30">사용<br>여부</td>
    <td width="40">개별<br>입력</td>
</tr>
<tr id="floating" style="position:absolute;margin:0px 0px; top:0px; display:none;" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20" ></td>
    <td width="30">구분</td>
	<td width="40">상품ID</td>
	<td width="40">옵션</td>
	<td width="50">이미지</td>
	<td>상품명<br>[옵션명]</td>
	<td>
	    <% if (C_IS_Maker_Upche) or (C_IS_SHOP) then %>
	    매장<br>매입가
	    <% ELSE %>
	    본사<br>매입가
	    <% END IF %>
	</td>
	<td>판매가</td>
	<!-- td width="40">센터<br>매입<br>구분</td -->
	<td width="40">센터<br>입고</td>
	<td width="40">센터<br>반품</td>
	<td width="40">브랜드<br>입고</td>
	<td width="40">브랜드<br>반품</td>
    <td width="40">매장<br>판매</td>
    <td width="40">매장<br>반품</td>
    <td width="40" bgcolor="F4F4F4">시스템<br>총재고</td>
    <td width="40">총<br>실사<br>오차</td>
    <td width="40" bgcolor="F4F4F4">실사<br>재고</td>
    <td width="60">재고금액<br>(매입가*실사)</td>
	<td width="40">배송중</td>
	<td width="40">반품중</td>
	<td width="40" bgcolor="F4F4F4">매장<br>재고<br>(현재)</td>
	<td width="40">재고<br>입력</td>
    <td width="40">샘플</td>
    <td width="40" bgcolor="F4F4F4">유효<br>재고</td>
    <td width="30">사용<br>여부</td>
    <td width="40">개별<br>입력</td>
</tr>

<% if oOffStock.FResultCount<1 then %>

<tr align="center" bgcolor="#FFFFFF" height="30">
    <% if (shopid="") and (makerid="") then %>
    	<td colspan="26" >[ 매장 및 브랜드를 선택 하세요. ]</td>
    <% else %>
    	<td colspan="26" >[ 검색 결과가 없습니다. ]</td>
    <% end if %>
</tr>

<% else %>

<%
Dim vTotCenterIp, vTotCenterBan, vTotBrandIp, vTotBrandBan, vTotMaeIp, vTotMaeBan
dim totalBuycash ,totalshopitemprice , totallogicsipgono , totalbrandreipgono ,totalresellno, totalbuycashSum
dim totalsysstockNo , totalerrrealcheckno , totalAvailStock, totalerrsampleitemno, totalRealStock
for i=0 to oOffStock.FResultCount - 1
%>
<%
IF (C_IS_FRN_SHOP) THEN
	totalBuycash = totalBuycash + oOffStock.FItemList(i).GetOfflineSuplycash
ELSE
	totalBuycash = totalBuycash + oOffStock.FItemList(i).GetOfflineBuycash
END IF

if Not IsNull(oOffStock.FItemList(i).fshopitemprice) then
	totalshopitemprice = totalshopitemprice + oOffStock.FItemList(i).fshopitemprice
end if

totallogicsipgono = totallogicsipgono + oOffStock.FItemList(i).Flogicsipgono + oOffStock.FItemList(i).Flogicsreipgono
totalbrandreipgono = totalbrandreipgono + oOffStock.FItemList(i).Fbrandipgono + oOffStock.FItemList(i).Fbrandreipgono
totalresellno = totalresellno + oOffStock.FItemList(i).Fsellno + oOffStock.FItemList(i).Fresellno
totalsysstockNo = totalsysstockNo + oOffStock.FItemList(i).FsysstockNo
totalerrrealcheckno = totalerrrealcheckno + oOffStock.FItemList(i).Ferrrealcheckno

totalRealStock       = totalRealStock + oOffStock.FItemList(i).Frealstockno
totalerrsampleitemno = totalerrsampleitemno + oOffStock.FItemList(i).Ferrsampleitemno
totalAvailStock = totalAvailStock + oOffStock.FItemList(i).getAvailStock

vTotCenterIp	= vTotCenterIp + oOffStock.FItemList(i).Flogicsipgono
vTotCenterBan	= vTotCenterBan + oOffStock.FItemList(i).Flogicsreipgono
vTotBrandIp		= vTotBrandIp + oOffStock.FItemList(i).Fbrandipgono
vTotBrandBan	= vTotBrandBan + oOffStock.FItemList(i).Fbrandreipgono
vTotMaeIp		= vTotMaeIp + oOffStock.FItemList(i).Fsellno
vTotMaeBan		= vTotMaeBan + oOffStock.FItemList(i).Fresellno

IF (C_IS_FRN_SHOP) THEN
	totalbuycashSum = totalbuycashSum + oOffStock.FItemList(i).Frealstockno*oOffStock.FItemList(i).GetOfflineSuplycash
ELSE
	totalbuycashSum = totalbuycashSum + oOffStock.FItemList(i).Frealstockno*oOffStock.FItemList(i).GetOfflineBuycash
END IF
%>
	<% if oOffStock.FItemList(i).Fisusing="Y" then %>
		<tr bgcolor="#FFFFFF" align="center">
    <% else %>
		<tr bgcolor="#FFFFFF" align="center">
    <% end if %>

        <td>
        	<input type="checkbox" name="cksel" onClick="AnCheckClick(this);" value="<%= i %>">
	        <input type="hidden" name="Arritemgubun" value="<%= oOffStock.FItemList(i).FItemGubun %>">
	        <input type="hidden" name="Arritemid" value="<%= oOffStock.FItemList(i).FItemID %>">
	        <input type="hidden" name="Arritemoption" value="<%= oOffStock.FItemList(i).FItemOption %>">
			<input type="hidden" name="Arrrealstock" value="<%= oOffStock.FItemList(i).Frealstockno %>">
			<input type="hidden" name="Arrlogischulgo" value="<%= oOffStock.FItemList(i).Flogischulgo %>">
			<input type="hidden" name="Arrlogisreturn" value="<%= oOffStock.FItemList(i).Flogisreturn %>">
        </td>
        <td><%= oOffStock.FItemList(i).FItemGubun %></td>
    	<td>
    	    <% if (C_ADMIN_USER or C_IS_Maker_Upche) then %>
    	    <a href="javascript:popOffItemEdit('<%= oOffStock.FItemList(i).getBarcode %>');"><%= oOffStock.FItemList(i).Fitemid %></a>
    	    <% else %>
    	    <%= oOffStock.FItemList(i).Fitemid %>
    	    <% end if %>
    	</td>
    	<td><%= oOffStock.FItemList(i).FItemOption %></td>
    	<td><img src="<%= oOffStock.FItemList(i).GetImageSmall %>" width=50 height=50> </td>
    	<td align="left">
          	<a href="javascript:popShopCurrentStock('<%= oOffStock.FItemList(i).FShopid %>','<%= oOffStock.FItemList(i).Fitemgubun %>','<%= oOffStock.FItemList(i).FItemID %>','<%= oOffStock.FItemList(i).FItemOption %>');"><%= oOffStock.FItemList(i).FShopitemname %></a>
          	<% if oOffStock.FItemList(i).FShopitemoptionName <>"" then %>
          		<br>
          		<font color="blue">[<%= oOffStock.FItemList(i).FShopitemoptionName %>]</font>
          	<% end if %>
        </td>
    	<td>
    	    <% if (C_IS_Maker_Upche) or (C_IS_SHOP) then %>
    	    <%= FormatNumber(oOffStock.FItemList(i).GetOfflineSuplycash,0) %>
    	    <% ELSE %>
    	    <%= FormatNumber(oOffStock.FItemList(i).GetOfflineBuycash,0) %>
    	    <% END IF %>
    	</td>
    	<td>
			<% if Not IsNull(oOffStock.FItemList(i).fshopitemprice) then  %>
			<%= FormatNumber(oOffStock.FItemList(i).fshopitemprice,0) %>
			<% end if  %>
		</td>
        <!-- td><%= fnColor(oOffStock.FItemList(i).FCenterMwdiv,"mw") %></td -->
    	<td><%= oOffStock.FItemList(i).Flogicsipgono %></td>
    	<td><%= oOffStock.FItemList(i).Flogicsreipgono %></td>
    	<td><%= oOffStock.FItemList(i).Fbrandipgono %></td>
    	<td><%= oOffStock.FItemList(i).Fbrandreipgono %></td>
    	<td><%= oOffStock.FItemList(i).Fsellno %></td>
    	<td><%= oOffStock.FItemList(i).Fresellno %></td>
    	<td bgcolor="F4F4F4"><b><%= oOffStock.FItemList(i).FsysstockNo %></b></td>
    	<td><%= oOffStock.FItemList(i).Ferrrealcheckno %></td>
    	<td bgcolor="F4F4F4"><b><font color="<%= ChkIIF(oOffStock.FItemList(i).Frealstockno<0,"#FF0000","#000000") %>"><%= oOffStock.FItemList(i).Frealstockno %></font></b></td>
    	<td >
    		<% if oOffStock.FItemList(i).GetOfflineSuplycash<>"" and oOffStock.FItemList(i).Frealstockno<>"" then %>
	    	    <% if (C_IS_Maker_Upche) or (C_IS_SHOP) then %>
					<%= FormatNumber(oOffStock.FItemList(i).GetOfflineSuplycash*oOffStock.FItemList(i).Frealstockno,0) %>
	    	    <% ELSE %>
					<%= FormatNumber(oOffStock.FItemList(i).GetOfflineBuycash*oOffStock.FItemList(i).Frealstockno,0) %>
	    	    <% END IF %>
			<% END IF %>
    	</td>
		<td <%= CHKIIF(oOffStock.FItemList(i).Flogischulgo<>0, "style='font-weight: bold; color: red;'", "")%>><%= oOffStock.FItemList(i).Flogischulgo %></td>
		<td <%= CHKIIF(oOffStock.FItemList(i).Flogisreturn<>0, "style='font-weight: bold; color: red;'", "")%>><%= oOffStock.FItemList(i).Flogisreturn %></td>
		<td><b><%= oOffStock.FItemList(i).getShopRealStockNoExc %></b></td>
		<td><input type="text" class="text" name="Arrshoprealstock" value="<%= oOffStock.FItemList(i).getShopRealStock %>" size="4" maxlength="4" AUTOCOMPLETE="off" style="text-align=center" onKeyDown="CheckThis('<%= i %>');" onFocusOut="CheckSet(<%= i %>, <%= oOffStock.FItemList(i).getShopRealStock %>)"></td>
    	<td><%= oOffStock.FItemList(i).Ferrsampleitemno %></td>
    	<td ><b><font color="<%= ChkIIF(oOffStock.FItemList(i).getAvailStock<0,"#FF0000","#000000") %>"><%= oOffStock.FItemList(i).getAvailStock %></font></b></td>
    	<td>
    	    <% if oOffStock.FItemList(i).Fisusing="N" then %>
    	    <strong><%= oOffStock.FItemList(i).Fisusing %></strong>
    	    <% else %>
    	    <%= oOffStock.FItemList(i).Fisusing %>
    	    <% end if %>
    	</td>
    	<td>
    		<input type="button" class="button" value="입력" onclick="popOffErrInput('<%= shopid %>','<%= oOffStock.FItemList(i).Fitemgubun %>','<%= oOffStock.FItemList(i).Fitemid %>','<%= oOffStock.FItemList(i).Fitemoption %>');">
    	</td>
    </tr>
<% next %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan=6></td>
	<td>
		<% if totalBuycash <> "" then %>
			<%= FormatNumber(totalBuycash,0) %>
		<% end if %>
	</td>
	<td>
		<% if totalshopitemprice <> "" then %>
			<%= FormatNumber(totalshopitemprice,0) %>
		<% end if %>
	</td>
	<!-- td></td -->
	<td>
		<% if vTotCenterIp <> "" then %>
			<%= FormatNumber(vTotCenterIp,0) %>
		<% end if %>
	</td>
    <td>
    	<% if vTotCenterBan <> "" then %>
    		<%= FormatNumber(vTotCenterBan,0) %>
    	<% end if %>
    </td>
	<td>
		<% if vTotBrandIp <> "" then %>
			<%= FormatNumber(vTotBrandIp,0) %>
		<% end if %>
	</td>
    <td>
    	<% if vTotBrandBan <> "" then %>
    		<%= FormatNumber(vTotBrandBan,0) %>
    	<% end if %>
    </td>
    <td>
    	<% if vTotMaeIp <> "" then %>
    		<%= FormatNumber(vTotMaeIp,0) %>
    	<% end if %>
    </td>
    <td>
    	<% if vTotMaeBan <> "" then %>
    		<%= FormatNumber(vTotMaeBan,0) %>
    	<% end if %>
    </td>
    <td>
    	<% if totalsysstockNo <> "" then %>
    		<%= FormatNumber(totalsysstockNo,0) %>
    	<% end if %>
    </td>
    <td>
    	<% if totalerrrealcheckno <> "" then %>
    		<%= FormatNumber(totalerrrealcheckno,0) %>
    	<% end if %>
    </td>
    <td>
    	<% if totalRealStock <> "" then %>
    		<%= FormatNumber(totalRealStock,0) %>
    	<% end if %>
    </td>
    <td>
    	<% if totalbuycashSum <> "" then %>
    		<%= FormatNumber(totalbuycashSum,0) %>
    	<% end if %>
    </td>
    <td></td>
	<td></td>
	<td></td>
	<td></td>
    <td>
    	<% if totalerrsampleitemno <> "" then %>
    		<%= FormatNumber(totalerrsampleitemno,0) %>
    	<% end if %>
    </td>
    <td>
    	<% if totalAvailStock <> "" then %>
    		<%= FormatNumber(totalAvailStock,0) %>
    	<% end if %>
    </td>
	<td></td>
    <td></td>
</tr>
<% end if %>

</table>
</form>

<%
set oOffStock = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
