<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/baditemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim makerid,mode, searchtype
makerid = request("makerid")
mode = request("mode")
searchtype = request("searchtype")

searchtype = "bad"

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectmakerid = makerid
osummarystock.FRectSearchType = searchtype

if (makerid<>"") then
    osummarystock.GetDailyErrItemListByBrand
else
    osummarystock.GetDailyErrBadItemListByBrandGroup
end if

dim i

%>
<script language='javascript'>
function PopBadItemReInput(makerid){
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid,'pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopBadItemLossInput(makerid){
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid + '&actType=actloss','pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

//function getOnLoad(){
//	document.frm.itemid.focus();
//	document.frm.itemid.select();
//}
//
//window.onload=getOnLoad;

function checkhL(e){
    if (e.value*1>0){
        hL(e);
    }else{
        dL(e);
    }
}

function SubmitSearchByBrand() {
        if (document.frm.makerid.selectedIndex == 0) {
                alert("브랜드를 선택하세요.");
                return;
        }
        document.frm.itemid.value = "";
        document.frm.submit();
}

function SubmitSearchByBrandNew(makerid) {
	document.frm.makerid.value = makerid;
	document.frm.submit();
}

function SubmitSearchByItemId() {
        if (document.frm.itemid.value.length != 12) {
                alert("상품코드를 정확히 입력하세요.");
                return;
        }
        document.frm.makerid.selectedIndex = 0;
        document.frm.submit();
}

function SubmitSearchItem() {
        if (document.frm.itemid.value.length != 12) {
                if (document.frm.makerid.selectedIndex == 0) {
                        alert("브랜드 또는 상품코드를 입력하세요.");
                        return;
                }
                document.frm.itemid.value = "";
                document.frm.submit();
        } else {
                document.frm.makerid.selectedIndex = 0;
                document.frm.submit();
        }
}

function SubmitInsert(){
    <% if (osummarystock.FResultCount < 1) then %>
        SubmitSearchByItemId();
    <% else %>
        if (document.frm.itemid.value.length != 12) {
                alert("상품코드를 정확히 입력하세요.");
                return;
        }

        var e;
        var t;
        var found = 0;

        e = document.frmlist.elements;
        t = document.frm.itemid.value;
	for (var i=0; i < e.length; i++){
		if (e[i].name == "itemgubun") {
        		if ((e[i].value == t.substring(0,2)) && (e[i+1].value == (1 * t.substring(2,8))) && (e[i+2].value == t.substring(8))) {
        			e[i+3].value = (1 * e[i+3].value) + 1;

        			if ((1 * e[i+3].value) > (e[i+4].value * -1)) {
        			        e[i+3].value = (e[i+4].value * -1);
        			        alert("불량등록된 수량보다 수량이 큽니다. 먼저 불량등록을 하세요.");
                                }

        			found = 1;
        			hL(e[i+3]);
        			break;
        		}
		}
	}
	if (found == 0) {
        alert("상품이 목록에 없습니다. 다른 브랜드이거나, 불량등록이 되어 있지 않습니다.");
    }
    <% end if %>
}

function CheckInsert(f, maxvalue){
        alert(f.value);
        if (f.value = "") {
                return;
        }

        if (f.value < 0) {
                alert("불량반품수량은 마이너스가 될수 없습니다.");
                f.value = 0;
                return;
        }

        if (f.value > (maxvalue * -1)) {
                alert("불량등록된 수량보다 수량이 큽니다. 먼저 불량등록을 하세요.");
                f.value = (maxvalue * -1);
                return;
        }


}

function SubmitList(){
	window.open('/common/pop_item_search.asp','pop_item_search','width=900,height=600');
}


function ReActItems(itemgubunarr,
                    itemarr,
                    itemoptionarr,
                    sellcasharr,
                    suplycasharr,
                    buycasharr,
                    itemnoarr,
                    itemnamearr,
                    itemoptionnamearr,
                    designerarr,
                    mwdivarr)
{
        document.frm.itemgubunarr.value = itemgubunarr;
        document.frm.itemidarr.value = itemarr;
        document.frm.itemoptionarr.value = itemoptionarr;
        document.frm.itemnoarr.value = itemnoarr;

        document.frm.method = "post";
        document.frm.mode.value = "arrinsert";
        document.frm.action = "do_bad_item_input.asp";
        document.frm.submit();

        return true;
}





function SubmitUpdateAll(){
    var e;
    var f;
    var found = 0;
    var pmwdiv = "";

    e = document.frmlist.elements;
    f = document.frm;

    f.itemgubunarr.value = "";
    f.itemidarr.value = "";
    f.itemoptionarr.value = "";
    f.itemnoarr.value = "";

	for (var i=0; i < e.length; i++){
		if (e[i].name == "itemgubun") {
        		if ((e[i+3].value * 0) != 0) {
    		        alert("수량이 잘못 입력되었습니다.");
    		        e[i+3].focus();
    		        e[i+3].select();
    		        return;
                }

                if (e[i+3].value == "") {
                        alert("수량이 잘못 입력되었습니다.");
    		        e[i+3].focus();
    		        e[i+3].select();
                        return;
                }

                if (e[i+3].value < 0) {
                        alert("불량반품수량은 마이너스가 될수 없습니다.");
    		        e[i+3].focus();
    		        e[i+3].select();
                        return;
                }

                if (e[i+3].value > (e[i+4].value * -1)) {
                        alert("불량등록된 수량보다 수량이 큽니다. 먼저 불량등록을 하세요.");
    		        e[i+3].focus();
    		        e[i+3].select();
                        return;
                }

        		if ((e[i+3].value * 1) != 0) {
        		        f.itemgubunarr.value = f.itemgubunarr.value + e[i].value + "|";
        		        f.itemidarr.value = f.itemidarr.value + e[i+1].value + "|";
        		        f.itemoptionarr.value = f.itemoptionarr.value + e[i+2].value + "|";
        		        f.itemnoarr.value = f.itemnoarr.value + e[i+3].value + "|";

        		        if (pmwdiv==""){
                		    pmwdiv = e[i+5].value;
                		}else{
                		    if (pmwdiv!=e[i+5].value){
                		        alert('매입 속성이 다른제품을 같이 처리 할 수 없습니다.');
                		        return;
                		    }
                		}
        		}


		}
	}

	if (f.itemgubunarr.value == "") {
        alert("반품처리할 상품이 없습니다.");
        return;
    }

    //매입속성이 다른내역은 같이 작성할 수 없음.


    if (confirm('반품 내역서를 작성하시겠습니까?')){
        document.frm.method = "post";
        document.frm.mode.value = "ipgoarr";
        document.frm.pmwdiv.value = pmwdiv;
        document.frm.action = "/common/do_bad_item_re_input.asp";
        document.frm.submit();
    }
}

function ChangePage(v) {
	var frm = document.frm;

	if (v == "bad") {
		frm.action = "baditem_return_list.asp";
	} else {
		frm.action = "erritem_loss_list.asp";
	}

	frm.submit();
}

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;
			<input type="radio" name="searchtype" value="bad" <% if (searchtype = "bad") then %>checked<% end if %> onClick="ChangePage('bad')" > 불량상품
			<input type="radio" name="searchtype" value="err" <% if (searchtype = "err") then %>checked<% end if %> onClick="ChangePage('err')"> 오차등록상품
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<% if makerid<>"" then %>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="불량상품반품" onclick="PopBadItemReInput('<%= makerid %>')" border="0">
			&nbsp;
        	<input type="button" class="button" value="불량상품로스처리" onclick="PopBadItemLossInput('<%= makerid %>')" border="0">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= osummarystock.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30">구분</td>
		<td width="50">상품코드</td>
		<td width="40">옵션</td>
		<td width="50">이미지</td>
    	<td width="100">브랜드ID</td>

		<td>아이템명</td>
		<td>옵션명</td>
		<td width="40">계약<br>구분</td>

		<td width="50">소비자가</td>
		<td width="40">불량<br>수량</td>
    </tr>

	<% for i=0 to osummarystock.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td><%= osummarystock.FItemList(i).FItemid %></td>
		<td><%= osummarystock.FItemList(i).FItemoption %></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" ></td>
    	<td><%= osummarystock.FItemList(i).Fmakerid %></td>

		<td align="left"><%= osummarystock.FItemList(i).FItemname %></td>
		<td align="left"><%= osummarystock.FItemList(i).FItemOptionName %></td>
		<td><%= osummarystock.FItemList(i).GetMwDivName %></td>

		<td align="right"><%= formatnumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td><%= osummarystock.FItemList(i).Ferrbaditemno %></td>
    </tr>
    <% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		</td>
	</tr>
</table>

<% else %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= osummarystock.FResultCount %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">브랜드</td>
		<td width="100">불량상품수On</td>
		<td width="100">불량상품수Off</td>
		<td >&nbsp;</td>
	</tr>
	<% for i=0 to osummarystock.FResultCount-1 %>
	<tr bgcolor="#FFFFFF" height="30">
	    <td><a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td align="center"><%= osummarystock.FItemList(i).FOnCnt %></td>
	    <td align="center"><%= osummarystock.FItemList(i).FOffCnt %></td>
	    <td align="left">
			<input type="button" class="button" value="불량상품반품" onclick="PopBadItemReInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
			&nbsp;
        	<input type="button" class="button" value="불량상품로스처리" onclick="PopBadItemLossInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
	    </td>
	</tr>
	<% next %>
</table>
<% end if %>

<p>




<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
