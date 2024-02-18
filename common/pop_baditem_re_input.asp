<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/baditemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim itemid, makerid, mode, actType

itemid  = requestCheckVar(request("itemid"),32)     '' length > 9
makerid = requestCheckVar(request("makerid"),32)
mode    = requestCheckVar(request("mode"),9)
actType = requestCheckVar(request("actType"),9)     '' actType="actloss" 로스처리 actType<>"actloss" 반품처리

dim osummarystock
set osummarystock = new CSummaryItemStock
if (Len(itemid) = 12) then
        osummarystock.FRectItemID =  Mid(itemid,3,6)
end if
if (Len(itemid) = 14) then
        osummarystock.FRectItemID =  Mid(itemid,3,8)
end if
osummarystock.FRectmakerid = makerid

if (makerid<>"") then
osummarystock.GetDailyErrItemListByBrand
end if

if (osummarystock.FResultCount > 0) then
        makerid = osummarystock.FItemList(0).Fmakerid
end if

dim i

dim LorRText
if (actType="actloss") then
    LorRText = "로스"
else
    LorRText = "반품"
end if

%>
<script language='javascript'>


function getOnLoad(){
	document.frm.itemid.focus();
	document.frm.itemid.select();
}

window.onload=getOnLoad;

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

function SubmitSearchByItemId() {
        if ((document.frm.itemid.value.length != 12) && (document.frm.itemid.value.length != 14)) {
                alert("상품코드를 정확히 입력하세요.");
                return;
        }
        document.frm.makerid.selectedIndex = 0;
        document.frm.submit();
}

function SubmitSearchItem() {
        if ((document.frm.itemid.value.length != 12) && (document.frm.itemid.value.length != 14)) {
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
        if ((document.frm.itemid.value.length != 12) && (document.frm.itemid.value.length != 14)) {
                alert("상품코드를 정확히 입력하세요.");
                return;
        }

        var e;
        var t;
    	var found = 0;
		var itemgubun = "";
		var itemid = "";
		var itemoption = "";

        e = document.frmlist.elements;
        t = document.frm.itemid.value;
	for (var i=0; i < e.length; i++){
		if (e[i].name == "itemgubun") {
			if (e[i].value != "10"){
				alert("오프 상품은 처리 할 수 없습니다.");
				return;
			}

			if (t.length == 12) {
				itemgubun = t.substring(0,2);
				itemid = (1 * t.substring(2,8));
				itemoption = t.substring(8);
			} else if (t.length == 14) {
				itemgubun = t.substring(0,2);
				itemid = (1 * t.substring(2,10));
				itemoption = t.substring(10);
			} else {
				alert("ERROR");
				return;
			}

			if ((e[i].value == itemgubun) && (e[i+1].value == itemid) && (e[i+2].value == itemoption)) {
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
    }else{
        frm.itemid.select();
        frm.itemid.focus();
    }
    <% end if %>
}

function CheckInsert(f, maxvalue){
        alert(f.value);
        if (f.value = "") {
                return;
        }

        if (f.value < 0) {
                alert("불량<%= LorRText %>수량은 마이너스가 될수 없습니다.");
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
                        alert("불량<%= LorRText %>수량은 마이너스가 될수 없습니다.");
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

        		        <% if (actType<>"actloss") then %>
        		        if (pmwdiv==""){
                		    pmwdiv = e[i+5].value;
                		}else{
                		    if (pmwdiv!=e[i+5].value){
                		        alert('매입 속성이 다른제품을 같이 처리 할 수 없습니다.');
                		        return;
                		    }
                		}
                		<% end if %>
        		}


		}
	}

	if (f.itemgubunarr.value == "") {
        alert("<%= LorRText %>처리할 상품이 없습니다.");
        return;
    }

    //매입속성이 다른내역은 같이 작성할 수 없음.


    if (confirm('<%= LorRText %> 내역서를 작성하시겠습니까?')){
        document.frm.method = "post";
        <% if (actType="actloss") then %>
        document.frm.mode.value = "lossarr";
        <% else %>
        document.frm.mode.value = "ipgoarr";
        <% end if %>
        document.frm.pmwdiv.value = pmwdiv;
        document.frm.action = "/common/baditem_re_input_process.asp";
        document.frm.submit();
    }
}
</script>

<!-- 검색 시작 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" >
<tr bgcolor="#FFFFFF">
    <td>** <%= chkIIF(actType="actloss"," 불량 상품 <strong>로스</strong> 처리 "," 불량 상품 <strong>반품</strong> 처리 ") %></td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="" onsubmit="return false;">
	<input type="hidden" name="actType" value="<%= actType %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="brandid" value="<%= makerid %>">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="pmwdiv" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드명 : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
			<input type="button" class="button_s" value="브랜드불량상품목록검색" onClick="SubmitSearchItem()">
		</td>
	</tr>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td>
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="14" size="14" onKeyPress="if (event.keyCode == 13) { SubmitInsert(); return false; }">
			<input type="button" class="button" value="<%= chkIIF(actType="actloss"," 로스추가 "," 반품추가 ") %>" onclick="SubmitInsert()">
			<!--
			&nbsp;
        	<input type="button" value=" 브랜드검색 " onclick="SubmitSearchByBrand()">&nbsp;&nbsp;<input type="button" value=" 상품코드검색 " onclick="SubmitSearchByItemId()"><br>
        	-->
        	&nbsp;
			* 한종류의 브랜드만 일괄반품입고가 가능합니다.
		</td>
		<td align="right">
			<input type="button" class="button" value="전체저장" onclick="SubmitUpdateAll()">
		</td>
	</tr>
	</form>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="100">브랜드ID</td>
      <td width="40">매입<br>구분</td>
      <td width="25">구분</td>
      <td width="40">상품<br>코드</td>
      <td width="30">옵션</td>
      <td>아이템명</td>
      <td>옵션명</td>
      <td width="50">소비자가</td>
      <td width="40">불량<br>수량</td>
      <td width="40">반품<br>수량</td>
    </tr>
    <form name="frmlist" method=get action="" onsubmit="return false;">
<% for i=0 to osummarystock.FResultCount - 1 %>
    <tr align="center" bgcolor="<%= chkIIF(osummarystock.FItemList(i).FItemgubun="10","#FFFFFF","#EEEEEE") %>">
      <td><%= osummarystock.FItemList(i).Fmakerid %></td>
      <td>
        <% if osummarystock.FItemList(i).FItemgubun="10" then %>
        <font color="<%= mwdivColor(osummarystock.FItemList(i).FMwdiv) %>"><%= osummarystock.FItemList(i).GetMwDivName %></font>
        <% end if %>
      </td>
      <td><%= osummarystock.FItemList(i).FItemgubun %></td>
      <td><%= osummarystock.FItemList(i).FItemid %></td>
      <td><%= osummarystock.FItemList(i).FItemoption %></td>
      <td align="left"><%= osummarystock.FItemList(i).FItemname %></td>
      <td align="left"><%= osummarystock.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= formatnumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
      <td>
        <%= osummarystock.FItemList(i).Ferrbaditemno %>
      </td>
      <input type="hidden" name="itemgubun" value="<%= osummarystock.FItemList(i).FItemgubun %>">
      <input type="hidden" name="itemid" value="<%= osummarystock.FItemList(i).FItemid %>">
      <input type="hidden" name="itemoption" value="<%= osummarystock.FItemList(i).FItemOption %>">
      <td>
        <input type="text" class="text" name="itemno" value="0" size="3" onKeyUP="checkhL(this);" <%= chkIIF(osummarystock.FItemList(i).FItemgubun="10","","disabled") %> >
      </td>
      <input type="hidden" name="itemmaxno" value="<%= osummarystock.FItemList(i).Ferrbaditemno %>" >
      <input type="hidden" name="mwdiv" value="<%= osummarystock.FItemList(i).FMwdiv %>">
    </tr>
   	<% next %>
<% if osummarystock.FResultCount = 0 then %>
    <tr align="center" bgcolor="#FFFFFF">
      <td colspan="11" align="center">검색된 상품이 없습니다.</td>
    </tr>
<% end if %>
    </form>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->


<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
