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
osummarystock.FRectmakerid = makerid
osummarystock.FRectSearchType = "err"

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
    if (e.value*1 != 0){
        hL(e);
    }else{
        dL(e);
    }
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
        alert("검색된 상품이 없습니다.");
        return;
    <% else %>
        if (document.frm.itemid.value.length != 12) {
			alert("상품코드를 정확히 입력하세요.");
			return;
        }

		var frm = document.frm;
		var itembarcode = frm.itemid.value;

		for (var i = 0; ; i++) {
			var itemgubun = document.getElementById("itemgubun_" + i);
			var itemid = document.getElementById("itemid_" + i);
			var itemoption = document.getElementById("itemoption_" + i);

			var itemno = document.getElementById("itemno_" + i);
			var itemmaxno = document.getElementById("itemmaxno_" + i);

			if (itemgubun == undefined) {
				alert("상품이 목록에 없습니다. 다른 브랜드이거나, 오차등록이 되어 있지 않습니다.");
				break;
			}

			if ((itemgubun.value == itembarcode.substring(0,2)) && (itemid.value*1 == (1 * itembarcode.substring(2,8))) && (itemoption.value == itembarcode.substring(8))) {
				itemno.value = (1 * itemno.value) + 1;

				/*
				if ((1 * itemno.value) > (itemmaxno.value * -1)) {
					itemno.value = (itemmaxno.value * -1);
					alert("오차등록된 수량보다 수량이 큽니다. 먼저 오차등록을 하세요.");
				}
				*/

				hL(itemno);
				break;
			}
		}

        frm.itemid.select();
        frm.itemid.focus();
    <% end if %>
}

function SubmitCheckInsert(v) {
	var curridx = v.value;
	var itemno = document.getElementById("itemno_" + curridx);
	var itemmaxno = document.getElementById("itemmaxno_" + curridx);

	if (v.checked == true) {
		itemno.value = itemmaxno.value*-1;
	} else {
		itemno.value = 0;
	}
	checkhL(itemno);
}

function SubmitCheckInsertAll(v) {
	for (var i = 0;; i++) {
		var chk = document.getElementById("chk_" + i);
		if (chk == undefined) {
			return;
		}
		chk.checked = v.checked;
		SubmitCheckInsert(chk);
	}
}

function CheckInsert(f, maxvalue){
        alert(f.value);
        if (f.value = "") {
                return;
        }

        if (f.value < 0) {
                alert("오차<%= LorRText %>수량은 마이너스가 될수 없습니다.");
                f.value = 0;
                return;
        }

        if (f.value > (maxvalue * -1)) {
                alert("오차등록된 수량보다 수량이 큽니다. 먼저 오차등록을 하세요.");
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
    var pmwdiv = "";

    var frm = document.frm;

    frm.itemgubunarr.value = "";
    frm.itemidarr.value = "";
    frm.itemoptionarr.value = "";
    frm.itemnoarr.value = "";

	for (var i = 0; ; i++) {
		var itemgubun = document.getElementById("itemgubun_" + i);
		var itemid = document.getElementById("itemid_" + i);
		var itemoption = document.getElementById("itemoption_" + i);

		var itemno = document.getElementById("itemno_" + i);
		var mwdiv = document.getElementById("mwdiv_" + i);

		if (itemgubun == undefined) {
			break;
		}

		if (itemno.value*1 != 0) {
			if (pmwdiv == "") {
				pmwdiv = mwdiv.value;
			} else {
				// 반품의 경우
				/*
				if (pmwdiv != mwdiv.value) {
					alert("매입 속성이 다른제품을 같이 처리 할 수 없습니다.");
					return;
				}
				*/
			}

			frm.itemgubunarr.value = frm.itemgubunarr.value + itemgubun.value + "|";
			frm.itemidarr.value = frm.itemidarr.value + itemid.value + "|";
			frm.itemoptionarr.value = frm.itemoptionarr.value + itemoption.value + "|";
			frm.itemnoarr.value = frm.itemnoarr.value + itemno.value + "|";
		}
	}

	if (frm.itemgubunarr.value == "") {
        alert("<%= LorRText %>처리할 상품이 없습니다.");
        return;
    }

    if (confirm('<%= LorRText %> 내역서를 작성하시겠습니까?')){
        document.frm.method = "post";
        <% if (actType="actloss") then %>
        document.frm.mode.value = "lossarr";
        <% else %>
        document.frm.mode.value = "notused";
        <% end if %>
        document.frm.pmwdiv.value = pmwdiv;
        document.frm.action = "/common/erritem_re_input_process.asp";
        document.frm.submit();
    }
}
</script>

<!-- 검색 시작 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" >
<tr bgcolor="#FFFFFF">
    <td>** 오차등록 상품 <strong>로스</strong> 처리</td>
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
			<input type="button" class="button_s" value="브랜드오차등록상품목록검색" onClick="SubmitSearchItem()">
		</td>
	</tr>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td>
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="12" size="14" onKeyPress="if (event.keyCode == 13) { SubmitInsert(); return false; }">
			<input type="button" class="button" value="로스추가" onclick="SubmitInsert()">
        	&nbsp;
			* 한종류의 브랜드만 일괄처리 가능합니다.
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
      <td width="20"><input type="checkbox" name="chkall" value="" onClick="SubmitCheckInsertAll(this);"></td>
      <td width="100">브랜드ID</td>
      <td width="40">매입<br>구분</td>
      <td width="25">구분</td>
      <td width="40">상품<br>코드</td>
      <td width="30">옵션</td>
      <td>아이템명</td>
      <td>옵션명</td>
      <td width="50">소비자가</td>
      <td width="40">오차<br>수량</td>
      <td width="40">출고<br>수량</td>
    </tr>
    <form name="frmlist" method=get action="" onsubmit="return false;">
<% for i=0 to osummarystock.FResultCount - 1 %>
    <tr align="center" bgcolor="<%= chkIIF(osummarystock.FItemList(i).FItemgubun="10","#FFFFFF","#EEEEEE") %>">
      <td><input type="checkbox" id="chk_<%= i %>" name="chk" value="<%= i %>" onClick="SubmitCheckInsert(this);"></td>
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
        <%= osummarystock.FItemList(i).Ferrrealcheckno %>
      </td>
      <input type="hidden" id="itemgubun_<%= i %>" name="itemgubun" value="<%= osummarystock.FItemList(i).FItemgubun %>">
      <input type="hidden" id="itemid_<%= i %>" name="itemid" value="<%= osummarystock.FItemList(i).FItemid %>">
      <input type="hidden" id="itemoption_<%= i %>" name="itemoption" value="<%= osummarystock.FItemList(i).FItemOption %>">
      <td>
        <input type="text" class="text" id="itemno_<%= i %>" name="itemno" value="0" size="3" onKeyUP="checkhL(this);" <%= chkIIF(osummarystock.FItemList(i).FItemgubun="10","","disabled") %> >
      </td>
      <input type="hidden" id="itemmaxno_<%= i %>" name="itemmaxno" value="<%= osummarystock.FItemList(i).Ferrrealcheckno %>" >
      <input type="hidden" id="mwdiv_<%= i %>" name="mwdiv" value="<%= osummarystock.FItemList(i).FMwdiv %>">
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