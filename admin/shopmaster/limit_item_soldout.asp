<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->

<%

dim designerid, itemid,  sellyn, isusing
dim SearchMode

designerid  = request("designerid")
itemid      = request("itemid")
sellyn      = request("sellyn")
isusing     = request("isusing")
SearchMode  = request("SearchMode")

if ((request("research") = "") and (isusing = "")) then
        isusing = "on"
        SearchMode = "S1"
end if


dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FPageSize=300
osummarystock.FRectMakerid = designerid
osummarystock.FRectItemID = itemid
osummarystock.FRectOnlyIsUsing = isusing
osummarystock.FRectSearchMode = SearchMode
osummarystock.GetCurrentStockByOnlineBrandLimitSoldout

dim i

%>


<script language='javascript'>
function CheckThisRow(comp){
    var frm = comp.form;
    frm.cksel.checked = true;
    AnCheckClick(frm.cksel);
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function changecontent(){
	// nothing
}

function Research(page){
	frm.page.value = page;
	frm.submit();
}

function CheckNSellDispYN(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}
    upfrm.sellyn.value = "";
    upfrm.itemid.value = "";
	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;

					if (frm.sellyn[0].checked){
					    alert('한정 품절인 상품을 판매할 수 없습니다.');
					    frm.sellyn[0].focus();
					    return;
						//upfrm.sellyn.value = upfrm.sellyn.value + "|" + "Y";
					}else if (frm.sellyn[1].checked){
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "S";
					}else{
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "N";
					}
                    /*
					if (frm.dispyn[0].checked){
						upfrm.dispyn.value = upfrm.dispyn.value + "|" + "Y";
					}else{
						upfrm.dispyn.value = upfrm.dispyn.value + "|" + "N";
					}
					*/
				}
			}
		}
		frm.submit();
	}
}
</script>

<!-- 헤더에 포함 예정 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("menubar") %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>한정판매 품절 관리</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			한정판매 상품중 품절된 상품에 대한 정보입니다.
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>
<!-- 헤더에 포함 예정 끝 -->

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	브랜드 : <% drawSelectBoxDesignerwithName "designerid",designerid %>&nbsp;
        	상품코드 : <input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">&nbsp;
        	<input type=checkbox name="isusing" value="on" <% if isusing="on" then response.write "checked" %> >사용상품만


        	<input type="radio" name="SearchMode" value="Y0" <%= ChkIIF(SearchMode="Y0","checked","") %> > 판매Y, 한정Y, 한정0&nbsp;&nbsp;
          	<input type="radio" name="SearchMode" value="S1" <%= ChkIIF(SearchMode="S1","checked","") %> > 판매S, 한정Y, 한정1이상&nbsp;&nbsp;

        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a><br>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frmttl" onsubmit="return false;">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	검색결과 : <b><%= FormatNumber(osummarystock.FresultCount,0) %></b> (최대 : <%= osummarystock.FPageSize %>)
        </td>
        <td align="right">
        	<input type="button" value="전체선택" onClick="AnSelectAllFrame(true)">&nbsp;<input type="button" value="선택상품저장" onClick="CheckNSellDispYN()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>선택</td>
		<td width="50">이미지</td>
		<td width="70">브랜드</td>
		<td width="40">상품<br>코드</td>
		<td>상품명<br>(옵션명)</td>
		<td width="35">배송<br>구분</td>
        <td width="35">전체<br>입고<br>반품</td>
        <td width="35">전체<br>판매<br>반품</td>
        <td width="35">전체<br>출고<br>반품</td>
        <td width="35">기타<br>출고<br>반품</td>
<!--    <td width="35">시스템<br>재고</td>	-->
		<td width="35">총<br>불량</td>
<!--    <td width="35">유효<br>재고</td>	-->
        <td width="35">총<br>실사<br>오차</td>
        <td width="35">실사<br>재고</td>
        <td width="35">총<br>상품<br>준비</td>
        <td width="35">재고<br>파악<br>재고</td>
        <td width="35">ON<br>결제<br>완료</td>
        <td width="35">ON<br>주문<br>접수</td>
        <td width="35">한정<br>비교<br>재고</td>
<!--    <td width="35">적정<br>한정<br>재고</td>	-->
        <td width="50">전시<br>여부</td>
		<td width="50">판매<br>여부</td>
		<td width="60">한정<br>여부</td>
		<td width="35">품절<br>여부</td>
		<td width="35">단종<br>여부</td>
    </tr>
<% for i=0 to osummarystock.FresultCount-1 %>
	<form name="frmBuyPrc_<%= osummarystock.FItemList(i).FItemID %>" method="post" onSubmit="return false;" action="dolimitsoldset.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemid" value="<%= osummarystock.FItemList(i).FItemID %>">
	<% if osummarystock.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align="center">
    <% else %>
    <tr bgcolor="#EEEEEE" align="center">
    <% end if %>
    	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <%= ChkIIf (osummarystock.FItemList(i).FItemOptionName <> "","disabled","") %> ></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
		<td align="left">
          <%= osummarystock.FItemList(i).FMakerID %>
        </td>
		<td>
          <a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemID %>');"><%= osummarystock.FItemList(i).FItemID %></a>
        </td>
		<td align="left">
          <a href="javascript:PopItemDetail('<%= osummarystock.FItemList(i).FItemID %>','<%= osummarystock.FItemList(i).FItemOption %>')"><%= osummarystock.FItemList(i).FItemName %></a>
        <% if (osummarystock.FItemList(i).FItemOptionName <> "") then %>
          <br><font color="blue">(<%= osummarystock.FItemList(i).FItemOptionName %>)</font>
        <% end if %>
        </td>
        <td><%= osummarystock.FItemList(i).GetMwDivName %></td>
		<td><%= osummarystock.FItemList(i).Ftotipgono %></td>
		<td><%= -1*osummarystock.FItemList(i).Ftotsellno %></td>
		<td><%= osummarystock.FItemList(i).Foffchulgono + osummarystock.FItemList(i).Foffrechulgono %></td>
        <td><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></td>
<!--    <td><%= osummarystock.FItemList(i).Ftotsysstock %></td>	-->
        <td><%= osummarystock.FItemList(i).Ferrbaditemno %></td>
<!--    <td><%= osummarystock.FItemList(i).Favailsysstock %></td>	-->
        <td><%= osummarystock.FItemList(i).Ferrrealcheckno %></td>
        <td><b><%= osummarystock.FItemList(i).Frealstock %></b></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv5 + osummarystock.FItemList(i).Foffconfirmno %></td>
        <td><b><%= osummarystock.FItemList(i).GetCheckStockNo %></b></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv4 %></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv2 %></td>
        <td><b><%= osummarystock.FItemList(i).GetLimitStockNo %></b></td>
<!--        <td><b><%= round(osummarystock.FItemList(i).GetLimitStockNo * 0.95,0) %></b></td>	-->

<!--        <td><b><font color="red"><%= osummarystock.FItemList(i).GetLimitStockNo - osummarystock.FItemList(i).GetLimitStr %></font></b></td>	-->
        <td>

        </td>
        <td>
			<input type="radio" name="sellyn" value="Y" onClick="CheckThisRow(this);" <% if osummarystock.FItemList(i).Fsellyn="Y" then response.write "checked" %> >Y
			<input type="radio" name="sellyn" value="S" onClick="CheckThisRow(this);" <% if osummarystock.FItemList(i).Fsellyn="S" then response.write "checked" %> >S
			<input type="radio" name="sellyn" value="N" onClick="CheckThisRow(this);" <% if osummarystock.FItemList(i).Fsellyn="N" then response.write "checked ><font color=red>N</font>" else response.write ">N" %>
        </td>

        <td>
          	한정(<%= osummarystock.FItemList(i).GetLimitStr %>)
            <% if (osummarystock.FItemList(i).Foptlimityn = "Y") then %>
            <br>(<%= osummarystock.FItemList(i).Foptlimitno %>/<%= osummarystock.FItemList(i).Foptlimitsold %>)
            <% else %>
            <br>(<%= osummarystock.FItemList(i).FLimitNo %>/<%= osummarystock.FItemList(i).FLimitSold %>)
          	<% end if %>
        </td>
        <td><% if osummarystock.FItemList(i).IsSoldOut  then %><font color="red">품절</font><% end if %></td>
        <td>
            <% if osummarystock.FItemList(i).FDanjongyn="Y" then %>
            <font color="red">단종</font>
            <% elseif osummarystock.FItemList(i).FDanjongyn="S" then %>
            <font color="blue">일시<br>품절</font>
            <% else %>
            <% end if %>
        </td>
	</tr>
	</form>
<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<form name="frmArrupdate" method="post" action="dolimitsoldset.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="dispyn" value="">
<input type="hidden" name="sellyn" value="">
</form>

<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->