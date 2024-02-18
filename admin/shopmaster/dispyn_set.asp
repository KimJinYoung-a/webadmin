<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : [검토]전시판매관리
' History	:  이상구 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim designerid, itemid, dispyn, sellyn, isusing, diffdiv, mwdiv, i, vPurchasetype, tplgubun, isSellStart
dim dispCate, StockMwDiv
	tplgubun	= requestCheckvar(request("tplgubun"),32)
	designerid = requestCheckvar(request("designerid"),32)
	itemid = requestCheckvar(request("itemid"),10)
	dispyn = requestCheckvar(request("dispyn"),2)
	sellyn = requestCheckvar(request("sellyn"),2)
	isusing = requestCheckvar(request("isusing"),1)
	isSellStart = requestCheckvar(request("isSellStart"),1)
	diffdiv = requestCheckvar(trim(request("diffdiv")),32)
	mwdiv = requestCheckvar(request("mwdiv"),2)
	vPurchasetype 	= requestCheckvar(request("purchasetype"),3)
	dispCate = requestCheckvar(request("disp"),16)
	StockMwDiv  	= RequestCheckVar(request("StockMwDiv"),1)

if (diffdiv = "") then diffdiv = "sellN"
if ((request("research") = "") and (isusing = "")) then isusing = "Y"
if (request("research") = "") and tplgubun="" then tplgubun="3X"
if (request("research") = "") and isSellStart="" then isSellStart="Y"

dim osummarystock
set osummarystock = new CSummaryItemStock
	osummarystock.FRectMakerid = designerid
	osummarystock.FRectItemID = itemid
	osummarystock.FRectOnlyIsUsing = isusing
	osummarystock.FRectdiffdiv = diffdiv
	osummarystock.FRectMWDiv = mwdiv
	osummarystock.FRectTplGubun = tplgubun
	osummarystock.FRectIsSellStart = isSellStart
	osummarystock.FRectPurchasetype = vPurchasetype
	osummarystock.FRectDispCate		= dispCate
	osummarystock.FRectStockMwDiv = StockMwDiv
	osummarystock.GetCurrentStockByOnlineBrandDispSell

%>


<!-- obuyprice.GetDispYNSet 에서 새로운 디비 클래스로 교체함 -->

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
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

	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;

					if (frm.sellyn[0].checked){
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "Y";
					}else if (frm.sellyn[1].checked){
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "S";
					}else{
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "N";
					}

				}
			}
		}
		frm.submit();
	}
}
</script>


<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드: <% drawSelectBoxDesignerwithName "designerid",designerid %>&nbsp;
		&nbsp;
		* 상품코드: <input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">&nbsp;
		&nbsp;
		* 거래구분: <% drawSelectBoxMWU "mwdiv",mwdiv %>&nbsp;
		&nbsp;
		* 구매유형 : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
	</td>	
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!--<input type="radio" name="diffdiv" value="sellSlimit1" <% if (diffdiv = "sellSlimit1") then %>checked<% end if %>>일시품절/한정1이상&nbsp;//-->
		<!--<input type="radio" name="diffdiv" value="sellY0" <% if (diffdiv = "sellY0") then %>checked<% end if %>>판매/한정 1미만//-->
		<input type="radio" name="diffdiv" value="sellN" <% if (diffdiv = "sellN") then %>checked<% end if %>>품절/한정비교재고 1이상
		<input type="radio" name="diffdiv" value="sellSlimit2" <% if (diffdiv = "sellSlimit2") then %>checked<% end if %>>일시품절/한정비교재고 1이상
		&nbsp;&nbsp;
		* 사용여부: <% drawSelectBoxUsingYN "isusing", isusing %>
		&nbsp;
		* 3PL구분 : <% Call drawSelectBoxTPLGubun("tplgubun", tplgubun) %>
		&nbsp;
		* 판매전환여부 : <% Call drawSelectBoxUsingYN("isSellStart", isSellStart) %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 재고매입구분 :
		<select class="select" name="StockMwDiv">
			<option value="">선택</option>
			<option value="M" <% if (StockMwDiv = "M") then %>selected<% end if %> >M</option>
			<option value="W" <% if (StockMwDiv = "W") then %>selected<% end if %> >W</option>
			<option value="X" <% if (StockMwDiv = "X") then %>selected<% end if %> >기타</option>
		</select>
		&nbsp;
		* 전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<Br>

<!-- 액션 시작 -->
<form name="frmttl" onsubmit="return false;" style="margin:0px;">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
        <!--
			<input type="button" value="전체선택" onClick="AnSelectAllFrame(true)">&nbsp;<input type="button" value="선택상품저장" onClick="CheckNSellDispYN()">        </td>
        -->
	</td>
	<td align="right">	
	</td>
</tr>
<tr>
	<td align="left">
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->
<style>
th {
  background: #E6E6E6;
  position: sticky;
  top: 0;
  box-shadow: 0 0 1px 0 rgba(0, 0, 0, 0.4);
}
</style>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<thead>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="26">
		검색결과 : <b><%= osummarystock.FresultCount %></b>
	</td>
</tr>
<% if osummarystock.FresultCount > 0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<th width="30"></th>
	<th width="50">이미지</th>
	<th width="120">브랜드</th>
	<th width="50">상품구분</th>
	<th width="60">상품코드</th>
	<th width="50">옵션코드</th>
	<th>상품명<br>(옵션명)</th>
	<th width="70">판매전환일</th>
	<th width="70">랙코드</th>
	<th width="70">보조랙코드</th>
	<th width="35">배송<br>구분</th>
	<th width="35">전체<br>입고<br>반품</th>
	<th width="35">전체<br>판매<br>반품</th>
	<th width="35">전체<br>출고<br>반품</th>
	<th width="35">기타<br>출고<br>반품</th>

	<th width="35">총<br>실사<br>오차</th>
	<th width="35">실사<br>재고</th>
	<th width="35">총<br>불량</th>
	<th width="35">유효<br>재고</th>

	<th width="35">총<br>상품<br>준비</th>
	<th width="35">재고<br>파악<br>재고</th>
	<th width="35">ON<br>결제<br>완료</th>
	<th width="35">ON<br>주문<br>접수</th>
	<th width="35">한정<br>비교<br>재고</th>
<!--    <th width="35">적정<br>한정<br>재고</th>	-->
	<th width="40">판매<br>여부</th>
	<th width="50">한정<br>여부</th>
<!--	<th width="35">품절<br>여부</th>	-->
</tr>
</thead>
<tbody>
<% for i=0 to osummarystock.FresultCount-1 %>
<form name="frmBuyPrc_<%= i %>" method="post" onSubmit="return false;" action="/admin/shopmaster/dolimitsoldset.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="<%= osummarystock.FItemList(i).FItemID %>">
<% if osummarystock.FItemList(i).Fisusing="Y" then %>
	<tr bgcolor="#FFFFFF" align="center">
<% else %>
	<tr bgcolor="#EEEEEE" align="center">
<% end if %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
	<td align="left">
		<%= osummarystock.FItemList(i).FMakerID %>
	</td>
	<td><%= osummarystock.FItemList(i).Fitemgubun %></td>
	<td>
		<a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemID %>');"><%= osummarystock.FItemList(i).FItemID %></a>
	</td>
	<td><%= osummarystock.FItemList(i).Fitemoption %></td>
	<td align="left">
		<a href="javascript:PopItemDetail('<%= osummarystock.FItemList(i).FItemID %>','<%= osummarystock.FItemList(i).FItemOption %>')"><%= osummarystock.FItemList(i).FItemName %></a>
		<% if (osummarystock.FItemList(i).FItemOptionName <> "") then %>
			<br>(<%= osummarystock.FItemList(i).FItemOptionName %>)
		<% end if %>
	</td>
	<td><%= chkIIF(isNull(osummarystock.FItemList(i).FsellStdate),"",left(osummarystock.FItemList(i).FsellStdate,10)) %></td>
	<td><%= osummarystock.FItemList(i).FItemRackCode %></td>
	<td><%= osummarystock.FItemList(i).FItemsubrackcode %></td>
	<td><font color="<%= mwdivColor(osummarystock.FItemList(i).Fmwdiv) %>"><%= mwdivName(osummarystock.FItemList(i).Fmwdiv) %></font></td>
	<td><%= osummarystock.FItemList(i).Ftotipgono %></td>
	<td><%= -1*osummarystock.FItemList(i).Ftotsellno %></td>
	<td><%= osummarystock.FItemList(i).Foffchulgono + osummarystock.FItemList(i).Foffrechulgono %></td>
	<td><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></td>

	<td align="right"><b><%= FormatNumber(osummarystock.FItemList(i).Ferrrealcheckno, 0) %></b>&nbsp;</td>
	<td align="right"><%= FormatNumber(osummarystock.FItemList(i).getErrAssignStock, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Ferrbaditemno, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Frealstock, 0) %>&nbsp;</td>

	<td><%= osummarystock.FItemList(i).Fipkumdiv5 + osummarystock.FItemList(i).Foffconfirmno %></td>
	<td><b><%= osummarystock.FItemList(i).GetCheckStockNo %></b></td>
	<td><%= osummarystock.FItemList(i).Fipkumdiv4 %></td>
	<td><%= osummarystock.FItemList(i).Fipkumdiv2 %></td>
	<td><b><%= osummarystock.FItemList(i).GetLimitStockNo %></b></td>
<!--    <td><b><%= round(osummarystock.FItemList(i).GetLimitStockNo * 0.95,0) %></b></td>	-->
<!--    <td><b><font color="red"><%= osummarystock.FItemList(i).GetLimitStockNo - osummarystock.FItemList(i).GetLimitStr %></font></b></td>	-->
	<td>
		<%= osummarystock.FItemList(i).Fsellyn %>
		<!--
		<input type="radio" name="sellyn" value="Y" <% if osummarystock.FItemList(i).Fsellyn="Y" then response.write "checked" %>>Y
		<input type="radio" name="sellyn" value="N" <% if osummarystock.FItemList(i).Fsellyn="N" then response.write "checked ><font color=red>N</font>" else response.write ">N" %>
		-->
	</td>

	<td>
	<% if (osummarystock.FItemList(i).Flimityn = "Y") then %>
		한정(<%= osummarystock.FItemList(i).GetLimitStr %>)
		<% if (osummarystock.FItemList(i).Foptlimityn = "Y") then %>
		<br>(<%= osummarystock.FItemList(i).Foptlimitno %>/<%= osummarystock.FItemList(i).Foptlimitsold %>)
		<% else %>
		<br>(<%= osummarystock.FItemList(i).FLimitNo %>/<%= osummarystock.FItemList(i).FLimitSold %>)
		<% end if %>
	<% end if %>
	</td>
<!--    <td><% if osummarystock.FItemList(i).IsSoldOut  then %><font color="red">품절</font><% end if %></td>	-->
</tr>
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="26" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</tbody>
</table>

<form name="frmArrupdate" method="post" action="/admin/shopmaster/dolimitsoldset.asp" style="margin:0px;">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="sellyn" value="">
</form>
<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
