<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<% Session.codepage="65001" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 상품매입정보
' History : 이상구 생성
'			2022.10.13 한용민 수정(오류수정, 보안취약부분수정, 쿼리튜닝)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008_UTF8.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
Dim itemgubun, itemid, itemoption, page, i, j, k, oitem
	itemgubun	= requestCheckvar(trim(request("itemgubun")),32)
	itemid	= requestCheckvar(getNumeric(trim(request("itemid"))),10)
	itemoption	= requestCheckvar(trim(request("itemoption")),32)

if (itemgubun <> "10") then
	response.write "잘못된 접근입니다."
	dbget.close() : response.end
end if

page = 1

set oitem = new CItem
oitem.FRectItemGubun    = itemgubun
oitem.FRectItemid       = itemid
oitem.FRectItemOption   = itemoption
oitem.FPageSize			= 1000
oitem.FCurrPage			= page
oitem.GetBuyItemListOn

%>
<script type='text/javascript'>

function jsCheckNSave() {
	var frm = document.frm;

	for (var i = 0; i < frm.itemgubun.length; i++) {
		if (frm.itemgubun[i].value == '') { continue; }

		if (frm.makerid[i].value == '') {
			alert('브랜드를 입력하세요.');
			frm.makerid[i].focus();
			return;
		}

		if (frm.makerid[i].value.indexOf(',') >= 0) {
			alert('쉼표를 입력할 수 없습니다.');
			frm.makerid[i].focus();
			return;
		}

		/*
		if (frm.upchemanagecode[i].value == '') {
			alert('업체코드를 입력하세요.');
			frm.upchemanagecode[i].focus();
			return;
		}
		*/

		if (frm.upchemanagecode[i].value.indexOf(',') >= 0) {
			alert('쉼표를 입력할 수 없습니다.');
			frm.upchemanagecode[i].focus();
			return;
		}

		if (frm.buyitemname[i].value == '') {
			alert('매입처 상품명을 입력하세요.');
			frm.buyitemname[i].focus();
			return;
		}

		if (frm.buyitemname[i].value.indexOf(',') >= 0) {
			alert('쉼표를 입력할 수 없습니다.');
			frm.buyitemname[i].focus();
			return;
		}

		if ((frm.itemoption[i].value != '0000') && (frm.buyitemoptionname[i].value == '')) {
			alert('매입처 옵션명을 입력하세요.');
			frm.buyitemoptionname[i].focus();
			return;
		}

		if (frm.itemoption[i].value.indexOf(',') >= 0) {
			alert('쉼표를 입력할 수 없습니다.');
			frm.itemoption[i].focus();
			return;
		}

		if (frm.currencyUnit[i].value == '') {
			alert('매입처 통화화페를 지정하세요.');
			frm.currencyUnit[i].focus();
			return;
		}

		if (frm.buyitemprice[i].value == '') {
			alert('매입처 매입가를 입력하세요.');
			frm.buyitemprice[i].focus();
			return;
		}

		if (frm.buyitemprice[i].value*0 != 0) {
			alert('매입처 매입가는 숫자만 가능합니다.');
			frm.buyitemprice[i].focus();
			return;
		}
	}

	if (confirm('저장하시겠습니까?')) {
		frm.submit();
	}
}

function jsInsFirst() {
	var frm = document.frm;
	var makerid, buyitemname, currencyUnit, buyitemprice;

	if (confirm('첫번째 상품정보로 나머지 상품정보를 모두 입력합니다.\n(업체코드,옵션명제외)\n\n진행하시겠습니까?') != true) { return; }

	for (var i = 0; i < frm.itemgubun.length; i++) {
		if (frm.itemgubun[i].value == '') { continue; }

		if (i == 1) {
			makerid = frm.makerid[i].value;
			buyitemname = frm.buyitemname[i].value;
			currencyUnit = frm.currencyUnit[i].value;
			buyitemprice = frm.buyitemprice[i].value;
		} else if (i > 1) {
			frm.makerid[i].value = makerid;
			frm.buyitemname[i].value = buyitemname;
			frm.currencyUnit[i].value = currencyUnit;
			frm.buyitemprice[i].value = buyitemprice;
		}
	}
}

function jsInsEng() {
	var frm = document.frm;
	var engItemname, engItemoptionname, engMakerid;

	if (confirm('영문 상품정보가 입력되어 있는 경우 자동으로 내역을 가져옵니다.\n\n진행하시겠습니까?') != true) { return; }

	for (var i = 0; i < frm.itemgubun.length; i++) {
		if (frm.itemgubun[i].value == '') { continue; }

		if (i >= 1) {
			frm.makerid[i].value = frm.engMakerid[i].value;
			frm.buyitemname[i].value = frm.engItemname[i].value;
			frm.buyitemoptionname[i].value = frm.engItemoptionname[i].value;
		}
	}
}
</script>
<form name="frm" action="pop_BuyItemEdit_process.asp" method="post" style="margin: 0px;">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<input type="hidden" name="mode" value="ins">
	<input type="hidden" name="itemgubun" value="">
	<input type="hidden" name="itemid" value="">
	<input type="hidden" name="itemoption" value="">
	<input type="hidden" name="makerid" value="">
	<input type="hidden" name="upchemanagecode" value="">
	<input type="hidden" name="buyitemname" value="">
	<input type="hidden" name="buyitemoptionname" value="">
	<input type="hidden" name="currencyUnit" value="">
	<input type="hidden" name="buyitemprice" value="">
	<input type="hidden" name="engItemname" value="">
	<input type="hidden" name="engItemoptionname" value="">
	<input type="hidden" name="engMakerid" value="">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="10">
			검색결과 : <b><%= oitem.FResultCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30" rowspan="2">구분</td>
		<td width="80" rowspan="2"> 상품코드</td>
		<td width="50" rowspan="2">옵션</td>
		<td rowspan="2">판매<br />상품명</td>
		<td colspan="6">
			매입처정보
			<input type="button" class="button" value=" 영문명입력 " style="width: 100px;" onClick="jsInsEng();">
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			&nbsp;
			<input type="button" class="button" value=" [첫 상품정보]로 일괄입력 " style="width: 200px;" onClick="jsInsFirst();">
		</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">브랜드</td>
		<td width="100">업체코드</td>
		<td width="200">상품명</td>
		<td width="200">옵션명</td>
		<td width="90">통화화페</td>
		<td width="200">매입가</td>
	</tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="16" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% else %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<input type="hidden" name="itemgubun" value="<%= oitem.FItemList(i).FItemGubun %>">
		<input type="hidden" name="itemid" value="<%= oitem.FItemList(i).FItemId %>">
		<input type="hidden" name="itemoption" value="<%= oitem.FItemList(i).FItemOption %>">
		<td align="center"><%= oitem.FItemList(i).FItemGubun %></td>
		<td align="center"><%= oitem.FItemList(i).FItemId %></td>
		<td align="center"><%= oitem.FItemList(i).FItemOption %></td>
		<td>
			<%= oitem.FItemList(i).FItemName %>
			<% if oitem.FItemList(i).FItemOption <> "0000" then %>
			<br />[<font color="blue"><%= oitem.FItemList(i).Foptionname %></font>]
			<% end if %>
		</td>
		<td align="center"><input type="text" class="text" name="makerid" value="<%= oitem.FItemList(i).Fmakerid %>" size="15"></td>
		<td align="center"><input type="text" class="text" name="upchemanagecode" value="<%= oitem.FItemList(i).Fupchemanagecode %>" size="12"></td>
		<td align="center"><input type="text" class="text" name="buyitemname" value="<%= oitem.FItemList(i).FBuyItemName %>"></td>
		<td align="center"><input type="text" class="text" name="buyitemoptionname" value="<%= oitem.FItemList(i).FBuyItemOptionname %>"></td>
		<td align="center">
			<% DrawexchangeRate "currencyUnit",oitem.FItemList(i).FcurrencyUnit,"" %>
		</td>
		<td align="center"><input type="text" class="text" name="buyitemprice" value="<%= oitem.FItemList(i).Fbuyitemprice %>" size="12"></td>
		<input type="hidden" name="engItemname" value="<%= oitem.FItemList(i).FengItemname %>">
		<input type="hidden" name="engItemoptionname" value="<%= oitem.FItemList(i).FengItemoptionname %>">
		<input type="hidden" name="engMakerid" value="<%= oitem.FItemList(i).FengMakerid %>">
	</tr>
	<% next %>
<% end if %>
</table>
</form>

<br />

<div style="width: 100%; text-align: center;"><input type="button" class="button" value=" 저 장 " style="width: 200px;" onClick="jsCheckNSave();">

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<% Session.codepage="949" %>
