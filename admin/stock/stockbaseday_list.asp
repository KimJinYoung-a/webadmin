<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 재고기준일수관리
' Hieditor : 2015.05.27 이상구 생성
'			 2016.10.10 한용민 수정(기계가 퍼감. 페이징 방식 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newshortagestockcls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
const C_STOCK_DAY=7

''아래 두 페이지는 검색조건을 동일하게 맞춰야 한다. XXXXXXXXXX
''/admin/stock/newshortagestock.asp
''/admin/newstorage/popjumunitemNew.asp

dim page, mode, makerid, shopid,itemid, itemname, research, onlynotmddanjong, onoffgubun, idx, changemakerid
dim onlynotupchebeasong, onlyusingitem, onlyusingitemoption, onlynotdanjong, soldoutover7days, onlysell
dim onlynottempdanjong, onlynotinputday, purchasetype, mwdiv, centermwdiv, iPageSize, iTotCnt, iTotalPage
dim i, shopsuplycash, buycash, IsAvailDelete, DayForSellCount
	shopid = request("shopid")
	page = request("page")
	mode = request("mode")
	itemid = request("itemid")
	itemname = Trim(request("itemname"))
	research = request("research")
	onlynotupchebeasong = request("onlynotupchebeasong")
	onlyusingitem = request("onlyusingitem")
	onlyusingitemoption = request("onlyusingitemoption")
	onlynotdanjong = request("onlynotdanjong")
	soldoutover7days = request("soldoutover7days")
	onoffgubun = request("onoffgubun")
	idx = request("idx")
	onlysell = request("onlysell")
	onlynottempdanjong = request("onlynottempdanjong")
	onlynotmddanjong = request("onlynotmddanjong")
	onlynotinputday = request("onlynotinputday")
	purchasetype = request("purchasetype")
	mwdiv = request("mwdiv")
	centermwdiv = request("centermwdiv")
	DayForSellCount = requestcheckvar(request("DayForSellCount"),10)

	changemakerid = "Y"
	if (changemakerid = "") then
		changemakerid = request("changemakerid")
	end if
	
	makerid = request("makerid")
	if (makerid = "") then
		makerid = request("suplyer")
	end if

iPageSize = 50

if (research<>"on") and (onlynotupchebeasong = "") then
	onlynotupchebeasong = "on"
end if

if (research<>"on") and (onlyusingitem = "") then
	onlyusingitem = "on"
end if

if (research<>"on") and (onlyusingitemoption="") then
	onlyusingitemoption = "on"
end if

if (research<>"on") and (onlynotdanjong = "") then
	onlynotdanjong = "on"
end if

if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if

if page="" then page=1
if mode="" then mode="bybrand"

'상품코드 유효성 검사(2008.07.31;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim ostockbaseday
set ostockbaseday  = new CShortageStock
ostockbaseday.FPageSize = iPageSize
ostockbaseday.FCurrPage = page
ostockbaseday.FRectOnlySell			= onlysell
ostockbaseday.FRectOnlyUsingItem		= onlyusingitem
ostockbaseday.FRectOnlyUsingItemOption	= onlyusingitemoption
ostockbaseday.FRectOnlyNotUpcheBeasong	= onlynotupchebeasong
ostockbaseday.FRectOnlyNotDanjong		= onlynotdanjong
ostockbaseday.FRectOnlyNotTempDanjong	= onlynottempdanjong
ostockbaseday.FRectOnlyNotMDDanjong	= onlynotmddanjong
ostockbaseday.FRectOnlyNotInputDay	= onlynotinputday
ostockbaseday.FRectPurchaseType		= purchasetype
ostockbaseday.FRectMakerid				= makerid
ostockbaseday.FRectItemId				= itemid
ostockbaseday.FRectItemName			= html2db(itemname)
ostockbaseday.FRectMWDiv = mwdiv
ostockbaseday.FRectCenterMWDiv = centermwdiv
ostockbaseday.FRectDayForSellCount = DayForSellCount

if onoffgubun = "offline" or onoffgubun = "etcitem" then
	if (onoffgubun = "offline") then
		ostockbaseday.FRectItemGubun = "90"
	else
		ostockbaseday.FRectItemGubunExclude = "90"
	end if

	ostockbaseday.GetStockBaseDayItemListOffline
else
	ostockbaseday.GetStockBaseDayItemListOnline
end if

iTotCnt = ostockbaseday.FResultCount
iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate, iStartDate, iEndDate

'재입고예정일
'오늘기준 +- 일주일은 검정색 표시 / 그 이외 회색표시
if (yyyy1="") then
    nowdate = Left(CStr(DateAdd("d",now(),-7)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

    nowdate = Left(CStr(DateAdd("d",now(),+7)),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

iStartDate  = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
iEndDate    = Left(CStr(DateSerial(yyyy2,mm2,dd2)),10)

%>

<script type="text/javascript">
	
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ChangeReqDay(frm){
	if (!(IsDigit(frm.maxsellday.value))){
		alert('숫자만 가능합니다.');
		return;
	}

	if (confirm('필요 재고 기준일을 변경하시겠습니까?')){
		frm.submit();
	}
}

function Research(page){
	document.frm.page.value= page;
	document.frm.submit();
}

function DeleteStockLog(itemgubun,itemid,itemoption){
    if (confirm('삭제 하시겠습니까?')){
        frmdelstock.target="_blank";
        frmdelstock.itemgubun.value = itemgubun;
        frmdelstock.itemid.value = itemid;
        frmdelstock.itemoption.value = itemoption;
        frmdelstock.submit();
    }
}

function search(frm){
	/*
	if ((frm.makerid.value.length<1)){
		if ((frm.mode[0].checked)&&(frm.designer.value.length<1)){
			alert('브랜드를 선택 하세요.');
			frm.designer.focus();
			return;
		}
	}
	*/
	document.frm.page.value = 1;
	frm.submit();
}

function CheckAll(v) {
	for (var i = 0;; i++) {
		var chk = document.getElementById("chk_" + i);
		if (chk == undefined) {
			break;
		}
		chk.checked = v.checked;
		checkhL(chk);
	}
}

function ApplyDefaultValueToAll() {
	var dayforsellcountall = document.getElementById("dayforsellcountall");
	var dayforsafestockall = document.getElementById("dayforsafestockall");
	var dayforleadtimeall = document.getElementById("dayforleadtimeall");
	var dayformaxstockall = document.getElementById("dayformaxstockall");

	for (var i = 0;; i++) {
		var chk = document.getElementById("chk_" + i);

		var dayforsellcount = document.getElementById("dayforsellcount_" + i);
		var dayforsafestock = document.getElementById("dayforsafestock_" + i);
		var dayforleadtime = document.getElementById("dayforleadtime_" + i);
		var dayformaxstock = document.getElementById("dayformaxstock_" + i);

		if (chk == undefined) {
			break;
		}

		if (chk.checked) {
			dayforsellcount.value = dayforsellcountall.value;
			dayforsafestock.value = dayforsafestockall.value;
			dayforleadtime.value = dayforleadtimeall.value;
			dayformaxstock.value = dayformaxstockall.value;
		}
	}
}

function SaveSelectedItems() {
	var result = "";
	var dayforsellcount = "";
	var dayforsafestock = "";
	var dayforleadtime = "";
	var dayformaxstock = "";

	for (var i = 0;; i++) {
		var chk = document.getElementById("chk_" + i);

		var oitemgubun = document.getElementById("itemgubun_" + i);
		var oitemid = document.getElementById("itemid_" + i);
		var oitemoption = document.getElementById("itemoption_" + i);

		var odayforsellcount = document.getElementById("dayforsellcount_" + i);
		var odayforsafestock = document.getElementById("dayforsafestock_" + i);
		var odayforleadtime = document.getElementById("dayforleadtime_" + i);
		var odayformaxstock = document.getElementById("dayformaxstock_" + i);

		if (chk == undefined) {
			break;
		}

		if (chk.checked) {
			dayforsellcount = removeComma(trim(odayforsellcount.value));
			dayforsafestock = removeComma(trim(odayforsafestock.value));
			dayforleadtime = removeComma(trim(odayforleadtime.value));
			dayformaxstock = removeComma(trim(odayformaxstock.value));

			if (isNumeric(dayforsellcount) != true) {
				alert("숫자만 가능합니다.");
				odayforsellcount.focus();
				return;
			}

			if (isNumeric(dayforsafestock) != true) {
				alert("숫자만 가능합니다.");
				odayforsafestock.focus();
				return;
			}

			if (isNumeric(dayforleadtime) != true) {
				alert("숫자만 가능합니다.");
				odayforleadtime.focus();
				return;
			}

			if (isNumeric(dayformaxstock) != true) {
				alert("숫자만 가능합니다.");
				odayformaxstock.focus();
				return;
			}

			result = result + "|" + oitemgubun.value + "," + oitemid.value + "," + oitemoption.value + "," + dayforsellcount + "," + dayforsafestock + "," + dayforleadtime + "," + dayformaxstock;
		}
	}

	if (result == "") {
		alert("선택된 상품이 없습니다.");
		return;
	}

	if (confirm("저장하시겠습니까?") == true) {
		var frm = document.frmAct;

		frm.mode.value = "saveestockbaseday";
		frm.itemArr.value = result;
		frm.submit();
	}
}

function trim(value) {
	return value.replace(/^\s+|\s+$/g,"");
}

function removeComma(value) {
	return value.replace(/,/g,"");
}


function isNumeric(value) {
	var v = trim(value);

	var regx = /^\d{1,10}$/;

	return regx.test(v);
}

function SetCheck(v) {
	var chk = document.getElementById(v);
	chk.checked = true;
	checkhL(chk);
}

function checkhL(e){
    if (e.checked){
        hL(e);
    }else{
        dL(e);
    }
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="page" value="1">
<% if (changemakerid <> "Y") then %>
<input type="hidden" name="makerid" value="<%= makerid %>" >
<% else %>
<input type="hidden" name="changemakerid" value="Y" >
<% end if %>
<input type="hidden" name="shopid" value="<%= shopid %>" >
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<% if (changemakerid <> "Y") then %>
		브랜드 : <b><%= makerid %></b>
		<% else %>
		브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		<% end if %>
		&nbsp;
		|
		&nbsp;
		구분 : <input type="radio" name="onoffgubun" value="online" <% if onoffgubun="online" then response.write "checked" %> >온라인
		<input type="radio" name="onoffgubun" value="offline" <% if onoffgubun="offline" then response.write "checked" %> >오프라인
		&nbsp;
		<input type="radio" name="onoffgubun" value="etcitem" <% if onoffgubun="etcitem" then response.write "checked" %> >기타(사은품 등)
		&nbsp;
		|
		&nbsp;
		<input type=checkbox name="onlyusingitem" <% if onlyusingitem = "on" then response.write "checked" %> >사용상품만
		<input type=checkbox name="onlyusingitemoption" <% if onlyusingitemoption = "on" then response.write "checked" %> >사용옵션만
		<input type=checkbox name="onlysell" <% if onlysell = "on" then response.write "checked" %> >판매상품만
		<input type=checkbox name="onlynotupchebeasong" <% if onlynotupchebeasong = "on" then response.write "checked" %> >업체배송제외
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:search(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type=checkbox name="onlynotdanjong" <% if onlynotdanjong = "on" then response.write "checked" %> >단종제외
		<input type=checkbox name="onlynottempdanjong" <% if onlynottempdanjong = "on" then response.write "checked" %> >일시품절제외
		<input type=checkbox name="onlynotmddanjong" <% if onlynotmddanjong = "on" then response.write "checked" %> >MD단종제외
		<input type=checkbox name="onlynotinputday" <% if onlynotinputday = "on" then response.write "checked" %> >기준일수 미입력상품만
		&nbsp;
		판매량산정기준일수 : <input type="text" class="text" name="DayForSellCount" value="<%= DayForSellCount %>" size=5 maxlength=5>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=8 maxlength=7>
		&nbsp;
		|
		&nbsp;
		상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size=20 maxlength=50>
		&nbsp;
		|
		&nbsp;
		구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		&nbsp;
		|
		&nbsp;
		* 거래구분(ON) :<% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;
		|
		&nbsp;
		* 센터매입구분(OFF) :
		<select class="select" name="centermwdiv">
			<option value="">전체</option>
			<option value="M" <% if centermwdiv="M" then response.write "selected" %> >매입</option>
			<option value="W" <% if centermwdiv="W" then response.write "selected" %> >특정</option>
			<option value="N" <% if centermwdiv="N" then response.write "selected" %> >미지정</option>
		</select>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frmAct" method=post action="stockbaseday_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemArr" value="">
<tr>
	<td align="left">
		* 입력이 안되는 상품은 로직스에서 바코드 등록 후 입력하실 수 있습니다.
	</td>
	<td align="right">
		<input type="button" class="button" value="선택상품 일괄설정" onClick="ApplyDefaultValueToAll();">
		<input type="button" class="button" value="선택상품 저장" onClick="SaveSelectedItems();">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%=iTotCnt%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=iTotalPage%></b> &nbsp;(최대검색건수 : <%= ostockbaseday.FTotalCount %>)
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"></td>
	<td>브랜드ID</td>
	<td width="50">이미지</td>
	<td width="30">구분</td>
	<td width="60">상품<br>코드</td>
	<td width="50">옵션</td>
	<td>상품명<font color="blue">[옵션명]</font></td>

	<td width="50">매입구분</td>

	<td width="60">판매량<br>산정<br>기준일수</td>
	<td width="60">안전재고<br>일수</td>
	<td width="60">리드타임</td>
	<td width="60">최대재고<br>일수</td>

	<td width="70">비고</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="chkall" onClick="CheckAll(this)"></td>
	<td height="30" colspan="7"></td>

    <% if makerid="ithinkso" then %>
    	<td><input type="text" class="text" id="dayforsellcountall" name="dayforsellcountall" size="2" value="30"></td>
    	<td><input type="text" class="text" id="dayforsafestockall" name="dayforsafestockall" size="2" value="7"></td>
    	<td><input type="text" class="text" id="dayforleadtimeall" name="dayforleadtimeall" size="2" value="30"></td>
    	<td><input type="text" class="text" id="dayformaxstockall" name="dayformaxstockall" size="2" value="30"></td>
    <% else %>
    	<td><input type="text" class="text" id="dayforsellcountall" name="dayforsellcountall" size="2" value="7"></td>
    	<td><input type="text" class="text" id="dayforsafestockall" name="dayforsafestockall" size="2" value="3"></td>
    	<td><input type="text" class="text" id="dayforleadtimeall" name="dayforleadtimeall" size="2" value="2"></td>
    	<td><input type="text" class="text" id="dayformaxstockall" name="dayformaxstockall" size="2" value="10"></td>
    <% end if %>

	<td></td>
</tr>

<% if ostockbaseday.FTotalCount>0 then %>
	<% For i = 0 To ostockbaseday.FTotalCount -1 %>
	<% if ostockbaseday.FItemList(i).IsInvalidOption then %>
	<tr align="center" bgcolor="#CCCCCC">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><input type="checkbox" id="chk_<%= i %>" name="chk_<%= i %>"></td>
		<input type="hidden" id="itemgubun_<%= i %>" name="itemgubun_<%= i %>" value="<%= ostockbaseday.FItemList(i).Fitemgubun %>">
		<input type="hidden" id="itemid_<%= i %>" name="itemid_<%= i %>" value="<%= ostockbaseday.FItemList(i).FItemID %>">
		<input type="hidden" id="itemoption_<%= i %>" name="itemoption_<%= i %>" value="<%= ostockbaseday.FItemList(i).FItemOption %>">
	
	
		<td><%= ostockbaseday.FItemList(i).FMakerID %></td>
		<td width="50" align=center>
			<% if (onoffgubun <> "offline") and (onoffgubun <> "etcitem") then %>
				<img src="<%= ostockbaseday.FItemList(i).FimageSmall %>" width=50 height=50>
			<% end if %>
		</td>
		<td><%= ostockbaseday.FItemList(i).Fitemgubun %></td>
		<td><a href="javascript:PopItemSellEdit('<%= ostockbaseday.FItemList(i).FItemID %>');"><%= ostockbaseday.FItemList(i).FItemID %></a></td>
		<td><%= ostockbaseday.FItemList(i).FItemOption %></td>
	
		<td align="left">
			<a href="/admin/stock/itemcurrentstock.asp?itemid=<%= ostockbaseday.FItemList(i).FItemID %>&itemoption=<%= ostockbaseday.FItemList(i).FItemOption %>" target=_blank ><%= ostockbaseday.FItemList(i).FItemName %></a>
			<% if ostockbaseday.FItemList(i).FItemOption <> "0000" then %>
				<% if ostockbaseday.FItemList(i).Foptionusing="Y" then %>
					<br><font color="blue">[<%= ostockbaseday.FItemList(i).FItemOptionName %>]</font>
				<% else %>
					<br><font color="#AAAAAA">[<%= ostockbaseday.FItemList(i).FItemOptionName %>]</font>
				<% end if %>
			<% end if %>
		</td>
	
		<td>
			<font color="<%= ostockbaseday.FItemList(i).getMwDivColor %>"><%= ostockbaseday.FItemList(i).getMwDivName %></font><br>
		</td>
	
		<td>
			<input type="text" class="text" id="dayforsellcount_<%= i %>" name="dayforsellcount_<%= i %>" size="2" value="<%= ostockbaseday.FItemList(i).FDayForSellCount %>" onChange="SetCheck('chk_<%= i %>')" >
		</td>
		<td>
			<input type="text" class="text" id="dayforsafestock_<%= i %>" name="dayforsafestock_<%= i %>" size="2" value="<%= ostockbaseday.FItemList(i).FDayForSafeStock %>" onChange="SetCheck('chk_<%= i %>')" >
		</td>
		<td>
			<input type="text" class="text" id="dayforleadtime_<%= i %>" name="dayforleadtime_<%= i %>" size="2" value="<%= ostockbaseday.FItemList(i).FDayForLeadTime %>" onChange="SetCheck('chk_<%= i %>')" >
		</td>
		<td>
			<input type="text" class="text" id="dayformaxstock_<%= i %>" name="dayformaxstock_<%= i %>" size="2" value="<%= ostockbaseday.FItemList(i).FDayForMaxStock %>" onChange="SetCheck('chk_<%= i %>')" >
		</td>
	
		<td></td>
	</tr>
	<% next %>

	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="25">
			<%sbDisplayPaging "page", page, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">검색결과가 없습니다.</td>
	</tr>
<% end if %>

</table>

<form name="frmdelstock" method="post" action="dostockbasedayrefresh.asp">
<input type="hidden" name="mode" value="dellog">
<input type="hidden" name="itemgubun">
<input type="hidden" name="itemid">
<input type="hidden" name="itemoption">
</form>
<%
set ostockbaseday = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
