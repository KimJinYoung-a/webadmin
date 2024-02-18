<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  해외상품속성관리
' History : 이상구 생성
'			2018.10.15 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, itemname, makerid, sellyn, usingyn, mwdiv, limityn, overSeaYn, weightYn, itemrackcode,research
dim cdl, cdm, cds, sortDiv, page, limitrealstock, stocktype, i, pojangok, itemManageType, sizeYn, chdeliverOverseas
dim itemdivNotexists
	itemid		= request("itemid")
	itemname	= requestCheckVar(request("itemname"),128)
	makerid		= requestCheckVar(request("makerid"),32)
	sellyn		= requestCheckVar(request("sellyn"),1)
	usingyn		= requestCheckVar(request("usingyn"),1)
	mwdiv		= requestCheckVar(request("mwdiv"),32)
	limityn		= requestCheckVar(request("limityn"),1)
	overSeaYn	= requestCheckVar(request("overSeaYn"),1)
	weightYn	= requestCheckVar(request("weightYn"),1)
	itemrackcode= requestCheckVar(request("itemrackcode"),32)
	sortDiv		= requestCheckVar(request("sortDiv"),32)
	research	=requestCheckVar(Request("research"),1)
	pojangok	=requestCheckVar(Request("pojangok"),1)
	cdl = requestCheckVar(request("cdl"),32)
	cdm = requestCheckVar(request("cdm"),32)
	cds = requestCheckVar(request("cds"),32)
	page = requestCheckVar(request("page"),32)
	limitrealstock = requestCheckVar(request("limitrealstock"),32)
	stocktype = requestCheckVar(request("stocktype"),32)
	itemManageType = requestCheckVar(request("itemManageType"),32)
	sizeYn = requestCheckVar(request("sizeYn"),32)
	chdeliverOverseas = requestCheckVar(request("chdeliverOverseas"),10)
	itemdivNotexists = requestCheckVar(request("itemdivNotexists"),32)

'기본값
if chdeliverOverseas="" then chdeliverOverseas="Y"
if (page="") then page=1
if sortDiv="" then sortDiv="new"
if research="" then
	if mwdiv="" then mwdiv="MW"
	if overSeaYn="" then overSeaYn="Y"
	if weightYn="" then weightYn="Y"
	'if pojangok="" then pojangok="Y"
	itemManageType = "I"
end if
if research="" and itemdivNotexists="" then
	itemdivNotexists="on"
end if
if (stocktype = "") then
	stocktype = "sys"
end if

if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim oitem
set oitem = new CItem
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.FRectMakerid      = makerid
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectIsOversea	= overSeaYn
	oitem.FRectIsWeight		= weightYn
	oitem.FRectRackcode		= itemrackcode
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectSortDiv		= sortDiv
	oitem.FRectlimitrealstock = limitrealstock
	oitem.FRectStockType = stocktype
	oitem.FRectpojangok = pojangok
	oitem.FRecItemManageType = itemManageType
	oitem.FRectSizeYn = sizeYn

	if itemdivNotexists="on" then
		oitem.frectitemdivNotexists="'08','21'"
	end if

	oitem.GetItemAboardList

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function chgSort(srt){
	document.frm.sortDiv.value= srt;
	document.frm.submit();
}

// 옵션수정 -교체
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//판매수정
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

// 이미지수정
function editItemImage(itemid, makerid) {
	var param = "itemid=" + itemid;

	//if(makerid =="ithinkso"){
		//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}else{
		popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}
	popwin.focus();
}

// 상품설명 이미지 등록/수정
function popItemContImage(itemid)
{
	var popwin = window.open("/admin/shopmaster/item_imgcontents_write.asp?mode=edit&itemid=" + itemid + "&menupos=423","popitemContImage","width=600 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// 재고현황 팝업
function PopItemStock(itemid){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// 기본정보 수정
function editItemBasicInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 판매가 및 공급가 설정
function editItemPriceInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('pop_ItemPriceInfo.asp?' + param ,'editItemPrice','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopItemWeightEdit(iitemid){
	var popwin = window.open('/warehouse/pop_ItemWeightEdit.asp?itembarcode=' + iitemid + '&menupos=<%=menupos%>','itemWeightEdit','width=800,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function jsSubmit(frm) {
	/*
	if ((frm.itemManageType.value == 'O') && (frm.makerid.value == '')) {
		alert('옵션별 검색은 브랜드를 지정해야만 검색가능합니다.');
		return;
	}
	*/

	frm.submit();
}

function downloadexcel() {
	alert('200000건까지 다운로드 가능. 로딩중 기다려 주세요.');
	frm.action='/common/item/itemAboard_exceldownload.asp';
	frm.target='view';
	frm.submit();
	frm.action='';
	frm.target='';
	return false;
}

function regdeliverOverseas(){
	if (frm.chdeliverOverseas.value==""){
		alert("일괄변경하실 해외배송여부를 선택해 주세요.");
		return;
	}
	frmlist.chdeliverOverseas.value=frm.chdeliverOverseas.value

    if ($('input[name="check"]:checked').length == 0) {
        alert('일괄변경하실 상품을 선택해 주세요.');
        return;
    }

	frmlist.action="/warehouse/itemWeight_process.asp";
	frmlist.target="view";
	frmlist.submit();
}

function toggleChecked(status) {
    $('[name="check"]').each(function () {
        $(this).prop("checked", status);
    });
}

$(document).ready(function () {
    var checkAllBox = $("#ckall");

    checkAllBox.click(function () {
        var status = checkAllBox.prop('checked');
        toggleChecked(status);
    });
});

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<input type="hidden" name="sortDiv" value="<%=sortDiv%>">
<input type="hidden" name="research" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드 : <%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br><br>
		렉코드 :
		<input type="text" class="text" name="itemrackcode" value="<%= itemrackcode %>" size="12" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
		&nbsp;
		상품명 :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="jsSubmit(document.frm)">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		판매 : <% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;
		사용 : <% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;
		한정 : <% drawSelectBoxLimitYN "limityn", limityn %>
		&nbsp;
		거래구분 : <% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;
		해외배송 : <% drawSelectBoxUsingYN "overSeaYn", overSeaYn %>
		&nbsp;
		포장가능여부 : <% drawSelectBoxUsingYN "pojangok", pojangok %>
		&nbsp;
		재고 <select name="stocktype" class="select">
			<option value="sys" <% if (stocktype = "sys") then %>selected<% end if %> >시스템재고</option>
			<option value="real" <% if (stocktype = "real") then %>selected<% end if %> >유효재고</option>
		</select>
		: <% drawSelectBoxexistsstock "limitrealstock", limitrealstock, "" %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		등록방식 :
		<select name="itemManageType" class="select">
			<option value="I" <% if (itemManageType = "I") then %>selected<% end if %> >상품별</option>
			<option value="O" <% if (itemManageType = "O") then %>selected<% end if %> >옵션별</option>
		</select>
		&nbsp;
		무게여부 : <% drawSelectBoxUsingYN "weightYn", weightYn %>
		&nbsp;
		사이즈여부 : <% drawSelectBoxUsingYN "sizeYn", sizeYn %>
		&nbsp;
		<input type="checkbox" name="itemdivNotexists" value="on" <% if itemdivNotexists="on" then response.write "checked" %> >티켓/클래스/딜상품제외
	</td>
</tr>
</table>

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			해외배송여부 : <% drawSelectBoxUsingYN "chdeliverOverseas", chdeliverOverseas %>
			<input type="button" onclick="regdeliverOverseas();" value="해외배송여부일괄저장" class="button">
		</td>
		<td align="right">
			<input type="button" onclick="downloadexcel(); return false;" value="엑셀 다운로드" class="button">&nbsp;&nbsp;
			정렬방법 :
			<select name="sort" class="select" onchange="chgSort(this.value)">
				<option value="new" <% if sortDiv="new" then Response.Write "selected" %>>신상품순</option>
				<option value="rack" <% if sortDiv="rack" then Response.Write "selected" %>>렉코드순</option>
				<option value="weight" <% if sortDiv="weight" then Response.Write "selected" %>>상품무게순</option>
			</select>
		</td>
	</tr>
</table>
</form>
<!-- 액션 끝 -->

<form name="frmlist" method="post" action="" style="margin:0px;">
<input type="hidden" name="mode" value="chdeliverOverseas">
<input type="hidden" name="chdeliverOverseas" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oitem.FTotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="30"><input type="checkbox" name="ckall" id="ckall"></td>
	<td width="50"> 이미지</td>
	<td width="60">Rack</td>
	<td width="60">No.</td>
	<td width="100">브랜드ID</td>
	<td>상품명</td>
	<td width="60">판매가</td>
	<td width="50">계약<br>구분</td>
	<td width="30">판매<br>여부</td>
	<td width="30">사용<br>여부</td>
	<td width="30">해외<br>여부</td>
	<td width="50">포장가능<br>여부</td>
	<td width="60">상품무게</td>
	<td width="120">상품사이즈</td>
	<td width="40">비고</td>
</tr>

<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center"><input type="checkbox" name="check" value="<%= oitem.FItemList(i).Fitemid %>" /></td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="center"><%= oitem.FItemList(i).Fitemrackcode %></td>
		<td align="center">
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left"><% =oitem.FItemList(i).Fitemname %></td>
		<td align="right">
		<%
			Response.Write "<a href=""javascript:editItemPriceInfo('" & oitem.FItemList(i).Fitemid & "')"" title='판매가 및 공급가 설정'>" & FormatNumber(oitem.FItemList(i).Forgprice,0) & "</a>"
			'할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>(할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'쿠폰가
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
				end Select
			end if
		%>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).FdeliverOverseas,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fpojangok,"yn") %></td>
		<td align="center"><%= FormatNumber(oitem.FItemList(i).FitemWeight,0) %>g</td>
		<td align="center">
			<%= oitem.FItemList(i).fvolX %> * <%= oitem.FItemList(i).fvolY %> * <%= oitem.FItemList(i).fvolZ %> cm
		</td>
	    <td align="center"><input type="button" onClick="PopItemWeightEdit('<%= oitem.FItemList(i).Fitemid %>');" value="수정" class="button"></td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
    			<% if i>oitem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oitem.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="300" frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height="0" frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
