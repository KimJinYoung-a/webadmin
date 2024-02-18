<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/company/ch/incGlobalVariable.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->
<%

dim itemid, itemname, makerid, sellyn, usingyn, mwdiv, limityn, overSeaYn, weightYn, itemrackcode, vRegUserID, vIsReg
dim cdl, cdm, cds, sortDiv, sortDiv2, sellcash1, sellcash2
dim page

itemid		= request("itemid")
itemname	= request("itemname")
makerid		= request("makerid")
sellyn		= request("sellyn")
usingyn		= request("usingyn")
mwdiv		= request("mwdiv")
limityn		= request("limityn")
overSeaYn	= request("overSeaYn")
weightYn	= request("weightYn")
itemrackcode= request("itemrackcode")
sortDiv		= request("sortDiv")
sortDiv2	= request("sortDiv2")
vRegUserID	= request("reguserid")
vIsReg		= request("isreg")
sellcash1	= request("sellcash1")
sellcash2	= request("sellcash2")

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

page = request("page")

'기본값
if (page="") then page=1
if mwdiv="" then mwdiv="MW"
if overSeaYn="" then overSeaYn="Y"
if weightYn="" then weightYn="Y"
if sortDiv="" then sortDiv="new"
if sortDiv2="" then sortDiv2="weightup"


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


'==============================================================================
dim oitem

set oitem = new COverSeasItem

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
oitem.FRectSortDiv2		= sortDiv2

oitem.FRectRegUserID	= vRegUserID
oitem.FRectIsReg		= vIsReg
oitem.FRectSellcash1	= sellcash1
oitem.FRectSellcash2	= sellcash2

oitem.GetOverSeasTargetItemList

dim i

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function chgSort(srt){
	document.frm.sortDiv.value= srt;
	document.frm.submit();
}

function chgReg(reg){
	document.frm.isreg.value= reg;
	document.frm.submit();
}

// ============================================================================
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

// ============================================================================
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
	popwin = window.open('/admin/itemmaster/pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 판매가 및 공급가 설정
function editItemPriceInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('/admin/itemmaster/pop_ItemPriceInfo.asp?' + param ,'editItemPrice','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}


function PopItemWeightEdit(iitemid){
	var popwin = window.open('/warehouse/pop_ItemWeightEdit.asp?itembarcode=' + iitemid,'itemWeightEdit','width=500,height=300,scrollbars=yes,resizable=yes')
}

function PopItemContent(iitemid){
	var popwin = window.open('/admin/itemmaster/overseas/popItemContent.asp?countrycd=kr&itemid=' + iitemid,'itemWeightEdit','width=700,height=700,scrollbars=yes,resizable=yes')
}

function jsSearchBrandID(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/company/ch/popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function itemlistXls()
{
	document.frm.action = "target_itemlist_xls.asp";
	document.frm.submit();
	
	document.frm.action = "target_itemlist.asp";
}
</script>
</head>
<body>
<table width="700" border="0" class="a">
	<tr>
		<td>&gt;&gt;판매대상상품리스트</td>
	</tr>
</table>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<input type="hidden" name="sortDiv" value="<%=sortDiv%>">
	<input type="hidden" name="isreg" value="<%=vIsReg%>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			<br>
			상품코드 :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
			&nbsp;
			상품명 :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	&nbsp;
	     	사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;
		</td>
	</tr>
    </form>
</table>

<p>
<%
	If Request.ServerVariables("REMOTE_ADDR") = "61.252.133.15" Then
%>
<a href="javascript:itemlistXls();"><img src="http://webadmin.10x10.co.kr/images/btn_excel.gif" border="0"></a>
<br>
<%
	End If
%>

<!-- 리스트 시작 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<table width="100%" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					검색결과 : <b><%= oitem.FTotalCount%></b>
					&nbsp;
					페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
				</td>
				<td align="right">
					등록여부 :
					<select name="reg" class="select" onchange="chgReg(this.value)">
						<option value="" <%= CHKIIF(vIsReg="","selected","") %>>전체보기</option>
						<option value="x" <%= CHKIIF(vIsReg="x","selected","") %>>미등록만</option>
						<option value="o" <%= CHKIIF(vIsReg="o","selected","") %>>등록만</option>
					</select>
					&nbsp;&nbsp;&nbsp;
					정렬방법 :
					<select name="sort" class="select" onchange="chgSort(this.value)">
						<option value="new" <% if sortDiv="new" then Response.Write "selected" %>>신상품순</option>
						<option value="best" <% if sortDiv="best" then Response.Write "selected" %>>인기상품순</option>
						<option value="min" <% if sortDiv="min" then Response.Write "selected" %>>낮은가격순</option>
						<option value="hi" <% if sortDiv="hi" then Response.Write "selected" %>>높은가격순</option>
						<option value="hs" <% if sortDiv="hs" then Response.Write "selected" %>>높은할인율순</option>
						<!--<option value="weight" <% if sortDiv="weight" then Response.Write "selected" %>>상품무게순</option>//-->
					</select>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">No.</td>
		<td width=50> 이미지</td>
		<td width="100">브랜드ID</td>
		<td> 상품명</td>
		<td width="60">판매가</td>
		<td width="30">계약<br>구분</td>
		<td width="30">판매<br>여부</td>
		<td width="30">사용<br>여부</td>
		<td width="40">해외<br>여부</td>
		<td width="60">상품<br>무게</td>
		<td width="100">등록여부</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center">
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">				
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left"><% =oitem.FItemList(i).Fitemname %></td>
		<td align="right">
		<%
			'Response.Write "<a href=""javascript:editItemPriceInfo('" & oitem.FItemList(i).Fitemid & "')"" title='판매가 및 공급가 설정'>" & FormatNumber(oitem.FItemList(i).Forgprice,0) & "</a>"
			Response.Write "" & FormatNumber(oitem.FItemList(i).Forgprice,0) & ""
			'할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>(할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'쿠폰가
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						'Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
					Case "2"
						'Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
				end Select
			end if
		%>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).FdeliverOverseas,"yn") %></td>
		<td align="center"><%= FormatNumber(oitem.FItemList(i).FitemWeight,0) %>g</td>
	    <td align="center">
	    	<% If oitem.FItemList(i).FExistMultiLang = "Y" Then %>
	    		등록
	    	<% Else %>
	    		미등록
	    	<% End If %>
	    </td>
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
	
</table>
<% end if %>

<% set oitem = nothing %>

</body>
</html>
<!-- 표 하단바 끝-->
<!-- #include virtual="/lib/db/dbclose.asp" -->